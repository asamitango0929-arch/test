"""Line charts in the user's own layout (26W必要数 sheets 1-7) per pay period, shift counts, and rosters."""
import sys, pickle, datetime as dt
from collections import Counter
from datetime import timedelta as td
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from ref import FLIGHTS, tm, m2c, X0
from periods import FL
import v3
from v3 import DEP, ARR, PSPT, codes, day_pat, days
from export import segments, order_flights
from v3 import ARR_N
def counts(k):
    f = FLIGHTS[k]; cr = f['cr']; n = DEP[cr]['ctr']; st = v3.CTR_STAY.get(cr, 0); ps = PSPT.get(k, 0)
    return dict(dc=1, gi=1, ctr=n, stay=st, pspt=ps, gate=1 + n - st + ps, arv=ARR_N.get(k, ARR[cr]['n']))
from model import INTERVAL, MAXCONSEC, MINWORK, CAP

OUT = sys.argv[1]
SOL = pickle.load(open(sys.argv[2], 'rb'))
SCEN = sys.argv[3]
SCEN_TXT = {'A': 'A案：全便自社（VZ はオーロラ委託のため対象外）', 'B': 'B案：HB・SL を委託（自社は D7・7C・RF・5J。VZ は対象外）'}[SCEN]
WJ = ['', '月', '火', '水', '木', '金', '土', '日']
NEWF = {k for k, f in FLIGHTS.items() if f.get('new')}
# colours of the user's line charts (26W必要数 sheets 1-7)
UCOL = {'SL': 'FFC00000', 'D7': 'FF0070C0', '7C': 'FFFFC000', 'RF': 'FF92D050', 'HB': 'FFFF99CC', '5J': 'FFFFFF00'}
DARK = {'FFC00000', 'FF0070C0'}
hc = lambda m: f"{m//60:02d}:{m%60:02d}"
h4 = lambda m: f"{m//60:02d}{m%60:02d}"
md = lambda d: f"{d.month}/{d.day}"

# global pattern ids (same numbering as the pattern page)
pats = {}
for d in days: pats.setdefault(day_pat[d], []).append(d)
PID = {F: f"P{i+1:02d}" for i, (F, _) in enumerate(sorted(pats.items(), key=lambda x: x[1][0]))}

# ---- styles copied from the user's sheet 1 ----
th = Side(style='thin', color='FF000000'); tk = Side(style='thick', color='FF000000'); md_ = Side(style='medium', color='FF000000')
hourln = Side(style='thin', color='FF808080')
A6 = Font(name='Arial', size=6); A6B = Font(name='Arial', size=6, bold=True); A6W = Font(name='Arial', size=6, bold=True, color='FFFFFFFF')
A7B = Font(name='Arial', size=7, bold=True); A8 = Font(name='Arial', size=8); A9B = Font(name='Arial', size=9, bold=True)
A10B = Font(name='Arial', size=10, bold=True); A12 = Font(name='Arial', size=12); G9B = Font(name='MS PGothic', size=9, bold=True)
T14 = Font(name='Arial', size=14, bold=True); M9 = Font(name='Meiryo UI', size=9); M9B = Font(name='Meiryo UI', size=9, bold=True)
CEN = Alignment(horizontal='center', vertical='center', wrap_text=True); LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
fill = lambda rgb: PatternFill('solid', fgColor=rgb)
NAOSF = fill('FFE4DFEC'); CALC = fill('FFD0E0E3'); HDRF = fill('FFD9E1F2'); NEWFILL = fill('FFFFF2CC'); GRAY = fill('FFF2F2F2'); OFF = fill('FFE2EFDA'); BAD = fill('FFFFC7CE')
BOX = Border(left=th, right=th, top=th, bottom=th)
GRID_END = m2c(24*60) - 1
tval = lambda m: dt.time(m // 60 % 24, m % 60)

def setw(ws):
    for col, w in zip('ABCDEFGHIJKLMNOPQ', [2.3, 4.9, 9.9, 6, 4.4, 6.7, 7.9, 13, 8, 8.6, 8, 8, 8, 8, 0.7, 0.7, 0.7]): ws.column_dimensions[col].width = w
    for c in range(X0, GRID_END + 2): ws.column_dimensions[CL(c)].width = 1.15

def hour_header(ws, r):
    for h in range(4, 24):
        c0 = m2c(h*60) - 3
        x = ws.cell(r, c0, h); x.font = A8; x.alignment = CEN
        ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0 + 5)
    for c in range(X0 - 3, GRID_END + 1): ws.cell(r, c).border = Border(bottom=tk)

def bar(ws, r, s, e, text, rgb=None, bold=False):
    if e <= s: return
    c1, c2 = m2c(s), max(m2c(e) - 1, m2c(s))
    x = ws.cell(r, c1, text)
    x.font = (A6W if rgb in DARK else (A6B if bold else A6)); x.alignment = CEN
    for cc in range(c1, c2 + 1):
        y = ws.cell(r, cc); y.border = BOX
        if rgb: y.fill = fill(rgb)
    if c2 > c1: ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)

def label(kind, lab):
    """translate to the vocabulary of the user's charts"""
    if kind == 'prep': return 'PRE'
    if kind == 'pw': return 'PW'
    if kind == 'rest': return 'R'
    if kind == 'gate': return lab.replace(' GATE', ' GA') if 'PSPT' not in lab else lab
    if lab.endswith(' CTR'): return lab[:-4] + ' CKIN'
    return lab

def fill_ends(p, segs):
    segs = sorted(segs)
    if segs and segs[0][0] > p['start']: segs.insert(0, [p['start'], segs[0][0], 'FREE ASGN', 'free', None])
    if segs and segs[-1][1] < p['end']: segs.append([segs[-1][1], p['end'], 'FREE ASGN', 'free', None])
    return segs

def shift_segments(p):
    base = dict(p, start=p['ostart'], end=p['oend'], rest=p.get('orest'))
    segs = segments(base)
    if p.get('orest') is None and p.get('rest'):
        segs = [s for s in segs if s[3] != 'rest'] + [[p['rest'][0], p['rest'][1], 'REST', 'rest', None]]
    if p.get('pad'):
        a, b = p['pad']
        if p['start'] < p['ostart'] and segs and segs[0][0] < p['ostart']: pass
        parts = [(a, b)]
        if p.get('rest') and a <= p['rest'][0] < b:
            ra, rb = p['rest']; parts = [(a, ra), (rb, b)]
        for x0, x1 in parts:
            if x1 > x0: segs.append([x0, x1, 'FREE ASGN', 'free', None])
    # 15-min rounding slack at the ends goes to PRE / PW
    segs.sort()
    return segs

def line_sheet(wb, name, title, sub, F, shifts, ds, naos):
    ws = wb.create_sheet(name); setw(ws); ws.sheet_view.zoomScale = 85
    ws['C1'] = title; ws['C1'].font = T14
    ws['C2'] = sub; ws['C2'].font = A8; ws['C2'].alignment = Alignment(vertical='top', wrap_text=True)
    ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=X0 + 12*12); ws.row_dimensions[2].height = 24
    H = 4
    for col, v in zip('CDEFGHIJKLMN', ['Terminal', '事業部', 'Carrier', 'Type', 'Wide\nNarrow', 'FLT', 'Frequency', 'Dest', 'STA', 'STD', 'CTR OPEN', 'GATE']):
        x = ws[f'{col}{H}']; x.value = v; x.font = A7B; x.alignment = CEN; x.border = Border(top=tk, bottom=tk)
    hour_header(ws, H); ws.row_dimensions[H].height = 18.75
    FD = dict(FL); r = H + 1
    for k in order_flights(F):
        f = FLIGHTS[k]; cr = f['cr']; arr, dep = codes(k); n = counts(k)
        dsk = [x for _, _, x in FD[k]]
        freq = 'Daily' if [1,2,3,4,5,6,7] in dsk else ' '.join(map(str, sorted({x for y in dsk for x in y})))
        std = tm(f['std']); o = std + DEP[cr]['win'][0]; c_ = std + DEP[cr]['win'][1]; g = std - v3.GATE_BEFORE
        vals = ['T1', 'CTS', cr, f['ac'], f['wn'], f"{arr}/{dep}", freq, f['dest'], tval(tm(f['sta'])), tval(std), tval(o), f"{n['gate']}名"]
        for col, v in zip('CDEFGHIJKLMN', vals):
            x = ws[f'{col}{r}']; x.value = v; x.font = A9B if col in 'HKLM' else Font(name='Arial', size=9); x.alignment = CEN; x.border = BOX
            if col in 'KLM': x.number_format = 'h:mm'
            if k in NEWF: x.fill = NEWFILL
        c0 = m2c(o) - 6
        x = ws.cell(r, c0, tval(tm(f['sta']))); x.font = A6; x.alignment = CEN; x.number_format = 'h:mm'
        ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0 + 5)
        bar(ws, r, o, g, f"{arr}\\{dep}\n(CTR OPEN {hc(o)})", UCOL[cr], True)
        bar(ws, r, g, c_, f"GATE\n{n['gate']}名", UCOL[cr], True)
        x = ws.cell(r, m2c(c_), tval(std)); x.font = A6; x.alignment = CEN; x.number_format = 'h:mm'
        ws.merge_cells(start_row=r, start_column=m2c(c_), end_row=r, end_column=m2c(c_) + 5)
        ws.row_dimensions[r].height = 18; r += 1
    r += 1; SH = r
    for col, v in zip('CDFKLMN', ['Shift Code', 'Duty', '名前', '開始', '終了', '休憩時間', '労働時間']):
        x = ws[f'{col}{SH}']; x.value = v; x.font = G9B if col in 'KLMN' else Font(name='Arial', size=9, bold=True); x.alignment = CEN
        x.border = Border(left=th, right=th, top=tk, bottom=th)
        if col in 'KLM': x.fill = HDRF
        if col == 'N': x.fill = CALC
    ws.merge_cells(f'D{SH}:E{SH}'); ws.merge_cells(f'F{SH}:J{SH}')
    hour_header(ws, SH); ws.row_dimensions[SH].height = 17.25
    r = SH + 1; first = r
    for i, p in enumerate(shifts):
        ws.cell(r, 2, i + 1).font = A12; ws.cell(r, 2).alignment = CEN
        x = ws.cell(r, 3, p['code']); x.font = A9B; x.alignment = CEN; x.border = BOX
        duty = '/'.join(dict.fromkeys(f"{j['cr']}" for j in p['jobs']))
        x = ws.cell(r, 4, duty); x.font = Font(name='Arial', size=8); x.alignment = CEN; x.border = BOX
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
        for c in range(6, 11): ws.cell(r, c).border = BOX
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=10)
        for col, m in (('K', p['start']), ('L', p['end']), ('M', p['brk'])):
            x = ws[f'{col}{r}']; x.value = tval(m); x.number_format = 'h:mm'; x.font = A9B; x.alignment = CEN; x.border = BOX; x.fill = HDRF
        x = ws[f'N{r}']; x.value = f'=L{r}-K{r}-M{r}'; x.number_format = 'h:mm'; x.font = A9B; x.alignment = CEN; x.border = BOX; x.fill = CALC
        for s, e, lab, kind, cr in shift_segments(p):
            t = label(kind, lab)
            txt = f"{t}\n{h4(s)}-{h4(e)}" if e - s >= 25 else t
            bar(ws, r, s, e, txt, UCOL[cr] if kind in ('main', 'gate', 'arv') else None, kind in ('main', 'gate', 'arv'))
        ws.row_dimensions[r].height = 17.25; r += 1
    last = r - 1
    if naos:
        x = ws.cell(r, 3, f'NAOS 派遣（7C増便 18:30発・{len(naos)}名・拘束4時間）'); x.font = A9B
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=10); r += 1
        for p in naos:
            x = ws.cell(r, 3, p['code']); x.font = A9B; x.alignment = CEN; x.border = BOX; x.fill = NAOSF
            x = ws.cell(r, 4, 'NAOS'); x.font = Font(name='Arial', size=8, bold=True); x.alignment = CEN; x.border = BOX; x.fill = NAOSF
            ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
            for c in range(6, 11): ws.cell(r, c).border = BOX
            for col, m in (('K', p['start']), ('L', p['end']), ('M', 0)):
                x = ws[f'{col}{r}']; x.value = tval(m); x.number_format = 'h:mm'; x.font = A9B; x.alignment = CEN; x.border = BOX; x.fill = NAOSF
            x = ws[f'N{r}']; x.value = f'=L{r}-K{r}-M{r}'; x.number_format = 'h:mm'; x.font = A9B; x.alignment = CEN; x.border = BOX; x.fill = NAOSF
            for s_, e_, lab, kind, cr in fill_ends(p, segments(dict(p))):
                t = label(kind, lab)
                bar(ws, r, s_, e_, f"{t}\n{h4(s_)}-{h4(e_)}" if e_ - s_ >= 25 else t, UCOL[cr] if kind in ('main', 'gate', 'arv') else None, kind in ('main', 'gate', 'arv'))
            ws.row_dimensions[r].height = 17.25; r += 1
    for lab, fm, fmt in (('AVERAGE', f'=AVERAGE(N{first}:N{last})', 'h:mm'), ('総労働時間', f'=SUM(N{first}:N{last})', '[h]:mm'), ('出面数', len(shifts), None)):
        x = ws.cell(r, 3, lab); x.font = A9B; x.alignment = CEN
        x = ws.cell(r, 14, fm); x.font = A9B; x.alignment = CEN; x.fill = CALC
        if fmt: x.number_format = fmt
        r += 1
    # shift term counts
    r += 1
    ws.cell(r, 3, 'シフト別 必要人数').font = A10B; r += 1
    for col, v in zip('CDFG', ['Shift', '開始', '終了', '人数']):
        x = ws[f'{col}{r}']; x.value = v; x.font = A9B; x.alignment = CEN; x.border = BOX; x.fill = HDRF
    r += 1
    cnt = Counter(p['code'] for p in shifts); se = {p['code']: (p['start'], p['end']) for p in shifts}
    for code in sorted(cnt, key=lambda c: se[c]):
        for col, v in zip('CDFG', [code, tval(se[code][0]), tval(se[code][1]), cnt[code]]):
            x = ws[f'{col}{r}']; x.value = v; x.font = Font(name='Arial', size=9); x.alignment = CEN; x.border = BOX
            if col in 'DF': x.number_format = 'h:mm'
        r += 1
    x = ws.cell(r, 3, '自社 計'); x.font = A9B; x.border = BOX; x.alignment = CEN
    x = ws.cell(r, 7, len(shifts)); x.font = A9B; x.border = BOX; x.alignment = CEN; r += 1
    if naos:
        for col, v in zip('CDFG', [f"NAOS {naos[0]['code']}", tval(naos[0]['start']), tval(naos[0]['end']), len(naos)]):
            x = ws[f'{col}{r}']; x.value = v; x.font = Font(name='Arial', size=9); x.alignment = CEN; x.border = BOX; x.fill = NAOSF
            if col in 'DF': x.number_format = 'h:mm'
        r += 1
    for rr in list(range(H, SH - 1)) + list(range(SH, last + 1 + (len(naos) + 1 if naos else 0))):
        for h in range(4, 24):
            x = ws.cell(rr, m2c(h*60)); b = x.border
            x.border = Border(left=b.left if b.left and b.left.style else hourln, right=b.right, top=b.top, bottom=b.bottom)
    ws.freeze_panes = ws.cell(H + 1, X0)
    return cnt, se

def offpairs(sched, ds):
    n = 0; i = 0
    while i < len(ds) - 1:
        if ds[i] not in sched and ds[i+1] not in sched: n += 1; i += 2
        else: i += 1
    return n
def maxconsec(sched, ds):
    m = c = 0
    for d in ds:
        c = c + 1 if d in sched else 0; m = max(m, c)
    return m

wb = openpyxl.Workbook()
ws = wb.active; ws.title = '前提'
ws['B2'] = f'CTS 26W 線表・シフト・ROSTER（11/16〜3/15 の4期間）　{SCEN_TXT}'; ws['B2'].font = Font(name='Meiryo UI', size=14, bold=True)
notes = [
 '【線表の形】「26W必要数」ファイルの線表（シート1〜7）と同じ配置・表記（PRE → 便 DC/CKIN → GA → PW、R、ARV）。便はキャリアごとにまとめ、CTR OPEN 順。',
 '【配置人数】CTS 作成の線表をもとに、出発便 DC 1・GI 1・CKIN 3（D7 は 6、うち2名は CTR STAY）、7C1504 PSPT CHK 2、到着 ARV 全便3名（OPEN SPOT 対応）。CKIN は STD 60分前に GATE（GA）へ移動。',
 '【準備・後作業】DC は CTR OPEN 60分前から PRE、出発後60分 PW。AGT は最低45分（空きがあれば60分）。D7551 は DC・GI・CKIN 3名が 5:30 出勤、残り CKIN 3名が 6:30 出勤。',
 f'【シフト】15分単位。拘束は最大{CAP//60}時間（休憩 R 1時間を含む）、最小勤務 {MINWORK//60}時間。',
 '【労働時間の調整】期間ごとに、短いシフトへ FREE ASGN（空き時間の業務・待機）を足して、全員の月間労働時間が所定時間（上限）に近づくようにした。期間ごとに足した後の最短の実働は「必要人数」シートに記載。',
 '【ROSTER】各期間、2026年Ver ①正社員の所定勤務日数どおりに勤務し、労働時間は所定時間以内。連続勤務は最大5日、2連休を2回以上、勤務間インターバル11時間以上。条件を満たす最少人数を最適化計算で求めた。',
 '【予備（FREE）】人数×所定日数が必要な出面より多い分は、9A18A の予備勤務（FREE）として割り当てた。応援・研修などに使える日。',
 '【NAOS】7C増便（18:30発）は NAOS から毎日7名（15:15〜19:15、拘束4時間）が GI 1・CKIN 3（→GA）・ARV 3 を担当。DC は自社要員で、この便のみ 15:00〜19:00（準備60分・後作業30分）。NAOS は ROSTER に含めない。',
 f'【対象便】{SCEN_TXT}',
 '【前提】2027年1/16以降の所定日数・時間は2026年の同じ期間の値を使用。期間をまたぐ連続勤務・インターバルはつないでいない。全員が全キャリア・全ポジションを担当できる前提。名前欄は空欄。',
]
for i, t in enumerate(notes):
    x = ws.cell(4 + i, 2, t); x.font = Font(name='Meiryo UI', size=10); x.alignment = LEFT; ws.row_dimensions[4 + i].height = 32
ws.column_dimensions['A'].width = 2; ws.column_dimensions['B'].width = 150

need = wb.create_sheet('必要人数')
need['B2'] = '期間ごとの必要人数とシフト別人数'; need['B2'].font = Font(name='Meiryo UI', size=12, bold=True)
sumr = 4
for col, v in enumerate(['期間', '必要人数', '所定勤務日数', '所定労働時間', '延べ出面', '予備(FREE)日数', '1人あたり労働時間（最小〜最大）', '最短シフト実働', '線表数'], 2):
    x = need.cell(sumr, col, v); x.font = M9B; x.fill = HDRF; x.border = BOX; x.alignment = CEN
need.row_dimensions[sumr].height = 30
for j, w in enumerate([16, 9, 10, 10, 9, 11, 22, 11, 8]): need.column_dimensions[CL(2 + j)].width = w
need.column_dimensions['A'].width = 2
blocks_row = sumr + len(SOL) + 3

ros_sheets = []
for k, S in enumerate(SOL):
    a, b, ds = S['a'], S['b'], S['ds']
    tag = f"{a.month}-{a.day}"
    groups = {}
    for d in ds: groups.setdefault(S['F'][d], []).append(d)
    gl = sorted(groups.items(), key=lambda x: (min(x[1]).isoweekday() if False else 0, x[1][0]))
    cols = []
    for gi, (F, gds) in enumerate(gl):
        wds = sorted({d.isoweekday() for d in gds})
        pidl = f"{gi+1:02d}"
        nm = f"{tag} {pidl}({''.join(WJ[w] for w in wds)})"
        dates = ', '.join(md(d) for d in gds)
        shifts = sorted(S['sh'][gds[0]], key=lambda p: (p['start'], p['end']))
        naos = S['naos'][gds[0]]
        cnt, se = line_sheet(wb, nm, f"CTS 26W 線表  {md(a)}〜{md(b)} 期間  パターン{pidl}（{'・'.join(WJ[w] for w in wds)}）　{SCEN_TXT[:2]}",
                             f"適用日（{len(gds)}日）: {dates}　　自社 {len(shifts)} 名" + (f"　＋ NAOS {len(naos)} 名" if naos else '') + ('　　HB・SL は委託（本線表に含まず）' if SCEN == 'B' else ''), F, shifts, gds, naos)
        cols.append((nm, pidl, wds, gds, cnt, se, len(naos)))
    # 必要人数: shift x day-type
    r = blocks_row
    x = need.cell(r, 2, f"{md(a)}〜{md(b)}　シフト別の必要人数（列 = 線表パターン）"); x.font = M9B; r += 1
    hdr = ['Shift', '開始', '終了', '実働'] + [f"{c[1]}({''.join(WJ[w] for w in c[2])})\n{len(c[3])}日" for c in cols] + ['期間延べ']
    for j, v in enumerate(hdr):
        x = need.cell(r, 2 + j, v); x.font = M9B; x.fill = HDRF; x.border = BOX; x.alignment = CEN
        if j >= 4 and j < 4 + len(cols): x.hyperlink = f"#'{cols[j-4][0]}'!A1"
    need.row_dimensions[r].height = 30; r += 1
    allse = {}
    for c in cols: allse.update(c[5])
    for code in sorted(allse, key=lambda c: allse[c]):
        s0, e0 = allse[code]; wk = (e0 - s0) - (60 if e0 - s0 > 360 else 0)
        vals = [code, tval(s0), tval(e0 % 1440), round(wk / 60, 2)] + [c[4].get(code, '') for c in cols] + [sum(c[4].get(code, 0) * len(c[3]) for c in cols)]
        for j, v in enumerate(vals):
            x = need.cell(r, 2 + j, v); x.font = M9; x.border = BOX; x.alignment = CEN
            if j in (1, 2): x.number_format = 'h:mm'
        r += 1
    vals = ['NAOS', tval(15*60+15), tval(19*60+15), 4.0] + [c[6] or '' for c in cols] + [sum(c[6] * len(c[3]) for c in cols)]
    for j, v in enumerate(vals):
        x = need.cell(r, 2 + j, v); x.font = M9; x.border = BOX; x.alignment = CEN; x.fill = NAOSF
        if j in (1, 2): x.number_format = 'h:mm'
    r += 1
    vals = ['自社 合計', '', '', ''] + [sum(c[4].values()) for c in cols] + [sum(sum(c[4].values()) * len(c[3]) for c in cols)]
    for j, v in enumerate(vals):
        x = need.cell(r, 2 + j, v); x.font = M9B; x.border = BOX; x.alignment = CEN; x.fill = CALC
    blocks_row = r + 3
    # roster
    R = S['R']; codes_ = S['codes']
    rs = wb.create_sheet(f"ROSTER {tag}")
    rs['B1'] = f"ROSTER {md(a)}〜{md(b)}　必要人数 {S['N']} 名　所定 {S['D']}日・{S['H']}h（連続勤務5日以内・2連休2回以上・インターバル11h）"; rs['B1'].font = M9B
    hdr = ['要員'] + [f"{d.month}/{d.day}\n{WJ[d.isoweekday()]}" for d in ds] + ['勤務日数', '公休数', '労働時間', '所定との差', '2連休', '最大連続']
    for j, v in enumerate(hdr):
        x = rs.cell(3, 2 + j, v); x.font = M9B; x.fill = HDRF; x.border = BOX; x.alignment = CEN
    rs.row_dimensions[3].height = 30
    rs.column_dimensions['A'].width = 2; rs.column_dimensions['B'].width = 8
    for j in range(len(ds)): rs.column_dimensions[CL(3 + j)].width = 7
    for j in range(6): rs.column_dimensions[CL(3 + len(ds) + j)].width = 8
    R = sorted(R, key=lambda p: min((codes_[c][0] for c in p['sched'].values()), default=0))
    for i, p in enumerate(R):
        rr = 4 + i
        x = rs.cell(rr, 2, f"要員{i+1:02d}"); x.font = M9; x.border = BOX; x.alignment = CEN
        for j, d in enumerate(ds):
            c = p['sched'].get(d)
            x = rs.cell(rr, 3 + j, c if c else '公'); x.border = BOX; x.alignment = CEN
            x.font = Font(name='Arial', size=8, bold=bool(c), color='FF000000' if c else 'FF548235')
            if not c: x.fill = OFF
            elif c == 'FREE': x.fill = GRAY
            elif codes_[c][0] < 7*60: x.fill = fill('FFDDEBF7')
            elif codes_[c][0] >= 12*60: x.fill = fill('FFFCE4D6')
        hrs = p['mins'] / 60
        vals = [p['days'], len(ds) - p['days'], round(hrs, 2), round(hrs - S['H'], 2), offpairs(p['sched'], ds), maxconsec(p['sched'], ds)]
        for j, v in enumerate(vals):
            x = rs.cell(rr, 3 + len(ds) + j, v); x.font = M9; x.border = BOX; x.alignment = CEN
        if vals[4] < 2 or vals[5] > MAXCONSEC: rs.cell(rr, 3 + len(ds) + 4).fill = BAD
    rr = 4 + len(R)
    x = rs.cell(rr, 2, '出勤数'); x.font = M9B; x.fill = CALC; x.border = BOX
    for j, d in enumerate(ds):
        x = rs.cell(rr, 3 + j, sum(1 for p in R if d in p['sched'])); x.font = M9B; x.fill = CALC; x.border = BOX; x.alignment = CEN
    x = rs.cell(rr + 1, 2, 'うち予備'); x.font = M9; x.border = BOX
    for j, d in enumerate(ds):
        x = rs.cell(rr + 1, 3 + j, S['free'][d] or ''); x.font = M9; x.border = BOX; x.alignment = CEN
    rs.cell(rr + 3, 2, '色: 青=7時前始業、橙=12時以降始業、灰=予備(FREE 9A18A)、緑=公休。').font = Font(name='Meiryo UI', size=8)
    rs.freeze_panes = 'C4'
    # summary row
    hrs = [p['mins'] / 60 for p in R]
    minwork = min(p['work'] for d in ds for p in S['sh'][d]) / 60
    vals = [f"{md(a)}〜{md(b)}", S['N'], S['D'], S['H'], sum(len(S['sh'][d]) for d in ds), sum(S['free'].values()),
            f"{min(hrs):.2f}〜{max(hrs):.2f}h", f"{minwork:.2f}h", len(cols)]
    for j, v in enumerate(vals):
        x = need.cell(sumr + 1 + k, 2 + j, v); x.font = M9; x.border = BOX; x.alignment = CEN
    need.cell(sumr + 1 + k, 2).hyperlink = f"#'ROSTER {tag}'!A1"
    # checks
    for p in R:
        sd_ = [d for d in ds if d in p['sched']]
        for d0, d1 in zip(sd_, sd_[1:]):
            if d1 == d0 + td(1):
                assert (1440 - codes_[p['sched'][d0]][1]) + codes_[p['sched'][d1]][0] >= INTERVAL
        assert maxconsec(p['sched'], ds) <= MAXCONSEC and offpairs(p['sched'], ds) >= 2 and p['days'] == S['D'] and p['mins'] <= S['H'] * 60
    for d in ds:
        assert Counter(c for p in R for c in [p['sched'].get(d)] if c and c != 'FREE') == Counter(s['code'] for s in S['sh'][d])
need.freeze_panes = 'B5'
# order: 前提, 必要人数, then for each period: ROSTER then line sheets
order = ['前提', '必要人数']
for S in SOL:
    tag = f"{S['a'].month}-{S['a'].day}"
    order += [f"ROSTER {tag}"] + [n for n in wb.sheetnames if n.startswith(tag + ' ')]
wb._sheets = [wb[n] for n in order]
wb.save(OUT)
print('saved', len(wb.sheetnames), 'sheets')
for S in SOL: print(S['a'], 'N', S['N'], 'T', S['T'] / 60)
