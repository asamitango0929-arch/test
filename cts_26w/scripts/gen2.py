import sys, math, datetime as dt
from datetime import date, timedelta as td
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from ref import FLIGHTS, tm, hm, m2c, X0
from periods import FL, D, S, E
from opt import build, jobs_for, days, day_pat, PAY, PRE, PW, GAP, brk

OUT = sys.argv[1]
CAP, MAXGAP = 13*60, 600
HOLD = 50  # 10末時点保有数（26W必要数シート）
WJ = ['', '月', '火', '水', '木', '金', '土', '日']
NEWF = {k for k, f in FLIGHTS.items() if f.get('new')}

def order_flights(F):
    """carrier blocks ordered by the carrier's earliest CTR OPEN, flights inside a block by CTR OPEN"""
    by = {}
    for k in F: by.setdefault(FLIGHTS[k]['cr'], []).append(k)
    for v in by.values(): v.sort(key=lambda k: (tm(FLIGHTS[k]['ctr']), tm(FLIGHTS[k]['std'])))
    crs = sorted(by, key=lambda c: tm(FLIGHTS[by[c][0]]['ctr']))
    return [k for c in crs for k in by[c]]

pats = {}
for d in days: pats.setdefault(day_pat[d], []).append(d)
plist = sorted(pats.items(), key=lambda x: x[1][0])
PI = {}
for i, (F, ds) in enumerate(plist):
    rows = build(jobs_for(F), CAP, MAXGAP)
    wds = sorted({d.isoweekday() for d in ds})
    PI[F] = dict(id=f"P{i+1:02d}", ds=ds, wds=wds, rows=rows, P=len(rows), H=sum(r['work'] for r in rows) / 60)

def ranges(ds):
    out = []; a = b = ds[0]
    for d in ds[1:]:
        if d == b + td(7): b = d
        else: out.append((a, b)); a = b = d
    out.append((a, b)); return out
fd = lambda d: f"{d.month}/{d.day}"
ddmon = lambda d: d.strftime('%d%b%y').upper()
def sheetname(P): return f"{P['id']}({''.join(WJ[w] for w in P['wds'])})"

# ---------- styles ----------
thin = Side(style='thin', color='FF8C8C8C'); hour = Side(style='thin', color='FF595959')
B_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
FN = 'Meiryo UI'
F8 = Font(name=FN, size=8); F8B = Font(name=FN, size=8, bold=True); F10 = Font(name=FN, size=10); F10B = Font(name=FN, size=10, bold=True)
F14 = Font(name=FN, size=14, bold=True); FW = Font(name=FN, size=8, color='FFFFFFFF', bold=True); FG = Font(name=FN, size=7, color='FF7F7F7F')
FLINK = Font(name=FN, size=8, color='FF0563C1', underline='single')
CEN = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
fill = lambda rgb: PatternFill('solid', fgColor=rgb)
HDR = fill('FFD9E1F2'); NEWFILL = fill('FFFFF2CC'); RFILL = fill('FFE7E6E6'); BAD = fill('FFFFC7CE'); GOOD = fill('FFC6EFCE'); CALC = fill('FFD0E0E3')
DARK = {'FF0070C0', 'FFC00000'}
def ffont(k): return FW if FLIGHTS[k]['fill'] in DARK else F8B
def head(ws, r, c0, heads, widths=None):
    for j, h in enumerate(heads):
        c = ws.cell(r, c0+j, h); c.font = F8B; c.fill = HDR; c.border = B_ALL; c.alignment = CEN
    if widths:
        for j, w in enumerate(widths): ws.column_dimensions[CL(c0+j)].width = w
def cell(ws, r, c, v, fmt=None, al=CEN, font=F8, fl=None):
    x = ws.cell(r, c, v); x.font = font; x.border = B_ALL; x.alignment = al
    if fmt: x.number_format = fmt
    if fl: x.fill = fl
    return x

wb = openpyxl.Workbook()

# ---------- 前提 ----------
ws = wb.active; ws.title = '前提'
ws['B2'] = 'CTS 26W 線表 — 作成の前提'; ws['B2'].font = F14
notes = [
 '【スケジュール欄】便はキャリアごとにまとめ、キャリアの並びはそのキャリアで一番早いカウンターオープン順、同じキャリア内もカウンターオープン順。',
 '【要員配置】少ない人数で回すことを優先して組み直した。各便の作業の固まり（DC、CKIN＋GC/GA/SC/CTR、AS/AR/BUS など）と人数は、元の線表シート「1」の配置をそのまま使用。',
 f'【勤務の組み方】始業は最初の作業の{PRE}分前（PRE）、終業は最後の作業の{PW}分後（PW）。作業と作業の間は移動として最低{GAP}分空ける。',
 '【休憩】拘束6時間超は45分、8時間45分超は60分の休憩（R）を、作業の空き時間に必ず入れる（空き時間がない組み合わせは作らない）。',
 f'【無視した条件】ご指示により、勤務間インターバルと1日の最大勤務時間は考慮していない（拘束は最大{CAP//60}時間まで許容）。',
 '【期間の調整】「Period別」シートで、給与期間（16日〜翌月15日）ごとに延べ出面（人日）と延べ労働時間を集計し、①正社員の勤務日数・時間の基準から必要人数を算出。',
 '【基準値】2026年Ver の①正社員を使用。2027年1/16以降の期間は表に値がないため、2026年の同じ期間の値（1/16〜2/15: 21日・168h、2/16〜3/15: 19日・152h、3/16〜4/15: 21日・168h）を仮に使用。',
 '【シーズンの端】10/16〜11/15 は 10/25 から、3/16〜4/15 は 3/27 までしかデータがないため、基準日数・時間を日数比で按分して比較。',
 f'【保有数】「26W必要数」シートの10末時点保有数 {HOLD} 名と比較。',
 '【便・期間】Operating Period シートに従う。D7550 は14DEC以前も毎日、HB は HB382/383（STD 16:15）、7C増便は 27MAR27 まで、RF増便18:30 は 31DEC26 までとして作成。VZ（UTS委託）と WE（RAMPのみ）は対象外。',
]
for i, t in enumerate(notes):
    c = ws.cell(4+i, 2, t); c.font = F10; c.alignment = LEFT; ws.row_dimensions[4+i].height = 32
ws.column_dimensions['A'].width = 2; ws.column_dimensions['B'].width = 150

# ---------- Period別 ----------
pp = wb.create_sheet('Period別')
pp['B2'] = '給与期間（Period）ごとの勤務日数・労働時間と必要人数（①正社員基準）'; pp['B2'].font = F10B
heads = ['Period', '暦日数', '26Wの日数', '基準\n勤務日数', '基準\n労働時間', '按分後\n勤務日数', '按分後\n労働時間', '延べ出面\n（人日）', '延べ労働時間\n（h）',
         '1出面あたり\n平均時間', '必要人数\n（日数で計算）', '必要人数\n（時間で計算）', '必要人数', '決め手',
         '必要人数での1人あたり\n勤務日数（Period換算）', '必要人数での1人あたり\n労働時間（Period換算）', f'保有{HOLD}名での1人あたり\n勤務日数（Period換算）', f'保有{HOLD}名での1人あたり\n労働時間（Period換算）', f'保有{HOLD}名との差']
head(pp, 4, 2, heads, [20, 7, 8, 8, 8, 8, 8, 9, 10, 9, 10, 10, 9, 9, 12, 12, 12, 12, 10])
pp.row_dimensions[4].height = 44
r = 5
for a, b, sd, sh in PAY:
    ds = [d for d in days if a <= d <= b]; pdays = (b - a).days + 1; frac = len(ds) / pdays
    P = sum(PI[day_pat[d]]['P'] for d in ds); H = sum(PI[day_pat[d]]['H'] for d in ds)
    nd, nh = P / (sd*frac), H / (sh*frac); N = math.ceil(max(nd, nh) - 1e-9)
    vals = [f'{a.month}/{a.day}〜{b.month}/{b.day}', pdays, len(ds), sd, sh, round(sd*frac, 1), round(sh*frac, 1), P, round(H, 1), round(H/P, 2),
            round(nd, 1), round(nh, 1), N, '勤務日数' if nd >= nh else '労働時間',
            round(P/N/frac, 1), round(H/N/frac, 1), round(P/HOLD/frac, 1), round(H/HOLD/frac, 1), HOLD - N]
    for j, v in enumerate(vals):
        x = cell(pp, r, 2+j, v)
    pp.cell(r, 14).font = Font(name=FN, size=10, bold=True)
    pp.cell(r, 20).fill = GOOD if HOLD - N >= 0 else BAD
    for j in (17, 18):
        over = (vals[j-2] > (sd if j == 17 else sh))
        pp.cell(r, 2+j).fill = BAD if over else GOOD
    r += 1
notes = [
 '・必要人数 = 「延べ出面 ÷ 按分後勤務日数」と「延べ労働時間 ÷ 按分後労働時間」の大きい方（切り上げ）。決め手はどちらで決まったか。',
 '・決め手が「勤務日数」の期間は、1出面あたりの平均時間が8時間に届かず、人数は日数で決まっている。この場合、必要人数で回すと1人あたりの労働時間は基準より短くなる（基準時間までの残りを他業務・研修等に充てられる）。',
 '・決め手が「労働時間」の期間は、1出面あたりの平均時間が8時間を超えており、時間側で人数が決まる。',
 f'・保有{HOLD}名での1人あたり勤務日数・労働時間が基準を超える期間（赤）は、保有数では足りない期間。',
]
for i, t in enumerate(notes):
    c = pp.cell(r+1+i, 2, t); c.font = F8
pp.column_dimensions['A'].width = 2

# ---------- 期間一覧 (Operating Period) ----------
ix = wb.create_sheet('期間一覧')
ix['B2'] = '運航期間（Operating Period）× 曜日 → 線表シート（カッコ内は出面数）'; ix['B2'].font = F10B
bps = sorted({D(a) for _, ps in FL for a, b, _ in ps} | {D(b) + td(1) for _, ps in FL for a, b, _ in ps} | {S})
bps = [d for d in bps if S <= d <= E] + [E + td(1)]
head(ix, 4, 2, ['#', '運航期間', '日数'] + [WJ[w] for w in range(1, 8)], [5, 22, 6] + [13]*7)
r = 5
for n, (a, b) in enumerate(zip(bps[:-1], bps[1:]), 1):
    b = b - td(1); seg = [a + td(i) for i in range((b - a).days + 1)]
    cell(ix, r, 2, n); cell(ix, r, 3, f'{ddmon(a)} – {ddmon(b)}' if a != b else ddmon(a)); cell(ix, r, 4, len(seg))
    for w in range(1, 8):
        dd = [d for d in seg if d.isoweekday() == w]
        x = cell(ix, r, 4+w, '')
        if dd:
            P = PI[day_pat[dd[0]]]; x.value = f"{P['id']} ({P['P']})"; x.hyperlink = f"#'{sheetname(P)}'!A1"; x.font = FLINK
    r += 1
ix.column_dimensions['A'].width = 2
r += 1
ix.cell(r, 2, '・P01〜は「その日に運航する便の組み合わせ」の番号。組み合わせが同じ日は同じ線表になるため、1枚にまとめている。').font = F8
ix.freeze_panes = 'E5'

# ---------- パターン一覧 ----------
pl = wb.create_sheet('パターン一覧')
pl['B2'] = '線表シート一覧（シート名クリックで移動）'; pl['B2'].font = F10B
head(pl, 4, 2, ['線表シート', '出面数', '延べ労働時間', '平均時間', '便数', '適用曜日', '適用日数', '適用日', '運航便（キャリア順）'], [12, 7, 10, 8, 6, 10, 8, 40, 80])
r = 5
for F, ds in plist:
    P = PI[F]
    rg = ', '.join(fd(a) + (f"–{fd(b)}" if b != a else '') for a, b in ranges(ds))
    vals = [sheetname(P), P['P'], round(P['H'], 1), round(P['H']/P['P'], 2), len(F), '/'.join(WJ[w] for w in P['wds']), len(ds), rg,
            ', '.join(FLIGHTS[k]['lab'].replace('【増便】', '増') for k in order_flights(F))]
    for j, v in enumerate(vals): cell(pl, r, 2+j, v, al=CEN if j < 7 else LEFT)
    pl.cell(r, 2).hyperlink = f"#'{sheetname(P)}'!A1"; pl.cell(r, 2).font = FLINK
    pl.row_dimensions[r].height = 36; r += 1
pl.column_dimensions['A'].width = 2; pl.freeze_panes = 'C5'

# ---------- 日別 ----------
dy = wb.create_sheet('日別')
ALLK = order_flights(set(FLIGHTS))
head(dy, 2, 2, ['日付', '曜日', 'Period', '線表シート', '出面数', '労働時間', '便数'] + [FLIGHTS[k]['lab'].replace('【増便】', '\n増便') for k in ALLK],
     [11, 5, 12, 11, 7, 8, 6] + [9]*len(ALLK))
dy.row_dimensions[2].height = 36
for i, d in enumerate(days):
    F = day_pat[d]; P = PI[F]; rr = 3+i
    per = next(f'{a.month}/{a.day}〜{b.month}/{b.day}' for a, b, _, _ in PAY if a <= d <= b)
    for j, v in enumerate([d, WJ[d.isoweekday()], per, sheetname(P), P['P'], round(P['H'], 1), len(F)]):
        cell(dy, rr, 2+j, v, fmt='yyyy/mm/dd' if j == 0 else None)
    dy.cell(rr, 5).hyperlink = f"#'{sheetname(P)}'!A1"; dy.cell(rr, 5).font = FLINK
    for j, k in enumerate(ALLK):
        x = cell(dy, rr, 9+j, '●' if k in F else '')
        if k in F: x.fill = fill(FLIGHTS[k]['fill']); x.font = FW if FLIGHTS[k]['fill'] in DARK else F8
dy.column_dimensions['A'].width = 2; dy.freeze_panes = 'D3'

# ---------- line chart sheets ----------
def bar(ws, r, s, e, text, rgb=None, font=F8, fl=None):
    c1, c2 = m2c(s), max(m2c(e) - 1, m2c(s))
    x = ws.cell(r, c1, text); x.font = font; x.alignment = CEN
    for cc in range(c1, c2+1):
        y = ws.cell(r, cc); y.border = B_ALL
        if rgb: y.fill = fill(rgb)
        elif fl: y.fill = fl
    if c2 > c1: ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
tval = lambda m: dt.time(m // 60, m % 60)
def hour_header(ws, r):
    for h in range(4, 24):
        c0 = m2c(h*60); x = ws.cell(r, c0, h); x.font = F8B; x.alignment = Alignment(horizontal='left', vertical='center')
        for cc in range(c0, c0+12): ws.cell(r, cc).fill = HDR
GRID_END = m2c(24*60) - 1

for F, ds in plist:
    P = PI[F]; ws = wb.create_sheet(sheetname(P)); fl = order_flights(F)
    ws.sheet_view.zoomScale = 70
    for col, w in zip('ABCDEFGHIJKLMN', [2.3, 4.9, 8, 6, 6, 6.7, 7.9, 22, 9, 9, 7, 7, 8, 8]): ws.column_dimensions[col].width = w
    for c in range(15, GRID_END + 2): ws.column_dimensions[CL(c)].width = 1.6 if c >= X0 else 0.7
    ws['C1'] = f"CTS 26W 線表  {P['id']}（{'・'.join(WJ[w] for w in P['wds'])}）"; ws['C1'].font = F14
    rg = ', '.join(ddmon(a) + (f"–{ddmon(b)}" if b != a else '') + f"({WJ[a.isoweekday()]})" for a, b in ranges(ds))
    ws['C2'] = f"適用日（{len(ds)}日）: {rg}"; ws['C2'].font = F8; ws['C2'].alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=X0 + 12*10)
    ws.row_dimensions[2].height = max(15, 13 * (1 + len(rg) // 230))
    ws['C3'] = f"出面数 {P['P']} 名   延べ労働時間 {P['H']:.1f} h   平均 {P['H']/P['P']:.2f} h/人   便数 {len(F)}"; ws['C3'].font = F8B
    H = 5
    head(ws, H, 3, ['Terminal', '空港', 'Carrier', 'Type', 'Wide\nNarrow', 'FLT', 'Frequency', 'Dest', 'STA', 'STD', 'CTR OPEN', 'BD'])
    hour_header(ws, H); ws.row_dimensions[H].height = 24
    FD = dict(FL); r = H + 1
    for k in fl:
        f = FLIGHTS[k]; dsk = [x for _, _, x in FD[k]]
        freq = 'Daily' if [1,2,3,4,5,6,7] in dsk else ' '.join(map(str, sorted({x for y in dsk for x in y})))
        for j, v in enumerate(['T1', 'CTS', f['cr'], f['ac'], f['wn'], f['lab'], freq, f['dest'], f['sta'], f['std'], f['ctr'], '']):
            cell(ws, r, 3+j, v, fl=NEWFILL if f.get('new') else None)
        a, b = tm(f['ctr']), tm(f['std'])
        bar(ws, r, a, b, f"{f['lab'].replace('【増便】','')}\n(STA {f['sta']} / CTR OPEN {f['ctr']})", f['fill'], ffont(k))
        ws.cell(r, m2c(b), f['std']).font = F8
        ws.row_dimensions[r].height = 30; r += 1
    r += 1; SH = r
    head(ws, SH, 2, ['No', 'Shift Code', 'Duty', '', '担当便（順）', '', '', '', '', '開始', '終了', '休憩時間', '労働時間'])
    ws.merge_cells(f'D{SH}:E{SH}'); ws.merge_cells(f'F{SH}:J{SH}')
    hour_header(ws, SH); ws.row_dimensions[SH].height = 24
    r = SH + 1; first = r
    for i, p in enumerate(P['rows']):
        cell(ws, r, 2, i + 1)
        route = '→'.join(dict.fromkeys(FLIGHTS[j['k']]['lab'].split('/')[0].replace('【増便】', '') + ('増' if j['k'] in NEWF else '') for j in p['jobs']))
        x = ws.cell(r, 6, route); x.font = FG; x.alignment = LEFT
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=10)
        for col, m in zip('KLM', [p['start'], p['end'], p['brk']]): cell(ws, r, ' KLM'.index(col) + 10, tval(m), fmt='hh:mm')
        cell(ws, r, 14, f'=L{r}-K{r}-M{r}', fmt='h:mm', fl=CALC)
        bar(ws, r, p['start'], p['jobs'][0]['s'], f"PRE\n{hm(p['start'])}-{hm(p['jobs'][0]['s'])}")
        for j in p['jobs']:
            k = j['k']; f = FLIGHTS[k]; code = f['arr'] if j['arr'] else f['dep']
            role = '/'.join(t['role'] for t in j['tasks'])
            bar(ws, r, j['s'], j['e'], f"{code} {role}\n{hm(j['s'])}-{hm(j['e'])}", f['fill'], ffont(k))
        if p['rest']:
            a, b = p['rest']; bar(ws, r, a, b, f"R\n{hm(a)}-{hm(b)}", fl=RFILL)
        bar(ws, r, p['jobs'][-1]['e'], p['end'], f"PW\n{hm(p['jobs'][-1]['e'])}-{hm(p['end'])}")
        ws.row_dimensions[r].height = 30; r += 1
    last = r - 1
    for lab, fm, fmt in [('AVERAGE', f'=AVERAGE(N{first}:N{last})', 'h:mm'), ('総労働時間', f'=SUM(N{first}:N{last})', '[h]:mm'), ('出面数', P['P'], None)]:
        ws.cell(r, 3, lab).font = F8B; x = ws.cell(r, 14, fm); x.font = F8B
        if fmt: x.number_format = fmt
        r += 1
    for rr in list(range(H, SH - 1)) + list(range(SH, last + 1)):
        for h in range(4, 24):
            x = ws.cell(rr, m2c(h*60)); b = x.border
            x.border = Border(left=hour, right=b.right, top=b.top, bottom=b.bottom)
    ws.freeze_panes = ws.cell(H + 1, X0)
    ws.sheet_properties.tabColor = 'FFFFD966' if F & NEWF else 'FFB4C6E7'

wb.save(OUT)
for F, ds in plist: print(sheetname(PI[F]), PI[F]['P'], round(PI[F]['H'], 1))
