import sys, copy, datetime as dt
from datetime import date, timedelta as td
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from ref import *
from periods import FL, D, S, E, ops

SRC, OUT = sys.argv[1], sys.argv[2]
wb_in = openpyxl.load_workbook(SRC)
REF = {int(n): parse_sheet(wb_in[n]) for n in '1234567'}
NEWF = {k for k, f in FLIGHTS.items() if f.get('new')}
ANALOG = {'7CXXX': '7C1503', 'RFX0910': 'RF352', 'RFX1830': 'RF356'}
WJ = ['', '月', '火', '水', '木', '金', '土', '日']
ORDER = ['D7550', 'RFX0910', '7C1501', 'RF352', 'SL392', '7C1551', 'HB382', '7C1503', 'RF356', '5J5128', 'D7552', '7CXXX', 'RFX1830']

def deltas(k, analog=None):
    """(dep delta, arr delta) moving reference times of `analog or k` to Operating Period times of k"""
    src = analog or k
    r = dict(FLIGHTS[src]); r.update(REFTIME.get(src, {}))
    return tm(FLIGHTS[k]['std']) - tm(r['std']), tm(FLIGHTS[k]['sta']) - tm(r['sta'])

def is_flight(t): return t.get('k') is not None

def retime(tasks, k):
    """apply Operating-Period timing to flight k's tasks inside one row (HB 16:00 -> 16:15 etc.)"""
    dd, da = deltas(k)
    if not dd and not da: return tasks
    out = [dict(t) for t in tasks]
    for i, t in enumerate(out):
        if t.get('k') != k: continue
        d = da if t['arr'] else dd
        if not d: continue
        t['s'] += d; t['e'] += d
        if i > 0 and not is_flight(out[i-1]) and out[i-1]['e'] == t['s'] - d: out[i-1]['e'] = t['s']
        if i+1 < len(out) and not is_flight(out[i+1]) and out[i+1]['s'] == t['e'] - d:
            out[i+1]['s'] = t['e']
            if out[i+1]['e'] - out[i+1]['s'] < 30: out[i+1]['e'] = out[i+1]['s'] + 30
    # keep neighbours from overlapping
    for i in range(1, len(out)):
        if out[i]['s'] < out[i-1]['e'] and not is_flight(out[i]): out[i]['s'] = out[i-1]['e']
    return [t for t in out if t['e'] > t['s']]

def trim(tasks):
    fl = [i for i, t in enumerate(tasks) if is_flight(t)]
    if not fl: return None
    a, b = fl[0], fl[-1]
    lead = [t for t in tasks[:a] if t['role'].startswith('PRE')][-1:] if a else []
    lead = [t for t in lead if t['e'] == tasks[a]['s']]
    tail = tasks[b+1:b+2] if b+1 < len(tasks) and tasks[b+1]['role'].startswith('PW') else []
    body = [dict(t) for t in lead + tasks[a:b+1] + tail]
    first = body[0]
    if first['role'].startswith('PRE'):
        untouched = bool(lead) and lead[0] is tasks[0]
        if not untouched and first['e'] - first['s'] < 60: first['s'] = first['e'] - 60
    else:
        body.insert(0, dict(s=first['s']-60, e=first['s'], role='PRE', k=None, arr=False))
    if not body[-1]['role'].startswith('PW'):
        body.append(dict(s=body[-1]['e'], e=body[-1]['e']+30, role='PW', k=None, arr=False))
    return body

def build_rows(ref_no, F):
    _, staff = REF[ref_no]
    rows = []
    for st in staff:
        tasks = [dict(t) for t in st['tasks'] if t['k'] is None or t['k'] in F]
        for k in F: tasks = retime(tasks, k)
        tasks = trim(tasks)
        if tasks: rows.append(dict(tasks=tasks, src=f"参照{ref_no}-{st['no']}"))
    # added flights: crew copied from the analogous flight in sheet 1
    for k in [k for k in ORDER if k in F and k in NEWF]:
        an = ANALOG[k]; dd, da = deltas(k, an)
        for st in REF[1][1]:
            blk = [t for t in st['tasks'] if t['k'] == an]
            if not blk: continue
            # split into contiguous blocks
            groups = [[blk[0]]]
            for t in blk[1:]:
                (groups[-1].append(t) if t['s'] == groups[-1][-1]['e'] else groups.append([t]))
            for g in groups:
                new = [dict(t, k=k, s=t['s'] + (da if t['arr'] else dd), e=t['e'] + (da if t['arr'] else dd)) for t in g]
                rows.append(dict(tasks=trim(new), src=f"増便 ({FLIGHTS[an]['lab']}の配置を流用)", new=True))
    for r in rows:
        r['start'] = r['tasks'][0]['s']; r['end'] = r['tasks'][-1]['e']
        r['brk'] = sum(t['e'] - t['s'] for t in r['tasks'] if t['role'] == 'R')
        r['warn'] = (r['end'] - r['start'] - r['brk'] > 360 and r['brk'] == 0)
    return rows

def choose_ref(F, wds):
    base = set(F) - NEWF
    ok = [n for n, (fl, _) in REF.items() if base <= set(fl)]
    best = min(len(set(REF[n][0]) - base) for n in ok)
    ok = [n for n in ok if len(set(REF[n][0]) - base) == best]
    same = [n for n in ok if n in wds]
    return (same or ok)[0]

# ---------- patterns ----------
days = [S + td(i) for i in range((E - S).days + 1)]
day_pat = {d: frozenset(ops(d)) for d in days}
pats = {}
for d in days: pats.setdefault(day_pat[d], []).append(d)
plist = sorted(pats.items(), key=lambda x: x[1][0])
PID = {F: f"P{i+1:02d}" for i, (F, _) in enumerate(plist)}
PINFO = {}
for F, ds in plist:
    wds = sorted({d.isoweekday() for d in ds})
    ref_no = choose_ref(F, wds)
    PINFO[F] = dict(id=PID[F], dates=ds, wds=wds, ref=ref_no, rows=build_rows(ref_no, F))

def ranges(ds):
    out = []; a = b = ds[0]
    for d in ds[1:]:
        if d == b + td(7): b = d
        else: out.append((a, b)); a = b = d
    out.append((a, b)); return out
def fd(d): return f"{d.month}/{d.day}"
def ddmon(d): return d.strftime('%d%b%y').upper()

# ---------- styles ----------
thin = Side(style='thin', color='FF8C8C8C'); hair = Side(style='hair', color='FFBFBFBF'); hour = Side(style='thin', color='FF595959')
B_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
F8 = Font(name='Meiryo UI', size=8); F8B = Font(name='Meiryo UI', size=8, bold=True); F9B = Font(name='Meiryo UI', size=10, bold=True)
F14 = Font(name='Meiryo UI', size=14, bold=True); FW = Font(name='Meiryo UI', size=8, color='FFFFFFFF', bold=True)
CEN = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
def fill(rgb): return PatternFill('solid', fgColor=rgb)
HDR = fill('FFD9E1F2'); GRAY = fill('FFF2F2F2'); NEWFILL = fill('FFFFF2CC'); WARN = fill('FFFFC7CE')
DARK = {'FF0070C0', 'FFC00000'}
GRID_END = m2c(24*60) - 1  # last column (23:55)

wb = openpyxl.Workbook()

# ---------- sheet: 前提 ----------
ws = wb.active; ws.title = '前提'
ws['B2'] = 'CTS 26W 線表（期間別）— 作成の前提'; ws['B2'].font = F14
notes = [
 '出典: CTS_26W必要数（D7増便.VZ UTS委託）シートの Operating Period、および曜日別線表シート「1」〜「7」（月〜日）。',
 'シーズン: 25OCT26〜27MAR27。運航便の組み合わせが同じ日を1パターン（P01〜）にまとめ、パターンごとに線表シートを作成。',
 '要員配置は、そのパターンの便をすべて含む曜日シート（参照シート）から、運航しない便の作業を外して作成。便がなくなった勤務は削除し、残った勤務は前後の PRE / PW を整えて開始・終了を再計算。',
 '増便（7C 17:30/18:30、RF 09:10/10:00、RF 18:30/19:30）は参照シート「1」の類似便（7C1503/1504、RF352/351、RF356/355）の配置を時間をずらして流用し、専用の勤務として追加（黄色の行）。',
 'D7 550/551: ご指示により 14DEC 以前も毎日運航として扱う。',
 'HB: Operating Period シートの HB382/383（STD 16:15、CTR OPEN 13:45）に合わせ、参照線表の HB880/881（STD 16:00）の出発作業を15分後ろにずらした。',
 '期間の誤記と思われる箇所: 7C増便「24DEC26-27DEC27」→ 27MAR27 まで、RF増便 18:30「20DEC26-31DEC27」→ 31DEC26 まで、として作成。',
 '7C1551/1552 は Operating Period シートの 01DEC26 開始（PDF は 06DEC）を採用。',
 '対象外: VZ570/571（UTS委託）、WE531/532（RAMP のみ。必要数シートに記載なし）。',
 '赤い行は休憩なしで実働6時間を超える勤務（要確認）。曜日番号は 1=月 … 7=日。',
]
for i, t in enumerate(notes):
    c = ws.cell(4+i, 2, f'・{t}'); c.font = Font(name='Meiryo UI', size=10); c.alignment = LEFT
    ws.row_dimensions[4+i].height = 30
ws.column_dimensions['A'].width = 2; ws.column_dimensions['B'].width = 140

r0 = 16
hdr = ['FLT', 'Carrier', 'Type', 'Dest', 'STA', 'STD', 'CTR OPEN', '運航期間・曜日']
tbl = wb.create_sheet('対象便')
for j, h in enumerate(hdr):
    c = tbl.cell(2, 2+j, h); c.font = F8B; c.fill = HDR; c.border = B_ALL; c.alignment = CEN
FLD = dict(FL)
for i, k in enumerate(ORDER):
    f = FLIGHTS[k]
    per = ' / '.join(f"{a}-{b} " + ('Daily' if ds == [1,2,3,4,5,6,7] else 'D' + '/'.join(map(str, ds))) for a, b, ds in FLD[k])
    vals = [f['lab'], f['cr'], f['ac'], f['dest'], f['sta'], f['std'], f['ctr'], per]
    for j, v in enumerate(vals):
        c = tbl.cell(3+i, 2+j, v); c.font = F8; c.border = B_ALL; c.alignment = CEN if j < 7 else LEFT
    tbl.cell(3+i, 2).fill = fill(f['fill']); tbl.cell(3+i, 2).font = FW if f['fill'] in DARK else F8B
for j, w in enumerate([22, 8, 8, 10, 8, 8, 10, 70]): tbl.column_dimensions[CL(2+j)].width = w
tbl.column_dimensions['A'].width = 2
ws.cell(r0, 2, '対象便の一覧は「対象便」シート、期間ごとの対応は「期間一覧」「日別」シートを参照。').font = Font(name='Meiryo UI', size=10)

# ---------- sheet: 期間一覧 ----------
ix = wb.create_sheet('期間一覧')
ix['B2'] = '期間 × 曜日 → 線表パターン（カッコ内は出面数）'; ix['B2'].font = F9B
bps = sorted({D(a) for _, ps in FL for a, b, _ in ps} | {D(b) + td(1) for _, ps in FL for a, b, _ in ps} | {S})
bps = [d for d in bps if S <= d <= E] + [E + td(1)]
heads = ['期間', '日数'] + [f'{WJ[w]}' for w in range(1, 8)]
for j, h in enumerate(heads):
    c = ix.cell(4, 2+j, h); c.font = F8B; c.fill = HDR; c.border = B_ALL; c.alignment = CEN
r = 5
for a, b in zip(bps[:-1], bps[1:]):
    b = b - td(1)
    seg = [a + td(i) for i in range((b - a).days + 1)]
    ix.cell(r, 2, f'{ddmon(a)} – {ddmon(b)}'); ix.cell(r, 3, len(seg))
    for w in range(1, 8):
        dd = [d for d in seg if d.isoweekday() == w]
        v = ''
        if dd:
            P = PINFO[day_pat[dd[0]]]; v = f"{P['id']} ({len(P['rows'])})"
            ix.cell(r, 3+w).hyperlink = f"#'{P['id']}'!A1"
        ix.cell(r, 3+w, v)
    for j in range(9):
        c = ix.cell(r, 2+j); c.font = F8; c.border = B_ALL; c.alignment = CEN
    r += 1
ix.column_dimensions['A'].width = 2; ix.column_dimensions['B'].width = 24; ix.column_dimensions['C'].width = 6
for w in range(1, 8): ix.column_dimensions[CL(3+w)].width = 10

pl = wb.create_sheet('パターン一覧')
r = 2
pl.cell(r, 2, 'パターン一覧（パターン名クリックで線表へ）').font = F9B; r += 2
heads = ['パターン', '出面数', '便数', '参照シート', '適用曜日', '適用日数', '適用日', '運航便']
for j, h in enumerate(heads):
    c = pl.cell(r, 2+j, h); c.font = F8B; c.fill = HDR; c.border = B_ALL; c.alignment = CEN
r += 1
for F, ds in plist:
    P = PINFO[F]
    rg = ', '.join(f"{fd(a)}" + (f"–{fd(b)}" if b != a else '') for a, b in ranges(ds))
    vals = [P['id'], len(P['rows']), len(F), f"「{P['ref']}」({WJ[P['ref']]})", '/'.join(WJ[w] for w in P['wds']), len(ds), rg,
            ', '.join(FLIGHTS[k]['lab'].replace('【増便】', '増') for k in ORDER if k in F)]
    for j, v in enumerate(vals):
        c = pl.cell(r, 2+j, v); c.font = F8; c.border = B_ALL; c.alignment = CEN if j < 6 else LEFT
    pl.cell(r, 2).hyperlink = f"#'{P['id']}'!A1"; pl.cell(r, 2).font = Font(name='Meiryo UI', size=8, color='FF0563C1', underline='single')
    pl.row_dimensions[r].height = 40
    r += 1
for col, w in zip('ABCDEFGHI', [2, 9, 7, 6, 11, 14, 8, 40, 80]): pl.column_dimensions[col].width = w
pl.freeze_panes = 'C5'
ix.freeze_panes = 'B5'

# ---------- sheet: 日別 ----------
dy = wb.create_sheet('日別')
heads = ['日付', '曜日', 'パターン', '出面数', '便数'] + [FLIGHTS[k]['lab'].replace('【増便】', '\n増便') for k in ORDER]
for j, h in enumerate(heads):
    c = dy.cell(2, 2+j, h); c.font = F8B; c.fill = HDR; c.border = B_ALL; c.alignment = CEN
dy.row_dimensions[2].height = 36
for i, d in enumerate(days):
    F = day_pat[d]; P = PINFO[F]; rr = 3+i
    vals = [d, WJ[d.isoweekday()], P['id'], len(P['rows']), len(F)] + ['●' if k in F else '' for k in ORDER]
    for j, v in enumerate(vals):
        c = dy.cell(rr, 2+j, v); c.font = F8; c.border = B_ALL; c.alignment = CEN
        if j >= 5 and v: c.fill = fill(FLIGHTS[ORDER[j-5]]['fill']); c.font = FW if FLIGHTS[ORDER[j-5]]['fill'] in DARK else F8
    dy.cell(rr, 2).number_format = 'yyyy/mm/dd'
    dy.cell(rr, 4).hyperlink = f"#'{P['id']}'!A1"
    if d.isoweekday() in (6, 7):
        for j in range(2): dy.cell(rr, 2+j).fill = GRAY
dy.column_dimensions['A'].width = 2; dy.column_dimensions['B'].width = 11
for j in range(1, 5): dy.column_dimensions[CL(2+j)].width = 7
for j in range(len(ORDER)): dy.column_dimensions[CL(7+j)].width = 9
dy.freeze_panes = 'C3'

# ---------- pattern sheets ----------
def put_bar(ws, r, s, e, text, rgb=None, font=F8, row_fill=None):
    c1, c2 = m2c(s), m2c(e) - 1
    if c2 < c1: c2 = c1
    c = ws.cell(r, c1, text); c.font = font; c.alignment = CEN
    if rgb: c.fill = fill(rgb)
    for cc in range(c1, c2+1):
        x = ws.cell(r, cc); x.border = B_ALL
        if rgb: x.fill = fill(rgb)
    if c2 > c1: ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)

def tval(m): return dt.time(m // 60, m % 60)

for F, ds in plist:
    P = PINFO[F]; ws = wb.create_sheet(P['id'])
    flights = [k for k in ORDER if k in F]
    ws.sheet_view.zoomScale = 70
    ws.column_dimensions['A'].width = 2.3; ws.column_dimensions['B'].width = 4.9
    for col, w in zip('CDEFGHIJKLMN', [9.9, 6, 6, 6.7, 7.9, 22, 9, 9, 7, 7, 8, 8]): ws.column_dimensions[col].width = w
    for c in range(15, GRID_END+2): ws.column_dimensions[CL(c)].width = 1.6 if c >= X0 else 0.7
    # title
    ws['C1'] = f"CTS 26W 線表  {P['id']}"; ws['C1'].font = F14
    rg = ', '.join(f"{ddmon(a)}" + (f"–{ddmon(b)}" if b != a else '') + f"({WJ[a.isoweekday()]})" for a, b in ranges(ds))
    ws['C2'] = f"適用日（{len(ds)}日）: {rg}"; ws['C2'].font = F8; ws['C2'].alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=X0+12*10)
    ws.row_dimensions[2].height = max(15, 13 * (1 + len(rg) // 260))
    ws['C3'] = f"要員配置の参照元: 曜日シート「{P['ref']}」({WJ[P['ref']]}曜)   出面数 {len(P['rows'])} 名   便数 {len(F)}"; ws['C3'].font = F8B
    # header
    H = 5
    for col, v in zip('CDEFGHIJKLMN', ['Terminal', '空港', 'Carrier', 'Type', 'Wide\nNarrow', 'FLT', 'Frequency', 'Dest', 'STA', 'STD', 'CTR OPEN', 'BD']):
        c = ws[f'{col}{H}']; c.value = v; c.font = F8B; c.fill = HDR; c.alignment = CEN; c.border = B_ALL
    def hour_header(r):
        for h in range(4, 24):
            c0 = m2c(h*60)
            c = ws.cell(r, c0, h); c.font = F8B; c.alignment = Alignment(horizontal='left', vertical='center')
            for cc in range(c0, c0+12): ws.cell(r, cc).fill = HDR
    hour_header(H)
    ws.row_dimensions[H].height = 24
    FDICT = dict(FL)
    r = H + 1
    for k in flights:
        f = FLIGHTS[k]
        ds_k = [ds_ for a, b, ds_ in FDICT[k]]
        freq = 'Daily' if [1,2,3,4,5,6,7] in ds_k else ' '.join(map(str, sorted({x for y in ds_k for x in y})))
        vals = ['T1', 'CTS', f['cr'], f['ac'], f['wn'], f['lab'], freq, f['dest'], f['sta'], f['std'], f['ctr'], '']
        for col, v in zip('CDEFGHIJKLMN', vals):
            c = ws[f'{col}{r}']; c.value = v; c.font = F8; c.alignment = CEN; c.border = B_ALL
            if f.get('new'): c.fill = NEWFILL
        a, b, sta = tm(f['ctr']), tm(f['std']), tm(f['sta'])
        put_bar(ws, r, a - 30, a, f['sta'] if False else '', None)
        put_bar(ws, r, a, b, f"{f['lab'].replace('【増便】','')}\n(STA {f['sta']} / CTR OPEN {f['ctr']})", f['fill'], FW if f['fill'] in DARK else F8B)
        ws.cell(r, m2c(b), f['std']).font = F8
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1
    SH = r
    for col, v in zip('CDFKLMN', ['Shift Code', 'Duty', '備考', '開始', '終了', '休憩時間', '労働時間']):
        c = ws[f'{col}{SH}']; c.value = v; c.font = F8B; c.fill = HDR; c.alignment = CEN; c.border = B_ALL
    ws.merge_cells(f'D{SH}:E{SH}'); ws.merge_cells(f'F{SH}:J{SH}')
    hour_header(SH)
    ws.row_dimensions[SH].height = 24
    r = SH + 1
    first = r
    for i, row in enumerate(P['rows']):
        ws.cell(r, 2, i+1).font = F8
        ws.cell(r, 6, row['src']).font = Font(name='Meiryo UI', size=7, color='FF7F7F7F')
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=10)
        for col, m in zip('KLM', [row['start'], row['end'], row['brk']]):
            c = ws[f'{col}{r}']; c.value = tval(m); c.number_format = 'hh:mm'; c.font = F8; c.alignment = CEN; c.border = B_ALL
        c = ws[f'N{r}']; c.value = f'=L{r}-K{r}-M{r}'; c.number_format = 'h:mm'; c.font = F8; c.alignment = CEN; c.border = B_ALL; c.fill = fill('FFD0E0E3')
        if row.get('new'):
            for col in 'BCDE': ws[f'{col}{r}'].fill = NEWFILL
        if row['warn']:
            for col in 'KLM': ws[f'{col}{r}'].fill = WARN
        prev_k = None
        for t in row['tasks']:
            k = t['k']; role = t['role']
            if k:
                f = FLIGHTS[k]
                code = f['arr'] if t['arr'] else f['dep']
                label = f"{code} {role}" if (prev_k != k or role in ('DC', 'CKIN') or t['arr']) else role
                txt = f"{label}\n{hm(t['s'])}-{hm(t['e'])}"
                put_bar(ws, r, t['s'], t['e'], txt, f['fill'], FW if f['fill'] in DARK else F8)
            else:
                put_bar(ws, r, t['s'], t['e'], f"{role}\n{hm(t['s'])}-{hm(t['e'])}" if t['e'] - t['s'] >= 30 else role, None)
            prev_k = k
        ws.row_dimensions[r].height = 30
        r += 1
    last = r - 1
    ws.cell(r, 3, 'AVERAGE').font = F8B; c = ws.cell(r, 14, f'=AVERAGE(N{first}:N{last})'); c.number_format = 'h:mm'; c.font = F8B
    ws.cell(r+1, 3, '総労働時間').font = F8B; c = ws.cell(r+1, 14, f'=SUM(N{first}:N{last})'); c.number_format = '[h]:mm'; c.font = F8B
    ws.cell(r+2, 3, '出面数').font = F8B; ws.cell(r+2, 14, len(P['rows'])).font = F8B
    # hour guide lines
    for rr in list(range(H, SH - 1)) + list(range(SH, last + 1)):
        for h in range(4, 24):
            c = ws.cell(rr, m2c(h*60)); b = c.border
            c.border = Border(left=hour, right=b.right, top=b.top, bottom=b.bottom)
    ws.freeze_panes = ws.cell(H+1, X0)
    ws.sheet_properties.tabColor = 'FFB4C6E7' if not (F & NEWF) else 'FFFFD966'

wb.save(OUT)
print('patterns', len(plist))
for F, ds in plist:
    P = PINFO[F]
    print(P['id'], len(ds), 'days', 'ref', P['ref'], 'staff', len(P['rows']), 'warn', sum(r['warn'] for r in P['rows']), sorted(F))
