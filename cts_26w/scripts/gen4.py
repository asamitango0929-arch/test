import sys, datetime as dt
from collections import Counter, defaultdict
from datetime import date, timedelta as td
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from ref import FLIGHTS, tm, hm, m2c, X0
hc = lambda m: f'{m//60:02d}:{m%60:02d}'
from periods import FL, D, S, E
import v3
from v3 import (flight_jobs, shifts_for, day_pat, days, PAY, solve_period, DEP, ARR, ARR_N, PSPT, PREPOST, AGT_MAX,
                COLOR, CRNAME, CAP, REST, INTERVAL, MAXCONSEC, codes)
import cts

OUT, CTSX, OLDX = sys.argv[1], sys.argv[2], sys.argv[3]
WJ = ['', '月', '火', '水', '木', '金', '土', '日']
NEWF = {k for k, f in FLIGHTS.items() if f.get('new')}
HOLD = 50

def ctr_open(k): f = FLIGHTS[k]; return tm(f['std']) + DEP[f['cr']]['win'][0]
def order_flights(F):
    by = {}
    for k in F: by.setdefault(FLIGHTS[k]['cr'], []).append(k)
    for v in by.values(): v.sort(key=lambda k: (ctr_open(k), tm(FLIGHTS[k]['std'])))
    return [k for c in sorted(by, key=lambda c: ctr_open(by[c][0])) for k in by[c]]

pats = {}
for d in days: pats.setdefault(day_pat[d], []).append(d)
plist = sorted(pats.items(), key=lambda x: x[1][0])
PI = {F: dict(id=f"P{i+1:02d}", ds=ds, wds=sorted({d.isoweekday() for d in ds}), rows=shifts_for(F)) for i, (F, ds) in enumerate(plist)}
for P in PI.values(): P['P'] = len(P['rows']); P['H'] = sum(r['work'] for r in P['rows']) / 60
sname = lambda P: f"{P['id']}({''.join(WJ[w] for w in P['wds'])})"

def ranges(ds):
    out = []; a = b = ds[0]
    for d in ds[1:]:
        if d == b + td(7): b = d
        else: out.append((a, b)); a = b = d
    out.append((a, b)); return out
fd = lambda d: f"{d.month}/{d.day}"
ddmon = lambda d: d.strftime('%d%b%y').upper()

# ---------- styles ----------
thin = Side(style='thin', color='FF8C8C8C'); hour = Side(style='thin', color='FF595959')
B_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
FN = 'Meiryo UI'
F8 = Font(name=FN, size=8); F8B = Font(name=FN, size=8, bold=True); F10 = Font(name=FN, size=10); F10B = Font(name=FN, size=10, bold=True)
F14 = Font(name=FN, size=14, bold=True); FG = Font(name=FN, size=7, color='FF7F7F7F'); FLINK = Font(name=FN, size=8, color='FF0563C1', underline='single')
FWHITE = Font(name=FN, size=8, bold=True, color='FFFFFFFF')
CEN = Alignment(horizontal='center', vertical='center', wrap_text=True); LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
fill = lambda rgb: PatternFill('solid', fgColor=rgb)
HDR = fill('FFD9E1F2'); NEWFILL = fill('FFFFF2CC'); REDF = fill('FFFF0000'); BAD = fill('FFFFC7CE'); GOOD = fill('FFC6EFCE'); CALC = fill('FFD0E0E3'); GRAY = fill('FFF2F2F2')
def head(ws, r, c0, heads, widths=None, h=None):
    for j, x in enumerate(heads):
        c = ws.cell(r, c0+j, x); c.font = F8B; c.fill = HDR; c.border = B_ALL; c.alignment = CEN
    if widths:
        for j, w in enumerate(widths): ws.column_dimensions[CL(c0+j)].width = w
    if h: ws.row_dimensions[r].height = h
def cell(ws, r, c, v, fmt=None, al=CEN, font=F8, fl=None):
    x = ws.cell(r, c, v); x.font = font; x.border = B_ALL; x.alignment = al
    if fmt: x.number_format = fmt
    if fl: x.fill = fl
    return x

wb = openpyxl.Workbook()
# ---------- 前提 ----------
ws = wb.active; ws.title = '前提'
ws['B2'] = 'CTS 26W 線表（パターン） — 作成の前提'; ws['B2'].font = F14
notes = [
 '【配置人数】CTS作成の線表（1DEC〜15DEC）で実際に配置しているポジションと人数をキャリアごとに集計し、その標準形で全便を配置（「キャリア別配置」シート）。',
 '【準備・後作業】DC は CTR OPEN の60分前から準備（DP）、出発後60分の後作業（PW）。出発便の AGT（GI・CTR・PSPT CHK）は準備・後作業とも最低45分、前後が空いていれば最大60分まで延ばして表示。到着（ARV）は CTS 線表の時間どおり。',
 f'【勤務】1日の拘束は最大{CAP//60}時間（休憩1時間を含む）。拘束6時間を超える勤務には休憩（REST）60分を作業の空き時間に入れる。勤務時間は CTS と同じく15分単位のシフトコード（例 5C14C = 5:30〜14:30、A=00/B=15/C=30/D=45）。',
 '【D7朝便】D7550/551 は6:00出勤を基本とし、DC の準備は 6:00〜6:55。勤務の開始・終了は15分単位。',
 '【到着】CTS の冬季 OPEN SPOT 対応のため、到着（ARV）は全便3名（D7 含む）。',
 '【GATE】CTR の要員は STD 60分前にゲートへ移動（26W必要数ファイルの線表の CKIN→GA/GC と同じ）。D7 は2名がカウンターに残る（CTR STAY）。GATE 人数 = GI ＋ ゲートへ移る CTR ＋ PSPT CHK。',
 f'【インターバル】前日の終業から翌日の始業まで{INTERVAL//60}時間以上空けること（勤務表で守る条件。線表パターンには影響しない）。',
 '【人の動き】1日の人数が最少になるように作業をつなぐ（キャリアをまたいで担当）。全員が全キャリア・全ポジションを担当できる前提。実際の個人スキル（資格）は反映していない。',
 '【便】Operating Period シートの便・時間。RF の便名は CTS 線表に合わせ、16:00着が RF355、16:55発が RF356。HB は382/383（STD 16:15）で、CTR OPEN は CTS 線表と同じ STD 180分前。VZ（UTS委託）・WE（RAMP）は対象外。',
 '【未反映】CTS 線表にある「翌日DP」「FREE ASGN」「HELP」「GI OJT/INST」など日によって入る業務は、標準配置から外した。',
]
for i, t in enumerate(notes):
    c = ws.cell(4+i, 2, t); c.font = F10; c.alignment = LEFT; ws.row_dimensions[4+i].height = 34
ws.column_dimensions['A'].width = 2; ws.column_dimensions['B'].width = 150

# ---------- キャリア別配置 ----------
ws = wb.create_sheet('キャリア別配置')
ws['B2'] = 'キャリア・便ごとの配置人数（本線表で使用する標準形）と CTS 線表の実績'; ws['B2'].font = F10B
# gate staff in the 26W必要数 line charts (sheet 1): GC/GA/SC in the last hour
import jobs as oldjobs
OLDGATE = {}
for k, js in oldjobs.extract(OLDX).items():
    OLDGATE[k] = sum(1 for j in js if not j['arr'] and any(t['role'] in ('GC', 'GA', 'SC') for t in j['tasks']))
# CTS observed
wbc = openpyxl.load_workbook(CTSX)
obs = defaultdict(lambda: defaultdict(list)); nd = Counter()
for w in wbc:
    _, fl, st = cts.parse(w)
    cnt = defaultdict(Counter)
    for p in st:
        seen = set()
        for s, e, v in p['tasks']:
            k, role = cts.role_of(v)
            if k and (k, role) not in seen: seen.add((k, role)); cnt[k][role] += 1
    for k, c in cnt.items():
        nd[k] += 1
        for role, n in c.items(): obs[k][role].append(n)
CTSMAP = {'D7550': ('D7550', 'D7551'), 'D7552': ('D7552', 'D7553'), '7C1501': ('7C1501', '7C1502'), '7C1551': ('7C1551', '7C1552'),
          '7C1503': ('7C1503', '7C1504'), 'RF352': ('RF352', 'RF351'), 'RF356': ('RF355', 'RF356'), 'SL392': ('SL392', 'SL393'),
          'HB382': ('HB880', 'HB881'), '5J5128': ('5J5128', '5J5129')}
def obs_txt(code, role):
    v = obs.get(code, {}).get(role)
    if not v: return '-'
    m = Counter(v).most_common(1)[0][0]
    return f"{m}" + (f"（{min(v)}〜{max(v)}）" if min(v) != max(v) else '')
heads = ['キャリア', '便（着/発）', 'STA', 'STD', 'CTR OPEN', 'DC', 'GI', 'CTR', 'PSPT CHK', 'ARV', '計', 'GATE\n（STD60分前〜）', 'うちCTR\nSTAY', '必要数ファイル\nGATE', 'DC 勤務帯', 'AGT 勤務帯（GI/CTR）', 'ARV 勤務帯',
         'CTS実績 DC', 'CTS実績 GI', 'CTS実績 CTR', 'CTS実績 PSPT', 'CTS実績 ARV', 'CTS線表の便名']
head(ws, 4, 2, heads, [12, 16, 7, 7, 8, 5, 5, 5, 7, 5, 5, 11, 8, 12, 20, 22, 18, 9, 9, 9, 9, 9, 16], 40)
r = 5
for k in order_flights(set(FLIGHTS)):
    f = FLIGHTS[k]; cr = f['cr']; arr, dep = codes(k); d = DEP[cr]; a = ARR[cr]
    std, sta = tm(f['std']), tm(f['sta']); o, c_ = std + d['win'][0], std + d['win'][1]
    nctr, npspt, narv = d['ctr'], PSPT.get(k, 0), ARR_N.get(k, a['n'])
    pd, pa = PREPOST['DC'], PREPOST['CTR']
    stay = v3.CTR_STAY.get(cr, 0); gate = 1 + (nctr - stay) + npspt
    vals = [f"{cr} {CRNAME[cr]}", f"{arr}/{dep}" + ('（増便）' if k in NEWF else ''), f['sta'], f['std'], hc(o), 1, 1, nctr, npspt or '', narv, 2 + nctr + npspt + narv,
            gate, stay or '', OLDGATE.get(k, '-'),
            f"{hc(max(o-pd[0], v3.EARLIEST if cr=='D7' else 0))}〜{hc(c_+pd[1])}", f"{hc(max(o-pa[0], v3.EARLIEST if cr=='D7' else 0))}〜{hc(c_+pa[1])}（最低45分）", f"{hc(sta+a['win'][0])}〜{hc(sta+a['win'][1]+a['pw'])}"]
    if k in CTSMAP:
        ca, cd = CTSMAP[k]
        vals += [obs_txt(cd, 'DC'), obs_txt(cd, 'GI'), obs_txt(cd, 'CTR'), obs_txt(cd, 'PSPT CHK'), obs_txt(ca, 'ARV'), f"{ca}/{cd}（{nd.get(cd, 0)}日分）"]
    else:
        vals += ['-'] * 5 + ['（線表なし・類似便で配置）']
    for j, v in enumerate(vals):
        x = cell(ws, r, 2+j, v, al=CEN)
    ws.cell(r, 2).fill = fill(COLOR[cr]); ws.cell(r, 3).fill = fill(COLOR[cr])
    if k in NEWF:
        for j in range(1, 6): ws.cell(r, 2+j).fill = NEWFILL
    r += 1
r += 1
ws.cell(r, 2, 'CTS実績は 1DEC〜15DEC の線表で、その便がある日の配置人数（最頻値、カッコ内は最小〜最大）。HB は CTS 線表では HB880/881（STD 15:00）。').font = F8; r += 1
ws.cell(r, 2, '必要数ファイル GATE は「26W必要数」ファイルの線表（シート1）で STD 60分前からゲート（GC/GA/SC）に入る人数。7C1503/1504 はうち2名がゲート専任（STD 60分前〜）で、CTS 線表の PSPT CHK 2名と同じ時間帯。').font = F8; r += 2
ws.cell(r, 2, 'ポジションと必要スキル（業務内容は線表の表記からの推定です。社内の定義に合わせて修正してください）').font = F10B; r += 1
head(ws, r, 2, ['ポジション', '線表の表記', '主な業務（推定）', '必要スキル（推定）', '人数（1便あたり）', '準備', '後作業'])
ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8); ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=12)
r += 1
skills = [
 ('DC', 'DP → ○○ DC → PW', '出発便の統括（DCS 操作、旅客・搭載の締め、書類、GI・CTR への指示）', '各キャリアの DCS（出発管理システム）操作資格、DC 経験', '1', 'CTR OPEN 60分前', '出発後60分'),
 ('GI', 'AP → ○○ GI → PW', 'カウンター・ゲートの取りまとめ（搭乗ゲート担当）', '各キャリアのチェックイン・搭乗業務資格', '1', '最低45分（最大60分）', '最低45分（最大60分）'),
 ('CTR', 'CTR準備 → ○○ → PW', 'チェックインカウンター（発券・手荷物受託）', '各キャリアのチェックイン業務資格', 'D7 6 ／ 他 3', '最低45分（最大60分）', '最低45分（最大60分）'),
 ('PSPT CHK', '○○ PSPT CHK', '搭乗前の旅券確認（7C1504）', '旅券確認の知識', '2（7C1504）', '最低45分', '最低45分'),
 ('GATE', '（CTR から移動）', 'STD 60分前からの搭乗ゲート業務（CTR 要員が移動。D7 は2名がカウンターに残る）', '各キャリアの搭乗業務資格', 'GI＋CTR（D7は2名残り）＋PSPT', '-', '-'),
 ('ARV', '○○ ARV → PW', '到着旅客の案内・手荷物', '到着業務', '3（OPEN SPOT対応・全便）', 'STA 30〜40分前から', 'CTS 線表どおり'),
]
for row in skills:
    for j, v in enumerate(row):
        col = [2, 3, 4, 9, 13, 14, 15][j]
        x = cell(ws, r, col, v, al=LEFT if j in (2, 3) else CEN)
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8); ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=12)
    ws.row_dimensions[r].height = 30; r += 1
ws.column_dimensions['A'].width = 2; ws.freeze_panes = 'D5'

# ---------- 期間一覧 / 日別 ----------
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
        dd = [d for d in seg if d.isoweekday() == w]; x = cell(ix, r, 4+w, '')
        if dd:
            P = PI[day_pat[dd[0]]]; x.value = f"{P['id']} ({P['P']})"; x.hyperlink = f"#'{sname(P)}'!A1"; x.font = FLINK
    r += 1
ix.cell(r+1, 2, '・P01〜は「その日に運航する便の組み合わせ」の番号。組み合わせが同じ日は同じ線表。').font = F8
ix.column_dimensions['A'].width = 2

dy = wb.create_sheet('日別')
ALLK = order_flights(set(FLIGHTS))
head(dy, 2, 2, ['日付', '曜日', 'Period', '線表シート', '出面数', '労働時間'] + [f"{codes(k)[0]}/{codes(k)[1]}".replace('増便/増便', '増便') + ('\n' + FLIGHTS[k]['std'] if k in NEWF else '') for k in ALLK],
     [11, 5, 12, 11, 7, 8] + [9]*len(ALLK), 36)
for i, d in enumerate(days):
    F = day_pat[d]; P = PI[F]; rr = 3 + i
    per = next(f'{fd(a)}〜{fd(b)}' for a, b, _, _ in PAY if a <= d <= b)
    for j, v in enumerate([d, WJ[d.isoweekday()], per, sname(P), P['P'], round(P['H'], 1)]):
        cell(dy, rr, 2+j, v, fmt='yyyy/mm/dd' if j == 0 else None)
    dy.cell(rr, 5).hyperlink = f"#'{sname(P)}'!A1"; dy.cell(rr, 5).font = FLINK
    for j, k in enumerate(ALLK):
        x = cell(dy, rr, 8+j, '●' if k in F else '')
        if k in F: x.fill = fill(COLOR[FLIGHTS[k]['cr']])
dy.column_dimensions['A'].width = 2; dy.freeze_panes = 'D3'

# ---------- 線表 ----------
def bar(ws, r, s, e, text, rgb=None, font=F8):
    if e <= s: return
    c1, c2 = m2c(s), max(m2c(e) - 1, m2c(s))
    x = ws.cell(r, c1, text); x.font = font; x.alignment = CEN
    for cc in range(c1, c2+1):
        y = ws.cell(r, cc); y.border = B_ALL
        if rgb: y.fill = fill(rgb)
    if c2 > c1: ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
def hour_header(ws, r):
    for h in range(4, 24):
        c0 = m2c(h*60); x = ws.cell(r, c0, h); x.font = F8B; x.alignment = Alignment(horizontal='left', vertical='center')
        for cc in range(c0, c0+12): ws.cell(r, cc).fill = HDR
tval = lambda m: dt.time(m // 60, m % 60)
GRID_END = m2c(24*60) - 1
def segments(p):
    """expand each job into prep / main / post work, stretching AGT prep & post up to 60 min into free time"""
    segs = []
    jobs = p['jobs']
    for i, j in enumerate(jobs):
        lo = p['start'] if i == 0 else jobs[i-1]['e']
        hi = p['end'] if i == len(jobs) - 1 else jobs[i+1]['s']
        if p['rest']:
            if p['rest'][1] <= j['s']: lo = max(lo, p['rest'][1])
            if p['rest'][0] >= j['e']: hi = min(hi, p['rest'][0])
        stretch = j['pos'] in ('GI', 'CTR', 'PSPT CHK')
        if segs: lo = max(lo, segs[-1][1])
        if j['prep']:
            a0, a1, nm = j['prep']
            if stretch and lo < a0: a0 = max(lo, a1 - AGT_MAX)
            segs.append((a0, a1, nm, 'prep', j))
        m0, m1 = j['main']; segs.append((m0, m1, None, 'main', j))
        b0, b1 = j['pw']
        if stretch and hi > b1: b1 = min(hi, b0 + AGT_MAX)
        segs.append((b0, b1, 'PW', 'pw', j))
    return segs

for F, ds in plist:
    P = PI[F]; ws = wb.create_sheet(sname(P)); fl = order_flights(F)
    ws.sheet_view.zoomScale = 70
    for col, w in zip('ABCDEFGHIJKLMN', [2.3, 7.5, 4.5, 6, 6, 6.7, 7.9, 16, 9, 7, 7, 7, 8, 8]): ws.column_dimensions[col].width = w
    for c in range(15, GRID_END + 2): ws.column_dimensions[CL(c)].width = 1.6 if c >= X0 else 0.7
    ws['B1'] = f"CTS 26W 線表  {P['id']}（{'・'.join(WJ[w] for w in P['wds'])}）"; ws['B1'].font = F14
    rg = ', '.join(ddmon(a) + (f"–{ddmon(b)}" if b != a else '') + f"({WJ[a.isoweekday()]})" for a, b in ranges(ds))
    ws['B2'] = f"適用日（{len(ds)}日）: {rg}"; ws['B2'].font = F8; ws['B2'].alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=X0 + 12*10); ws.row_dimensions[2].height = max(15, 13 * (1 + len(rg) // 230))
    ws['B3'] = f"出面数 {P['P']} 名   延べ労働時間 {P['H']:.1f} h   平均 {P['H']/P['P']:.2f} h/人"; ws['B3'].font = F8B
    H = 5
    head(ws, H, 2, ['Carrier', '', 'Type', 'Dest', 'Freq', 'W/N', '便（着/発）', 'STA', 'STD', 'CTR OPEN', 'DC/GI/CTR', 'GATE', 'ARV'])
    hour_header(ws, H); ws.row_dimensions[H].height = 22
    FD = dict(FL); r = H + 1
    for k in fl:
        f = FLIGHTS[k]; cr = f['cr']; arr, dep = codes(k); dsk = [x for _, _, x in FD[k]]
        freq = 'Daily' if [1,2,3,4,5,6,7] in dsk else ' '.join(map(str, sorted({x for y in dsk for x in y})))
        o, c_ = tm(f['std']) + DEP[cr]['win'][0], tm(f['std']) + DEP[cr]['win'][1]
        vals = [cr, '', f['ac'], f['dest'], freq, f['wn'], f"{arr}/{dep}", f['sta'], f['std'], hc(o), f"1/1/{DEP[cr]['ctr']}", 1 + DEP[cr]['ctr'] - v3.CTR_STAY.get(cr, 0) + PSPT.get(k, 0), ARR_N.get(k, ARR[cr]['n'])]
        for j, v in enumerate(vals): cell(ws, r, 2+j, v, fl=NEWFILL if k in NEWF else None)
        ws.cell(r, 2).fill = fill(COLOR[cr])
        bar(ws, r, o, c_, f"STD {f['std']}  {dep}", COLOR[cr], F8B)
        sta = tm(f['sta']); a = ARR[cr]['win']
        bar(ws, r + 1, sta + a[0], sta + a[1], f"STA {f['sta']}  {arr}", COLOR[cr], F8)
        ws.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=2)
        for j in range(3, 15): ws.merge_cells(start_row=r, start_column=j, end_row=r+1, end_column=j)
        ws.row_dimensions[r].height = 16; ws.row_dimensions[r+1].height = 16
        r += 2
    r += 1; SH = r
    head(ws, SH, 2, ['Shift', 'No', '担当（順）', '', '', '', '', '開始', '終了', 'REST', '労働時間'])
    ws.merge_cells(start_row=SH, start_column=4, end_row=SH, end_column=8)
    hour_header(ws, SH); ws.row_dimensions[SH].height = 22
    r = SH + 1; first = r
    for i, p in enumerate(P['rows']):
        cell(ws, r, 2, p['code'], font=F8B); cell(ws, r, 3, i + 1)
        route = ' → '.join(f"{j['code']} {j['pos']}" for j in p['jobs'])
        x = ws.cell(r, 4, route); x.font = FG; x.alignment = LEFT
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)
        for col, m in ((9, p['start']), (10, p['end']), (11, p['brk'])): cell(ws, r, col, tval(m), fmt='hh:mm')
        cell(ws, r, 12, f'=J{r}-I{r}-K{r}', fmt='h:mm', fl=CALC)
        for s, e, nm, kind, j in segments(p):
            cr = j['cr']
            if kind == 'main':
                if j['pos'] == 'CTR' and j.get('gate'):
                    bar(ws, r, s, j['gate'], f"{j['code']} CTR", COLOR[cr], F8B)
                    bar(ws, r, j['gate'], e, f"{j['code']} GATE", COLOR[cr], F8B)
                else:
                    txt = j['code'] + (' CTR STAY' if j.get('stay') else ' CTR' if j['pos'] == 'CTR' else ' ' + j['pos'])
                    bar(ws, r, s, e, txt, COLOR[cr], F8B)
            elif kind == 'prep':
                bar(ws, r, s, e, nm, COLOR[cr] if nm in ('DP', 'AP') else None)
            else:
                bar(ws, r, s, e, nm)
        if p['rest']: bar(ws, r, *p['rest'], 'REST', 'FFFF0000', FWHITE)
        ws.row_dimensions[r].height = 22; r += 1
    last = r - 1
    for lab, fm, fmt in [('AVERAGE', f'=AVERAGE(L{first}:L{last})', 'h:mm'), ('総労働時間', f'=SUM(L{first}:L{last})', '[h]:mm'), ('稼働数', P['P'], None)]:
        ws.cell(r, 2, lab).font = F8B; x = ws.cell(r, 12, fm); x.font = F8B
        if fmt: x.number_format = fmt
        r += 1
    for rr in list(range(H, SH - 1)) + list(range(SH, last + 1)):
        for h in range(4, 24):
            x = ws.cell(rr, m2c(h*60)); b = x.border
            x.border = Border(left=hour, right=b.right, top=b.top, bottom=b.bottom)
    ws.freeze_panes = ws.cell(H + 1, X0)
    ws.sheet_properties.tabColor = 'FFFFD966' if F & NEWF else 'FFB4C6E7'

wb.save(OUT)
print('ok', len(plist))
