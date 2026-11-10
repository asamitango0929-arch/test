import sys, pickle
from collections import Counter, defaultdict
from datetime import timedelta as td
OUT = sys.argv[1]
src = open('gen5.py').read().split('\nwb = openpyxl.Workbook()')[0]
src = src.replace("""        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=10)
        for col, m in (('K', p['start'])""", """        if not isn and p.get('who'):
            x = ws.cell(r, 6, p['who']); x.font = Font(name='Meiryo UI', size=8, bold=p['who'].startswith('応援'), color='FFC55A11' if p['who'].startswith('応援') else 'FF000000'); x.alignment = CEN
            if p['who'].startswith('応援'):
                for c in range(6, 11): ws.cell(r, c).fill = fill('FFFCE4D6')
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=10)
        for col, m in (('K', p['start'])""")
sys.argv = ['gen5', OUT, '../out/solved_B.pkl', 'B']
G = {}; exec(src, G)
import export
export.AGT_MAX = 45
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter as CL
fill = G['fill']; BOX = G['BOX']; CEN = G['CEN']; LEFT = G['LEFT']; HDRF = G['HDRF']; CALC = G['CALC']; NAOSF = G['NAOSF']; BAD = G['BAD']; OFF = G['OFF']
M9 = G['M9']; M9B = G['M9B']; WJ = G['WJ']; md = G['md']
SUPF = fill('FFFCE4D6')
def segs_of(p):
    s = export.segments(dict(p, rest=None))
    for r0, r1 in p['rests']: s.append([r0, r1, 'REST', 'rest', None])
    return sorted(s)
G['shift_segments'] = segs_of
import os
def _ld(k):
    f = f'../out/p5b_{k}.pkl'
    a = pickle.load(open(f'../out/p5_{k}.pkl', 'rb'))
    if os.path.exists(f):
        b = pickle.load(open(f, 'rb'))
        sa = sum(sum(1 for w in a['res'][d] if w[0] == 'S') for d in a['ds']); sb = sum(sum(1 for w in b['res'][d] if w[0] == 'S') for d in b['ds'])
        if sb < sa: return b
    return a
P5 = [_ld(k) for k in (1, 2, 3)]
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '結論'
ws['B2'] = 'CTS 26W 12/16〜3/15 線表・ROSTER（SL・HB 委託／11月末在籍49名＋研修計画／NAOS 7名 15:00〜19:00）'; ws['B2'].font = Font(name='Meiryo UI', size=13, bold=True)
notes = [
 '自社は D7・7C・RF・5J。HB・SL は委託、VZ はオーロラ委託。NAOS 7名は毎日 15:00〜19:00（7C増便の GI・CKIN 3・ARV 3。増便のない日は時間内の到着業務など）。',
 '要員：11月末在籍49名。資格は SKILL 表＋研修計画（D7 DC/GI、7C・5J・RF DC、RF GI、資格のない13名に D7・7C の CKIN/GATE/ARR、△ の完了）。',
 'シフト：拘束最大12時間・最小4時間・15分単位。拘束6時間超は休憩60分（作業と作業の間）。DC 前後60分、AGT 前後45分、D7551 は 5:30／6:30 出勤。',
 '組み方：先に各人の出勤日と公休を決め（所定日数〜＋1日、連続勤務5日以内、2連休2回以上）、日ごとにその日出勤する人の資格に合わせて作業を直接割り当てた（長い勤務と短い勤務の組み合わせ）。前日の終業から11時間以上空けている。',
 '最適化の優先順位：①全作業に必ず担当者を付ける、②応援の人数を最小にする、③自社の拘束時間を短くする（非稼働時間を減らして月間労働時間の超過を抑える）。',
 '49名で担当できない作業は「応援」として明示（必要な時間と業務をセルに表示）。各線表シートは代表日の実際の担当者入り。',
]
for i, t in enumerate(notes):
    x = ws.cell(4 + i, 2, f'・{t}'); x.font = Font(name='Meiryo UI', size=10); x.alignment = LEFT; ws.row_dimensions[4 + i].height = 32
ws.column_dimensions['A'].width = 2; ws.column_dimensions['B'].width = 140
r = 11
for j, h in enumerate(['Period', '必要人日（自社＋応援）', '自社49名', '応援（延べ）', '応援 1日最大', '応援 1日平均', '自社 月間労働時間（最小〜最大）', '所定時間', '所定超過の人数', '公休を1日削った人']):
    x = ws.cell(r, 2 + j, h); x.font = M9B; x.fill = HDRF; x.border = BOX; x.alignment = CEN
ws.row_dimensions[r].height = 30
order = ['結論']
for P in P5:
    ds = P['ds']; res = P['res']; people = P['people']; tag = f"{P['a'].month}-{P['a'].day}"
    own = defaultdict(dict); sup = defaultdict(list)
    for d in ds:
        for w, s in res[d].items():
            if w[0] == 'P': own[w[1]][d] = s
            else: sup[d].append(s)
    for d in ds: sup[d].sort(key=lambda s: (s['start'], s['end']))
    nsup = max((len(v) for v in sup.values()), default=0); totsup = sum(len(v) for v in sup.values())
    hrs = {i: sum(s['work'] for s in own[i].values()) / 60 for i in range(len(people))}
    r += 1
    vals = [f"{md(P['a'])}〜{md(P['b'])}", sum(len(res[d]) for d in ds), sum(len(v) for v in own.values()), totsup, nsup, round(totsup / len(ds), 1),
            f"{min(hrs.values()):.0f}〜{max(hrs.values()):.0f}h", P['H'], sum(1 for h in hrs.values() if h > P['H']), sum(1 for i in own if len(own[i]) > P['D'])]
    for j, v in enumerate(vals):
        x = ws.cell(r, 2 + j, v); x.font = M9; x.border = BOX; x.alignment = CEN
    ws.cell(r, 5).fill = BAD if totsup else fill('FFC6EFCE')
    # roster
    rs = wb.create_sheet(f"ROSTER {tag}"); order.append(rs.title)
    rs['B1'] = f"ROSTER {md(P['a'])}〜{md(P['b'])}　自社49名＋応援{nsup}枠＋NAOS 7名　所定 {P['D']}日・{P['H']}h"; rs['B1'].font = M9B
    hdr = ['名前', '会社'] + [f"{d.month}/{d.day}\n{WJ[d.isoweekday()]}" for d in ds] + ['勤務日数', '公休', '公休削減', '労働時間', '所定との差', '2連休', '最大連続']
    for j, h in enumerate(hdr):
        x = rs.cell(3, 2 + j, h); x.font = M9B; x.fill = HDRF; x.border = BOX; x.alignment = CEN
    rs.row_dimensions[3].height = 30; rs.column_dimensions['A'].width = 2; rs.column_dimensions['B'].width = 15; rs.column_dimensions['C'].width = 9
    for j in range(len(ds)): rs.column_dimensions[CL(4 + j)].width = 17
    for j in range(7): rs.column_dimensions[CL(4 + len(ds) + j)].width = 9
    def put(rr_, c_, s, f=None):
        x = rs.cell(rr_, c_, f"{s['code']}\n{s['label']}" if s else '公'); x.border = BOX; x.alignment = CEN
        x.font = Font(name='Arial', size=7, bold=True) if s else Font(name='Meiryo UI', size=8, color='FF548235')
        if not s: x.fill = OFF
        elif f: x.fill = f
    row = 4
    for i in sorted(range(len(people)), key=lambda i: (-len(own[i]), people[i]['name'])):
        p = people[i]; a_ = own[i]
        x = rs.cell(row, 2, p['name']); x.font = M9; x.border = BOX; x.alignment = CEN
        x = rs.cell(row, 3, p.get('company') or ''); x.font = M9; x.border = BOX; x.alignment = CEN
        for j, d in enumerate(ds): put(row, 4 + j, a_.get(d))
        dd = sorted(a_)
        for d0, d1 in zip(dd, dd[1:]):
            if d1 == d0 + td(1): assert (1440 - a_[d0]['end']) + a_[d1]['start'] >= 660, (p['name'], d0)
        pairs = 0; k_ = 0
        while k_ < len(ds) - 1:
            if ds[k_] not in a_ and ds[k_ + 1] not in a_: pairs += 1; k_ += 2
            else: k_ += 1
        mc = c = 0
        for d in ds: c = c + 1 if d in a_ else 0; mc = max(mc, c)
        h = hrs[i]
        vals = [len(a_), len(ds) - len(a_), max(0, len(a_) - P['D']), round(h, 1), round(h - P['H'], 1), pairs, mc]
        for j, v in enumerate(vals):
            x = rs.cell(row, 4 + len(ds) + j, v); x.font = M9; x.border = BOX; x.alignment = CEN
            if (j == 2 and v) or (j == 4 and v > 0): x.fill = BAD
        rs.row_dimensions[row].height = 34; row += 1
    for k in range(nsup):
        x = rs.cell(row, 2, f'応援{k+1}'); x.font = Font(name='Meiryo UI', size=9, bold=True, color='FFC55A11'); x.border = BOX; x.alignment = CEN; x.fill = SUPF
        x = rs.cell(row, 3, '外部'); x.font = M9; x.border = BOX; x.fill = SUPF
        for j, d in enumerate(ds):
            s = sup[d][k] if k < len(sup[d]) else None
            if s: put(row, 4 + j, s, SUPF)
            else: x = rs.cell(row, 4 + j, ''); x.border = BOX
        rs.row_dimensions[row].height = 34; row += 1
    for k in range(7):
        x = rs.cell(row, 2, f'NAOS{k+1}'); x.font = Font(name='Arial', size=8, bold=True, color='FF5B3B8C'); x.border = BOX; x.fill = NAOSF; x.alignment = CEN
        x = rs.cell(row, 3, '派遣'); x.font = M9; x.border = BOX; x.fill = NAOSF
        for j, d in enumerate(ds):
            n = P['naos'][d][k]
            lab = ' → '.join(f"{(G['codes'](jb['k'])[0] if jb['pos'] == 'ARV' else G['codes'](jb['k'])[1])} {'CKIN' if jb['pos'] == 'CTR' else jb['pos']}" for jb in n['jobs']) or '待機'
            x = rs.cell(row, 4 + j, f"{n['code']}\n{lab}"); x.font = Font(name='Arial', size=7, bold=True); x.border = BOX; x.alignment = CEN; x.fill = NAOSF
        rs.row_dimensions[row].height = 30; row += 1
    for lab, fn in (('自社出勤', lambda d: sum(1 for w in res[d] if w[0] == 'P')), ('応援', lambda d: len(sup[d]) or ''), ('NAOS', lambda d: 7)):
        x = rs.cell(row, 2, lab); x.font = M9B; x.fill = CALC; x.border = BOX
        for j, d in enumerate(ds):
            x = rs.cell(row, 4 + j, fn(d)); x.font = M9B; x.fill = BAD if lab == '応援' and sup[d] else CALC; x.border = BOX; x.alignment = CEN
        row += 1
    rs.freeze_panes = 'D4'
    # line charts: representative (first) date of each flight pattern, with names
    groups = {}
    for d in ds: groups.setdefault(P['F'][d], []).append(d)
    for F, gds in sorted(groups.items(), key=lambda x: x[1][0]):
        d0 = gds[0]; nm = d0.strftime('%d%b(%a)').upper(); order.append(nm)
        wds = sorted({d.isoweekday() for d in gds})
        shifts = []
        for w, s in res[d0].items():
            s = dict(s); s['who'] = people[w[1]]['name'] if w[0] == 'P' else f'応援{w[1]+1}'; shifts.append(s)
        shifts.sort(key=lambda s: (s['start'], s['end']))
        naos = [dict(n) for n in P['naos'][d0]]
        G['line_sheet'](wb, nm, f"CTS 26W 線表  {nm}  （{md(P['a'])}〜{md(P['b'])}・{'・'.join(WJ[w] for w in wds)}曜の便パターン）　SL・HB 委託",
                        f"この日の実際の担当者。同じ便パターンの日（{len(gds)}日）: {', '.join(md(d) for d in gds)}　　自社 {sum(1 for w in res[d0] if w[0] == 'P')} 名＋応援 {len(sup[d0])} 名＋NAOS 7 名",
                        F, shifts, gds, naos)
wb._sheets = [wb[n] for n in order]
wb.save(OUT); print('saved', len(order))
