import pickle, sys
from collections import Counter
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from analyze_sq import analyze, candidates
from skillq import load_people, SK
from skillq_train import PLAN, NEWBIE
OUT = sys.argv[1]
th = Side(style='thin', color='FF8C8C8C'); BOX = Border(left=th, right=th, top=th, bottom=th)
F9 = Font(name='Meiryo UI', size=9); F9B = Font(name='Meiryo UI', size=9, bold=True); F12 = Font(name='Meiryo UI', size=12, bold=True)
CEN = Alignment(horizontal='center', vertical='center', wrap_text=True); LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
fill = lambda c: PatternFill('solid', fgColor=c); HDR = fill('FFD9E1F2'); BAD = fill('FFFFC7CE'); GOOD = fill('FFC6EFCE'); WARN = fill('FFFFEB9C')
def put(ws, r, c, v, f=F9, al=CEN, fl=None):
    x = ws.cell(r, c, v); x.font = f; x.alignment = al; x.border = BOX
    if fl: x.fill = fl
    return x
def head(ws, r, c, hs):
    for j, h in enumerate(hs): put(ws, r, c + j, h, F9B, fl=HDR)
people = load_people('../xl5/in.xlsx')
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '結論'
ws['B2'] = '資格考慮パターンの検証（現有49名・SKILL 表ベース）'; ws['B2'].font = F12
notes = [
 '検証条件：拘束最大13時間（休憩1時間）、公休を所定から3日削減まで、連続勤務5日以内、2連休2回以上、インターバル11時間。● を資格あり、△ は訓練中（担当不可）として計算。',
 '必要資格：DC→DC、GI→GI、CKIN（→GATE へ移動）→CKIN＋GATE、CTR STAY→CKIN、PSPT CHK→GATE、ARV→ARR。シフトは49名の誰かが持つ資格の組み合わせでこなせるものだけで組んだ。',
 '結果：現状の資格では、HB委託・SL・HB委託とも12〜1月に毎日シフトが埋まらない（1日最大17名分）。最大の穴は D7（DC 1名・GI 1名しか資格者がいない）、次いで 7C DC、RF DC/GI。',
 '冬季キャリア（D7・7C・RF・5J・SL・HB）の資格を ● で持たない4名（ホンツォンミン、ジョン スンミン、本間菜桜、村田萌恵）は現状どのシフトにも入れず、実際に使えるのは45名。',
 '「研修計画」シートの研修（DC 9名分、D7 GI 3名、新人4名の D7・7C 基本資格、△ の完了）を行うと、SL・HB委託の1/16〜2/15 は不足が延べ29（1日最大3名）まで減る。12/16〜1/15 は延べ77（最大9名/日）残る。',
 'それでも埋まらない分は、公休のさらなる削減か応援が必要。12/20・12/27 など D7 2便と全便がそろう日に集中する。',
 '数字は最適化計算（各4分）の結果で、計算時間を延ばすと不足がさらに少し減る可能性がある。',
]
for i, t in enumerate(notes):
    x = ws.cell(4 + i, 2, f'・{t}'); x.font = Font(name='Meiryo UI', size=10); x.alignment = LEFT; ws.row_dimensions[4 + i].height = 34
ws.column_dimensions['A'].width = 2; ws.column_dimensions['B'].width = 150
r = 13
ws.cell(r, 2, '埋まらないシフト（延べ人日）').font = F9B; r += 1
head(ws, r, 2, ['パターン', '資格', 'Period', '必要シフト（延べ）', '埋まらない（延べ）', '1日の最大不足', '不足の多い資格（延べ人日）'])
for j, w in enumerate([150]): pass
r += 1
for lab, f in [('HB委託', 'sq_H_13_3_1.pkl'), ('HB委託', 'sqt_H.pkl'), ('SL・HB委託', 'sq_B_13_3_1.pkl'), ('SL・HB委託', 'sqt_B.pkl')]:
    _, res = analyze('../out/' + f)
    for x in res:
        put(ws, r, 2, lab); put(ws, r, 3, '研修後' if 'sqt' in f else '現状')
        put(ws, r, 4, x['a']); put(ws, r, 5, x['shifts']); put(ws, r, 6, x['uncovered'], fl=BAD if x['uncovered'] > 50 else WARN)
        put(ws, r, 7, x['max_unc_day']); put(ws, r, 8, '、'.join(f'{k} {v}' for k, v in list(x['lack'].items())[:8]), al=LEFT)
        r += 1
ws.column_dimensions['B'].width = 150
# skill counts
s2 = wb.create_sheet('資格者数')
s2['B2'] = 'キャリア別・資格別の人数（● 資格あり ／ △ 訓練中）'; s2['B2'].font = F12
head(s2, 4, 2, ['キャリア'] + SK)
c = Counter(); t = Counter()
for p in people:
    for s in p['skills']: c[s] += 1
    for s in p['train']: t[s] += 1
for i, cr in enumerate(['D7', '7C', 'RF', '5J', 'SL', 'HB', 'VZ', 'JD', 'HU']):
    put(s2, 5 + i, 2, cr, F9B)
    for j, s in enumerate(SK):
        v = c[(cr, s)]; put(s2, 5 + i, 3 + j, f"{v}" + (f"（+{t[(cr, s)]}）" if t[(cr, s)] else ''), fl=BAD if v <= 2 and cr in ('D7', '7C', 'RF', '5J', 'SL', 'HB') and s in ('DC', 'GI') else None)
for j, w in enumerate([10] + [11] * 6): s2.column_dimensions[CL(2 + j)].width = w
s2.cell(15, 2, '赤：冬季に毎日必要なのに資格者が2名以下。D7 は DC・GI とも1名で、毎日（12/14以降は1日2便）を1名では回せない。').font = F9
# training plan
s3 = wb.create_sheet('研修計画')
s3['B2'] = '超効率パターンに近づけるための研修計画（検証に使った案）'; s3['B2'].font = F12
head(s3, 4, 2, ['習得する資格', '対象者', '会社', '選んだ理由'])
r = 5
for (cr, s), names in PLAN.items():
    cand = {n: (co, why) for _, n, co, why in candidates(people, cr, s)}
    for n in names:
        co, why = cand.get(n, ('', ''))
        put(s3, r, 2, f'{cr} {s}', F9B); put(s3, r, 3, n); put(s3, r, 4, co); put(s3, r, 5, why, al=LEFT); r += 1
for n in NEWBIE:
    p = next(p for p in people if p['name'] == n)
    put(s3, r, 2, 'D7・7C CKIN/GATE/ARR', F9B); put(s3, r, 3, n); put(s3, r, 4, p['company'])
    put(s3, r, 5, '冬季キャリアの ● 資格なし（現状シフトに入れない）。訓練中：' + ('、'.join(sorted(f'{a} {b}' for a, b in p['train'])) or 'なし'), al=LEFT); r += 1
put(s3, r, 2, '△ の完了', F9B); put(s3, r, 3, '△ のある全員'); put(s3, r, 4, ''); put(s3, r, 5, '訓練中の資格をシーズン前に完了させる', al=LEFT); r += 2
s3.cell(r, 2, '資格ごとの候補者（優先順。△訓練中 → 同キャリアCKIN/GATE有資格 → 他社で同じ資格を持つ人）').font = F9B; r += 1
head(s3, r, 2, ['資格', '候補者', '会社', '理由']); r += 1
for cr in ('D7', '7C', 'RF', '5J', 'SL'):
    for s in ('DC', 'GI', 'CKIN'):
        for _, n, co, why in candidates(people, cr, s)[:5]:
            put(s3, r, 2, f'{cr} {s}'); put(s3, r, 3, n); put(s3, r, 4, co); put(s3, r, 5, why, al=LEFT); r += 1
for j, w in enumerate([22, 18, 12, 70]): s3.column_dimensions[CL(2 + j)].width = w
# rosters with names (trained, SL・HB委託)
for lab, f in [('SL・HB委託 研修後', 'sqt_B.pkl'), ('HB委託 研修後', 'sqt_H.pkl')]:
    D = pickle.load(open('../out/' + f, 'rb'))
    for o in D['out']:
        if 'R' not in o: continue
        nm = f"{'SLHB' if 'B' in f else 'HB'} 研修後 {o['a'].month}-{o['a'].day}"
        rs = wb.create_sheet(nm); ds = o['ds']
        rs['B1'] = f"{lab}  {o['a']:%m/%d}〜{o['b']:%m/%d}  実名 ROSTER（公休最大3日削減・拘束最大13h）"; rs['B1'].font = F9B
        head(rs, 3, 2, ['名前', '会社'] + [f"{d.month}/{d.day}" for d in ds] + ['勤務日数', '公休削減', '労働時間', '所定超過'])
        for i, rr in enumerate(sorted(o['R'], key=lambda x: -x['days'])):
            row = 4 + i; p = rr['person']
            put(rs, row, 2, p['name']); put(rs, row, 3, p['company'])
            for j, d in enumerate(ds):
                t_ = rr['sched'].get(d); put(rs, row, 4 + j, t_[0] if t_ else '公', fl=None if t_ else fill('FFE2EFDA'))
            put(rs, row, 4 + len(ds), rr['days']); put(rs, row, 5 + len(ds), max(0, rr['days'] - o['D']))
            hrs = rr['mins'] / 60; put(rs, row, 6 + len(ds), round(hrs, 1)); put(rs, row, 7 + len(ds), round(max(0, hrs - o['H']), 1), fl=BAD if hrs > o['H'] + 20 else None)
        row = 4 + len(o['R'])
        put(rs, row, 2, '埋まらない', F9B, fl=BAD); put(rs, row, 3, '')
        un = Counter()
        for (d, t_), n in o['U'].items(): un[d] += n
        for j, d in enumerate(ds): put(rs, row, 4 + j, un.get(d, '') or '', F9B, fl=BAD if un.get(d) else None)
        rs.column_dimensions['B'].width = 14; rs.column_dimensions['C'].width = 10
        for j in range(len(ds)): rs.column_dimensions[CL(4 + j)].width = 7
        rs.freeze_panes = 'D4'
wb.save(OUT); print('saved', wb.sheetnames)
