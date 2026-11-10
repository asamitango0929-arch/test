import math, random, copy, sys
import v3
from v3 import flight_jobs, is_naos, day_pat, days, q_down, q_up
from model import PERIODS
def rest_need(span, mode):
    if span <= 360: return 0
    if mode == '45' and span - 45 <= 480: return 45
    return 60
def build1(jobs, mode, rnd=None):
    people = []
    for j in sorted(jobs, key=lambda j: (j['s'], -j['e'], rnd.random() if rnd else 0)):
        best = None; bkey = None
        for p in people:
            g = j['s'] - p['last']
            if g < 0: continue
            span = q_up(j['e']) - q_down(p['first'])
            if span > v3.CAP: continue
            r = rest_need(span, mode)
            if r and max(p['gmax'], g) < r: continue
            key = (p['last'], rnd.random() if rnd else 0)
            if best is None or key > bkey: best, bkey = p, key
        if best is None: best = dict(first=j['s'], last=j['e'], jobs=[], gmax=0); people.append(best)
        else: best['gmax'] = max(best['gmax'], j['s'] - best['last'])
        best['jobs'].append(j); best['last'] = j['e']
    for p in people:
        p['span'] = q_up(p['last']) - q_down(p['first'])
    return people
def build(jobs, mode, tries=40):
    b = build1(jobs, mode)
    for i in range(tries):
        ps = build1(jobs, mode, random.Random(i))
        if len(ps) < len(b): b = ps
    return b
def run(excl, mode='60', hb150=False, part=False, peakhelp=False):
    if hb150: v3.DEP['HB']['win'] = (-150, 0)
    cache = {}; out = []
    for a, b, D, H in PERIODS:
        ds = [d for d in days if a <= d <= b]; cnt = []; shortc = []
        for d in ds:
            F = frozenset(day_pat[d] - excl)
            if F not in cache: cache[F] = build([j for k in F for j in flight_jobs(k) if not is_naos(j)], mode)
            ps = cache[F]
            sh = [p for p in ps if p['span'] <= 360] if part else []
            cnt.append(len(ps) - len(sh)); shortc.append(len(sh))
        P = sum(cnt); N = math.ceil(P / D) if peakhelp else max(math.ceil(P / D), max(cnt))
        out.append((N, max(shortc) if part else 0))
    v3.DEP['HB']['win'] = (-180, 0)
    return out
SC = {'全便自社': set(), 'SL委託': {'SL392'}, 'HB委託': {'HB382'}, 'SL・HB委託': {'SL392', 'HB382'}}
for name, ex in SC.items():
    base = run(ex)
    print(name, 'base', [n for n, _ in base], flush=True)
    for lab, kw in [('休憩45分', dict(mode='45')), ('HB CTR150', dict(hb150=True)), ('ピーク日応援', dict(peakhelp=True)),
                    ('短時間シフトをパート', dict(part=True)), ('全部', dict(mode='45', hb150=True, peakhelp=True, part=True))]:
        if 'HB' in lab and 'HB382' in ex: continue
        r = run(ex, **kw)
        print('  ', lab, [n for n, _ in r], ('パート/日 最大 ' + str(max(s for _, s in r))) if kw.get('part') else '', flush=True)
