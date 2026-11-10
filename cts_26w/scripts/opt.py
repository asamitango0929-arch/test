import math
from datetime import date, timedelta as td
from jobs import extract, hm
from periods import S, E, ops
PRE, PW, GAP = 30, 30, 15
PAY = [  # (start, end, period days, std days, std hours)  ① 正社員 2026年Ver（2027分は2026年の同期間の値）
 (date(2026,10,16), date(2026,11,15), 21, 168),
 (date(2026,11,16), date(2026,12,15), 20, 160),
 (date(2026,12,16), date(2027,1,15), 22, 176),
 (date(2027,1,16), date(2027,2,15), 21, 168),
 (date(2027,2,16), date(2027,3,15), 19, 152),
 (date(2027,3,16), date(2027,4,15), 21, 168),
]
def brk(span): return 0 if span <= 360 else (45 if span <= 525 else 60)
def build1(jobs, cap, maxgap, rnd=None):
    """greedy best-fit: fewest people, each shift = PRE + jobs + PW, span <= cap, idle gap <= maxgap"""
    people = []
    order = sorted(jobs, key=lambda j: (j['s'], -j['e'], rnd.random() if rnd else 0))
    for j in order:
        best = None; bkey = None
        for p in people:
            g = j['s'] - p['last']
            if g < GAP or g > maxgap: continue
            span = (j['e'] + PW) - (p['first'] - PRE)
            if span > cap: continue
            if brk(span) and max(p['gmax'], g) < brk(span) + GAP: continue
            key = (p['last'], rnd.random() if rnd else 0)
            if best is None or key > bkey: best, bkey = p, key
        if best is None:
            best = dict(first=j['s'], last=j['e'], jobs=[], gmax=0); people.append(best)
        else: best['gmax'] = max(best['gmax'], j['s'] - best['last'])
        best['jobs'].append(j); best['last'] = j['e']
    for p in people:
        p['start'] = p['first'] - PRE; p['end'] = p['last'] + PW
        span = p['end'] - p['start']; p['brk'] = brk(span); p['work'] = span - p['brk']
        gaps = [(a['e'], b['s']) for a, b in zip(p['jobs'], p['jobs'][1:])]
        big = [g for g in gaps if g[1] - g[0] >= p['brk'] + GAP] if p['brk'] else []
        p['rest'] = None
        if p['brk'] and big:
            a, b = max(big, key=lambda g: g[1]-g[0]); m = a + max(0, ((b - a - p['brk']) // 2 // 5) * 5)
            p['rest'] = (m, m + p['brk'])
        p['warn'] = bool(p['brk'] and not p['rest'])
    people.sort(key=lambda p: (p['start'], p['end']))
    return people
import random
def build(jobs, cap, maxgap, tries=60):
    best = build1(jobs, cap, maxgap)
    sc = lambda ps: (len(ps), sum(p['warn'] for p in ps), sum(p['work'] for p in ps))
    for i in range(tries):
        ps = build1(jobs, cap, maxgap, random.Random(i))
        if sc(ps) < sc(best): best = ps
    return best
J = extract('../xl/in.xlsx')
def jobs_for(F): return [j for k in F for j in J[k]]
days = [S + td(i) for i in range((E - S).days + 1)]
day_pat = {d: frozenset(ops(d)) for d in days}
CAPS = [13*60]
GAPS = [600]
cache = {}
def stat(F, cap, mg):
    key = (F, cap, mg)
    if key not in cache:
        ps = build(jobs_for(F), cap, mg); cache[key] = (len(ps), sum(p['work'] for p in ps), sum(p['warn'] for p in ps))
    return cache[key]
def period_eval(a, b, sd, sh, cap, mg):
    ds = [d for d in days if a <= d <= b]
    frac = len(ds) / ((b - a).days + 1)
    P = sum(stat(day_pat[d], cap, mg)[0] for d in ds); H = sum(stat(day_pat[d], cap, mg)[1] for d in ds) / 60
    need_d, need_h = P / (sd * frac), H / (sh * frac)
    return dict(days=len(ds), frac=frac, P=P, H=H, need_d=need_d, need_h=need_h, N=math.ceil(max(need_d, need_h) - 1e-9))
if __name__ == '__main__':
    for a, b, sd, sh in PAY:
        res = sorted(((period_eval(a, b, sd, sh, c, g)['N'], max(period_eval(a,b,sd,sh,c,g)['need_d'],period_eval(a,b,sd,sh,c,g)['need_h']), c, g) for c in CAPS for g in GAPS))[:4]
        print(a, b, [(n, round(x,1), c//60, g) for n, x, c, g in res])
    for c in CAPS:
        for g in GAPS:
            Ns = [period_eval(a, b, sd, sh, c, g)['N'] for a, b, sd, sh in PAY]
            print('cap', c//60, 'gap', g, Ns, 'max', max(Ns))
