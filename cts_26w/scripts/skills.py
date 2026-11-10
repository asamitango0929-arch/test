"""Qualification-aware staffing and the cross-training path to the fully efficient pattern."""
import math, random, json, sys
from collections import Counter, defaultdict
from v3 import flight_jobs, is_naos, day_pat, days, q_down, q_up, CAP, REST
from model import PERIODS

DEPPOS = ('DC', 'GI', 'CTR', 'PSPT CHK')
def build_q(jobs, groups, rnd=None):
    """greedy shift building; departure jobs of one shift must all be covered by one skill group (ARV free)"""
    people = []
    for j in sorted(jobs, key=lambda j: (j['s'], -j['e'], rnd.random() if rnd else 0)):
        dep = j['pos'] in DEPPOS
        best = None; bkey = None
        for p in people:
            g = j['s'] - p['last']
            if g < 0: continue
            span = q_up(j['e']) - q_down(p['first'])
            if span > CAP: continue
            if span > 360 and max(p['gmax'], g) < REST: continue
            if dep and groups is not None:
                crs = p['crs'] | {j['cr']}
                if not any(crs <= G for G in groups): continue
            key = (p['last'], rnd.random() if rnd else 0)
            if best is None or key > bkey: best, bkey = p, key
        if best is None:
            best = dict(first=j['s'], last=j['e'], jobs=[], gmax=0, crs=set()); people.append(best)
        else:
            best['gmax'] = max(best['gmax'], j['s'] - best['last'])
        best['jobs'].append(j); best['last'] = j['e']
        if dep: best['crs'].add(j['cr'])
    return people
def best_q(jobs, groups, tries=30):
    best = build_q(jobs, groups)
    for i in range(tries):
        ps = build_q(jobs, groups, random.Random(i))
        if len(ps) < len(best): best = ps
    return best

SC = {'全便自社': set(), 'SL委託': {'SL392'}, 'HB委託': {'HB382'}, 'SL・HB委託': {'SL392', 'HB382'}}
_c = {}
def day_shifts(F, groups):
    key = (F, None if groups is None else frozenset(groups))
    if key not in _c:
        _c[key] = best_q([j for k in F for j in flight_jobs(k) if not is_naos(j)], groups)
    return _c[key]
def evaluate(excl, groups):
    out = []
    for a, b, D, H in PERIODS:
        ds = [d for d in days if a <= d <= b]
        cnt = [len(day_shifts(frozenset(day_pat[d] - excl), groups)) for d in ds]
        P = sum(cnt); N = max(math.ceil(P / D), max(cnt))
        out.append(dict(P=P, N=N, peak=max(cnt)))
    return out
def carriers(excl): return sorted({k[:2] if k[:2] != '7C' else '7C' for d in days for k in day_pat[d] - excl})

if __name__ == '__main__':
    from ref import FLIGHTS
    res = {}
    for name, excl in SC.items():
        crs = sorted({FLIGHTS[k]['cr'] for d in days for k in day_pat[d] - excl})
        eff = evaluate(excl, None)
        ded = evaluate(excl, [frozenset([c]) for c in crs])
        # greedy cross-training path: merge two skill groups at a time
        groups = [frozenset([c]) for c in crs]; path = []
        cur = sum(r['N'] for r in ded)
        while len(groups) > 1:
            cands = []
            for i in range(len(groups)):
                for k in range(i + 1, len(groups)):
                    g2 = [g for t, g in enumerate(groups) if t not in (i, k)] + [groups[i] | groups[k]]
                    ev = evaluate(excl, g2); cands.append((sum(r['N'] for r in ev), [max(r['N'] for r in ev)], i, k, g2, ev))
            cands.sort(key=lambda c: (c[0], len(c[4][-1])))
            tot, _, i, k, g2, ev = cands[0]
            path.append(dict(merge=[sorted(groups[i]), sorted(groups[k])], groups=[sorted(g) for g in g2], N=[r['N'] for r in ev], total=tot, gain=cur - tot))
            cur = tot; groups = g2
        res[name] = dict(carriers=crs, efficient=[r['N'] for r in eff], dedicated=[r['N'] for r in ded], peak_eff=[r['peak'] for r in eff], peak_ded=[r['peak'] for r in ded], path=path)
        print(name, 'eff', res[name]['efficient'], 'ded', res[name]['dedicated'], flush=True)
        for st in path: print('   +', st['merge'], st['N'], 'gain', st['gain'], flush=True)
    json.dump(res, open(sys.argv[1], 'w'), ensure_ascii=False, indent=1)
