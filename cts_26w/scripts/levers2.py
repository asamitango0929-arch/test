import sys, random
from collections import Counter
import plan4, v3
from plan4 import naos_split, build, finish, CAP
from skillq import load_active, PERIODS, req
from sq2 import train2
from v3 import day_pat, days, tm, FLIGHTS
plan4.need_rest = lambda span: 0 if span <= 360 else 60
plan4.rest_ok = lambda gaps, span: plan4.need_rest(span) == 0 or any(g >= 60 for g in gaps)
EXCL = {'HB382', 'SL392'}
def quick(k):
    f = FLIGHTS[k]; return tm(f['std']) - tm(f['sta']) <= 80
def jobs_variant(F, arv_dep):
    own, naos = naos_split(F)
    if not arv_dep: return own
    out = []; cnt = Counter()
    for j in own:
        if j['pos'] == 'ARV' and quick(j['k']):
            cnt[j['k']] += 1
            if cnt[j['k']] > 1: continue          # 2 of 3 arrival agents = departure gate team of the same aircraft
        out.append(j)
    return out
def run(ds, people, arv_all, arv_dep, tries=30):
    if arv_all:
        for p in people: p['skills'] = p['skills'] | {(c, 'ARR') for c in ('D7', '7C', 'RF', '5J')}
    sk = [p['skills'] for p in people]; cache = {}; tot = 0; peak = 0
    for d in ds:
        F = frozenset(day_pat[d] - EXCL)
        if F not in cache:
            js = jobs_variant(F, arv_dep); best = build(js, sk)
            for i in range(tries):
                ps = build(js, sk, random.Random(i))
                if len(ps) < len(best): best = ps
            cache[F] = len(best)
        tot += cache[F]; peak = max(peak, cache[F])
    return tot, peak
if __name__ == '__main__':
    for k in (1, 2, 3):
        a, b, D, H = PERIODS[k]; ds = [d for d in days if a <= d <= b]
        cap = 49 * (D + 1)
        for lab, aa, ad in [('現状', False, False), ('到着チーム横断', True, False), ('出発が到着兼任', False, True), ('両方', True, True)]:
            tot, peak = run(ds, train2(load_active('../xl5/in.xlsx')), aa, ad)
            print(a, lab, '必要人日', tot, '上限', cap, '差', cap - tot, 'peak', peak, flush=True)
