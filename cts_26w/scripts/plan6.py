"""No support: departure gate team covers 2 of 3 arrival agents on quick turns; no days-off cut if possible."""
import sys, pickle, time, random, os
from collections import Counter
import plan5, plan4
from plan5 import dayoff_plan, assign_day, SCARCE, EXCL
from skillq import load_active, req, PERIODS
from sq2 import train2
from v3 import day_pat, days
import levers2
from levers2 import jobs_variant
k = int(sys.argv[1]); EXTRA = int(sys.argv[2]); TL = int(sys.argv[3]) if len(sys.argv) > 3 else 30
a, b, D, H = PERIODS[k]
people = train2(load_active('../xl5/in.xlsx')); sk = [p['skills'] for p in people]
ds = [d for d in days if a <= d <= b]
jobs = {}; naos = {}; cache = {}; need = {}
for d in ds:
    F = frozenset(day_pat[d] - EXCL)
    _, naos[d] = plan4.naos_split(F)
    jobs[d] = jobs_variant(F, True)
    if F not in cache:
        best = plan4.build(jobs[d], sk)
        for i in range(30):
            ps = plan4.build(jobs[d], sk, random.Random(i))
            if len(ps) < len(best): best = ps
        cache[F] = len(best)
    need[d] = cache[F] + 1
prevf = os.environ.get('PREV', f'../out/p6_{k}_0.pkl')
if os.path.exists(prevf):
    pr = pickle.load(open(prevf, 'rb')); need = {d: len(pr['res'][d]) + int(os.environ.get('NEEDPLUS', '0')) for d in ds}
skneed = {d: Counter(s for j in jobs[d] for s in req(j) if s in SCARCE) for d in ds}
src = plan5.dayoff_plan
def dayoff(people, ds, need, skneed, D_, tlimit=90):
    return src(people, ds, need, skneed, D_ - 1 + EXTRA, tlimit)
plan = dayoff(people, ds, need, skneed, D)
print('dayoff', sum(len(v) for v in plan.values()), 'need', sum(need.values()), flush=True)
res = {}; prev_end = {}
for d in ds:
    earliest = {p: max(0, prev_end[p] - 1440 + 660) for p in plan[d] if p in prev_end}
    r, stt = assign_day(jobs[d], people, plan[d], earliest, TL)
    res[d] = r; prev_end = {w[1]: s['end'] for w, s in r.items() if w[0] == 'P'}
    print(d, 'avail', len(plan[d]), 'own', sum(1 for w in r if w[0] == 'P'), 'support', sum(1 for w in r if w[0] == 'S'), stt, flush=True)
pickle.dump(dict(people=people, a=a, b=b, D=D, H=H, ds=ds, plan=plan, res=res, naos=naos, F={d: frozenset(day_pat[d] - EXCL) for d in ds}), open(os.environ.get('OUTF', f'../out/p6_{k}_{EXTRA}.pkl'), 'wb'))
print('saved', sum(sum(1 for w in res[d] if w[0] == 'S') for d in ds))
