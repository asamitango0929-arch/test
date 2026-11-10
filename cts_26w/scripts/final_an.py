import math, json
from collections import Counter
import v3
from skills import best_q, SC, DEPPOS
from v3 import flight_jobs, is_naos, day_pat, days
from model import PERIODS
from ref import FLIGHTS
res = json.load(open('../out/skills.json'))
out = {}
for name, ex in SC.items():
    path = res[name]['path']
    # skill groups after the first steps that reach the efficient level
    eff = res[name]['efficient']; steps = []
    for st in path:
        steps.append(st)
        if st['N'] == eff: break
    groups = [frozenset(g) for g in steps[-1]['groups']] if steps else None
    a, b, D, H = PERIODS[2]                      # 1/16-2/15 (peak period)
    ds = [d for d in days if a <= d <= b]
    need = Counter()
    for d in ds:
        F = frozenset(day_pat[d] - ex)
        ps = best_q([j for k in F for j in flight_jobs(k) if not is_naos(j)], groups)
        c = Counter()
        for p in ps:
            crs = sorted({j['cr'] for j in p['jobs'] if j['pos'] in DEPPOS})
            if len(crs) > 1:
                dc = [j['cr'] for j in p['jobs'] if j['pos'] == 'DC']
                c[('+'.join(crs), 'DC含む' if dc else 'AGT')] += 1
        for k2, v in c.items(): need[k2] = max(need[k2], v)
    L = len(ds)
    out[name] = dict(steps=[dict(merge=s['merge'], N=s['N'], gain=s['gain']) for s in steps],
                     cross={f"{k[0]}（{k[1]}）": dict(per_day=v, persons=math.ceil(v * L / D)) for k, v in sorted(need.items())})
    print(name, out[name])
# CTR OPEN 150 for D7/HB/5J
for cr in ('D7', 'HB', '5J'): v3.DEP[cr]['win'] = (-150, v3.DEP[cr]['win'][1])
from skills import evaluate, _c
_c.clear()
for name, ex in SC.items():
    r = evaluate(ex, None); out[name]['ctr150'] = [x['N'] for x in r]; print(name, 'CTR150', out[name]['ctr150'])
json.dump(out, open('../out/final_an.json', 'w'), ensure_ascii=False, indent=1)
