"""Rest: span 6-8h -> 60 min, 8-12h -> 120 min (one block or two). NAOS 7 people every day 15:00-19:00."""
import sys, random, pickle, time, copy
from collections import Counter
import v3
from v3 import flight_jobs, day_pat, days, q_down, q_up, code_of
from skillq import load_active, req, roster_skill, PERIODS
from sq2 import train2
import plan3
from plan3 import finish as finish3
CAP = 12 * 60; NA0, NA1 = 15 * 60, 19 * 60
def need_rest(span): return 0 if span <= 360 else (60 if span <= 480 else 120)
def rest_ok(gaps, span):
    r = need_rest(span)
    if r == 0: return True
    if r == 60: return any(g >= 60 for g in gaps)
    return any(g >= 120 for g in gaps) or sum(1 for g in gaps if g >= 60) >= 2
def naos_split(F):
    """NAOS: 7 fixed people 15:00-19:00. First the 7C extra flight (GI, CKIN x3, ARV x3, cut to 19:00), then other jobs that fit."""
    jobs = [j for k in F for j in flight_jobs(k)]
    rows = [[] for _ in range(7)]; taken = set()
    def fits(j): return j['pos'] != 'DC' and j['s'] >= NA0 and min(j['e'], NA1) - 0 >= j['main'][1] and j['main'][1] <= NA1
    cand = [j for j in jobs if j['k'] == '7CXXX' and j['pos'] in ('GI', 'CTR', 'ARV')] + sorted([j for j in jobs if j['k'] != '7CXXX' and j['pos'] in ('ARV', 'GI', 'CTR', 'PSPT CHK')], key=lambda j: j['s'])
    for j in cand:
        if id(j) in taken or not fits(j): continue
        jj = dict(j); jj['e'] = min(j['e'], NA1); jj['pw'] = (j['pw'][0], min(j['pw'][1], NA1))
        for row in rows:
            if all(jj['s'] >= o['e'] or jj['e'] <= o['s'] for o in row):
                row.append(jj); taken.add(id(j)); break
    naos = [dict(first=NA0, last=NA1, jobs=sorted(r, key=lambda j: j['s']), start=NA0, end=NA1, brk=0, work=NA1 - NA0, rest=None, rests=[], code=code_of(NA0, NA1), naos=True) for r in rows]
    own = [j for j in jobs if id(j) not in taken]
    return own, naos
def build(jobs, skillsets, rnd=None, pref_long=True):
    people = []
    for j in sorted(jobs, key=lambda j: (j['s'], -j['e'], rnd.random() if rnd else 0)):
        rq = req(j); best = None; bkey = None
        for p in people:
            g = j['s'] - p['last']
            if g < 0: continue
            span = q_up(j['e']) - q_down(p['first'])
            if span > CAP or not rest_ok(p['gaps'] + [g], span): continue
            need = p['req'] | rq
            nq = sum(1 for S in skillsets if need <= S)
            if nq == 0: continue
            key = (nq >= 3, p['last'], rnd.random() if rnd else 0)
            if best is None or key > bkey: best, bkey = p, key
        if best is None:
            best = dict(first=j['s'], last=j['e'], jobs=[], gaps=[], req=set()); people.append(best)
        else: best['gaps'].append(j['s'] - best['last'])
        best['jobs'].append(j); best['last'] = max(best['last'], j['e']); best['req'] |= rq
    return people
def finish(p):
    p = finish3(p)
    span = p['end'] - p['start']; r = need_rest(span)
    gl = sorted([(b['s'] - a['e'], a['e'], b['s']) for a, b in zip(p['jobs'], p['jobs'][1:]) if b['s'] - a['e'] >= 60], reverse=True)
    rests = []
    if r == 60 or (r == 120 and gl and gl[0][0] >= 120):
        g, a, b = gl[0]; off = ((g - r) // 2 // 5) * 5; rests = [(a + off, a + off + r)]
    elif r == 120:
        rests = sorted((a + ((g - 60) // 2 // 5) * 5, a + ((g - 60) // 2 // 5) * 5 + 60) for g, a, b in gl[:2])
    p['rests'] = rests; p['brk'] = sum(b - a for a, b in rests); p['work'] = span - p['brk']
    return p
def shifts(F, skillsets, tries=80):
    jobs, naos = naos_split(F)
    best = build(jobs, skillsets)
    for i in range(tries):
        ps = build(jobs, skillsets, random.Random(i))
        if len(ps) < len(best) or (len(ps) == len(best) and sum(q_up(p['last']) - q_down(p['first']) for p in ps) > sum(q_up(p['last']) - q_down(p['first']) for p in best)): best = ps
    return sorted([finish(dict(p)) for p in best], key=lambda p: (p['start'], p['end'])), naos
if __name__ == '__main__':
    k = int(sys.argv[1]); a, b, D, H = PERIODS[k]
    excl = {'HB382', 'SL392'}
    people = train2(load_active('../xl5/in.xlsx')); sk = [p['skills'] for p in people]
    ds = [d for d in days if a <= d <= b]; cache = {}; sh = {}; na = {}
    for d in ds:
        F = frozenset(day_pat[d] - excl)
        if F not in cache: cache[F] = shifts(F, sk)
        sh[d], na[d] = cache[F]
    t = time.time()
    r = roster_skill(ds, sh, people, D + 1, tlimit=420)
    print(a, 'shifts', sum(len(sh[d]) for d in ds), 'peak', max(len(sh[d]) for d in ds), 'support', sum(r['U'].values()), r['status'], round(time.time() - t), flush=True)
    pickle.dump(dict(people=people, a=a, b=b, D=D, H=H, ds=ds, sh=sh, naos=na, F={d: frozenset(day_pat[d] - excl) for d in ds}, **r), open(f'../out/p4_{k}.pkl', 'wb'))
