"""Qualification-aware check with the real skill matrix (49 staff)."""
import sys, math, random, pickle, time, json
from collections import Counter, defaultdict
import openpyxl
from openpyxl.utils import column_index_from_string as CI
from ortools.sat.python import cp_model
import v3, model
from v3 import flight_jobs, is_naos, day_pat, days, q_down, q_up, REST, code_of
from model import PERIODS, INTERVAL, MAXCONSEC, pad, MINWORK

GROUPS = {'7C': 'I', 'RF': 'O', '5J': 'U', 'VZ': 'AA', 'D7': 'AG', 'JD': 'AM', 'HU': 'AS', 'SL': 'AY', 'HB': 'BE'}
SK = ['DC', 'GI', 'CKIN', 'GATE', 'ARR', 'AP']
def load_people(path):
    ws = openpyxl.load_workbook(path)['SKILL']; P = []
    for r in range(6, 200):
        no = ws.cell(r, 2).value
        if not no or ws.cell(r, 6).value == '休職': continue
        sk, tr = set(), set()
        for cr, col in GROUPS.items():
            c0 = CI(col)
            for i, s in enumerate(SK):
                v = ws.cell(r, c0 + i).value
                if v == '●': sk.add((cr, s))
                elif v == '△': tr.add((cr, s))
        P.append(dict(no=no, name=ws.cell(r, 8).value or ws.cell(r, 7).value, company=ws.cell(r, 5).value, rmk=ws.cell(r, 6).value, skills=sk, train=tr))
    return P
def req(j):
    cr = j['cr']; p = j['pos']
    if p == 'DC': return {(cr, 'DC')}
    if p == 'GI': return {(cr, 'GI')}
    if p == 'CTR': return {(cr, 'CKIN')} if j.get('stay') else {(cr, 'CKIN'), (cr, 'GATE')}
    if p == 'PSPT CHK': return {(cr, 'GATE')}
    if p == 'ARV': return {(cr, 'ARR')}
def build_q(jobs, skillsets, rnd=None):
    people = []
    for j in sorted(jobs, key=lambda j: (j['s'], -j['e'], rnd.random() if rnd else 0)):
        rq = req(j); best = None; bkey = None
        for p in people:
            g = j['s'] - p['last']
            if g < 0: continue
            span = q_up(j['e']) - q_down(p['first'])
            if span > v3.CAP: continue
            if span > 360 and max(p['gmax'], g) < REST: continue
            need = p['req'] | rq
            nq = sum(1 for S in skillsets if need <= S)
            if nq == 0: continue
            key = (nq >= 3, p['last'], rnd.random() if rnd else 0)
            if best is None or key > bkey: best, bkey = p, key
        if best is None:
            best = dict(first=j['s'], last=j['e'], jobs=[], gmax=0, req=set()); people.append(best)
        else: best['gmax'] = max(best['gmax'], j['s'] - best['last'])
        best['jobs'].append(j); best['last'] = j['e']; best['req'] |= rq
    for p in people: v3.finish(p) if False else None
    return people
def shifts_q(F, skillsets, tries=40):
    jobs = [j for k in F for j in flight_jobs(k) if not is_naos(j)]
    best = build_q(jobs, skillsets)
    for i in range(tries):
        ps = build_q(jobs, skillsets, random.Random(i))
        if len(ps) < len(best): best = ps
    out = []
    for p in best:
        p = dict(p); p['jobs'].sort(key=lambda j: j['s'])
        p['start'], p['end'] = q_down(p['first']), q_up(p['last'])
        span = p['end'] - p['start']; p['brk'] = REST if span > 360 else 0; p['work'] = span - p['brk']; p['code'] = code_of(p['start'], p['end'])
        gaps = [(a['e'], b['s']) for a, b in zip(p['jobs'], p['jobs'][1:]) if b['s'] - a['e'] >= REST]
        p['rest'] = (max(gaps, key=lambda g: g[1] - g[0])[0],) if gaps else None
        if p['rest']: p['rest'] = (p['rest'][0], p['rest'][0] + REST)
        p['key'] = (p['code'], frozenset(p['req']))
        out.append(p)
    return out
def roster_skill(ds, sh, people, dcap, tlimit=240):
    m = cp_model.CpModel(); N = len(people)
    types = {}; dem = {}
    for d in ds:
        dem[d] = Counter(s['key'] for s in sh[d])
        for s in sh[d]: types[s['key']] = (s['start'], s['end'], s['work'])
    elig = {(p, t): t[1] <= people[p]['skills'] for p in range(N) for t in types}
    x = {}; unc = {}
    for d in ds:
        for t, n in dem[d].items():
            vs = [x.setdefault((p, d, t), m.NewBoolVar('')) for p in range(N) if elig[p, t]]
            u = m.NewIntVar(0, n, ''); unc[d, t] = u
            m.Add(sum(vs) + u == n)
    W = {}; hrs = []
    for p in range(N):
        for d in ds:
            W[p, d] = sum(x[p, d, t] for t in dem[d] if (p, d, t) in x); m.Add(W[p, d] <= 1)
        m.Add(sum(W[p, d] for d in ds) <= dcap)
        hrs.append(sum(types[t][2] * x[p, d, t] for d in ds for t in dem[d] if (p, d, t) in x))
        for i in range(len(ds) - MAXCONSEC): m.Add(sum(W[p, ds[i+k]] for k in range(MAXCONSEC + 1)) <= MAXCONSEC)
        for i in range(len(ds) - 1):
            d0, d1 = ds[i], ds[i+1]
            for t2 in dem[d1]:
                if (p, d1, t2) not in x: continue
                bad = [x[p, d0, t1] for t1 in dem[d0] if (p, d0, t1) in x and (1440 - types[t1][1]) + types[t2][0] < INTERVAL]
                if bad: m.Add(sum(bad) + x[p, d1, t2] <= 1)
        pair = [m.NewBoolVar('') for i in range(len(ds) - 1)]
        for i, v in enumerate(pair):
            m.Add(v + W[p, ds[i]] <= 1); m.Add(v + W[p, ds[i+1]] <= 1)
            if i: m.Add(pair[i-1] + v <= 1)
        m.Add(sum(pair) >= 2)
    mh = m.NewIntVar(0, 500 * 60, '')
    for h in hrs: m.Add(h <= mh)
    m.Minimize(sum(unc.values()) * 1000000 + mh)
    sv = cp_model.CpSolver(); sv.parameters.max_time_in_seconds = tlimit; sv.parameters.num_workers = 4
    st = sv.Solve(m)
    if st not in (cp_model.OPTIMAL, cp_model.FEASIBLE): return None
    R = [dict(sched={d: t for d in ds for t in dem[d] if (p, d, t) in x and sv.Value(x[p, d, t])}) for p in range(N)]
    for p, r in zip(people, R): r['person'] = p; r['days'] = len(r['sched']); r['mins'] = sum(types[t][2] for t in r['sched'].values())
    U = {(d, t): sv.Value(u) for (d, t), u in unc.items() if sv.Value(u)}
    return dict(R=R, U=U, types=types, status=sv.StatusName(st))
if __name__ == '__main__':
    SC, CAPH, extra = sys.argv[1], int(sys.argv[2]), int(sys.argv[3])
    v3.CAP = CAPH * 60
    excl = {'B': {'HB382', 'SL392'}, 'H': {'HB382'}}[SC]
    people = load_people('../xl5/in.xlsx')
    skillsets = [p['skills'] for p in people]
    out = []
    for a, b, D, H in PERIODS[int(sys.argv[4]):int(sys.argv[5])]:
        ds = [d for d in days if a <= d <= b]; cache = {}; sh = {}
        for d in ds:
            F = frozenset(day_pat[d] - excl)
            if F not in cache: cache[F] = shifts_q(F, skillsets)
            sh[d] = cache[F]
        t = time.time()
        r = roster_skill(ds, sh, people, D + extra)
        tot = sum(len(sh[d]) for d in ds)
        unc = sum(r['U'].values()) if r else None
        print(SC, CAPH, 'h 公休-', extra, a, 'shifts', tot, 'peak', max(len(sh[d]) for d in ds), 'uncovered', unc, r and r['status'], round(time.time() - t), flush=True)
        out.append(dict(a=a, b=b, D=D, H=H, ds=ds, sh=sh, extra=extra, **(r or {})))
    pickle.dump(dict(people=people, out=out), open(f'../out/sq_{SC}_{CAPH}_{extra}_{sys.argv[4]}.pkl', 'wb'))
