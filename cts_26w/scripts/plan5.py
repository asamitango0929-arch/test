"""Two-step: (1) days-off plan per person, (2) per-day exact assignment of jobs to the people at work (skills, 12h, rests, 4h min, 11h interval); leftovers -> support."""
import sys, math, pickle, time
from collections import Counter, defaultdict
from datetime import timedelta as td
from ortools.sat.python import cp_model
import v3
from v3 import day_pat, days, code_of, q_down, q_up
from skillq import load_active, req, PERIODS
from sq2 import train2
from plan4 import naos_split, need_rest
EXCL = {'HB382', 'SL392'}; CAPQ = 48; MINQ = 16; NSUP = 20
SCARCE = [('D7', 'DC'), ('D7', 'GI'), ('7C', 'DC'), ('RF', 'DC'), ('RF', 'GI'), ('5J', 'DC'), ('5J', 'GI'), ('7C', 'GI')]

def dayoff_plan(people, ds, need, skneed, D, tlimit=90):
    m = cp_model.CpModel(); N = len(people)
    W = {(p, d): m.NewBoolVar('') for p in range(N) for d in ds}
    for p in range(N):
        tot = sum(W[p, d] for d in ds)
        m.Add(tot <= D + 1); m.Add(tot >= D - 1)
        for i in range(len(ds) - 5): m.Add(sum(W[p, ds[i+k]] for k in range(6)) <= 5)
        pair = [m.NewBoolVar('') for i in range(len(ds) - 1)]
        for i, v in enumerate(pair):
            m.Add(v + W[p, ds[i]] <= 1); m.Add(v + W[p, ds[i+1]] <= 1)
            if i: m.Add(pair[i-1] + v <= 1)
        m.Add(sum(pair) >= 2)
    short = []
    for d in ds:
        s = m.NewIntVar(0, 100, ''); m.Add(sum(W[p, d] for p in range(N)) + s >= need[d]); short.append(10 * s)
        for sk, n in skneed[d].items():
            have = [W[p, d] for p in range(N) if sk in people[p]['skills']]
            s2 = m.NewIntVar(0, 10, ''); m.Add(sum(have) + s2 >= n); short.append(30 * s2)
    import os
    base = int(os.environ.get('DBASE', '0'))
    over = []
    if base:
        for p in range(N):
            o = m.NewIntVar(0, 2, ''); m.Add(o >= sum(W[p, d] for d in ds) - base); over.append(o)
    plus = sum(W.values())
    m.Minimize(sum(short) * 100 + sum(over) * 20 - plus)
    sv = cp_model.CpSolver(); sv.parameters.max_time_in_seconds = tlimit; sv.parameters.num_workers = 4
    st = sv.Solve(m)
    return {d: [p for p in range(N) if sv.Value(W[p, d])] for d in ds}

def assign_day(jobs, people, avail, earliest, tlimit=30, latest=None, extra=()):
    m = cp_model.CpModel(); J = sorted(jobs, key=lambda j: (j['s'], j['e'])); n = len(J)
    ws = [('P', p) for p in avail] + [('S', k) for k in range(NSUP)]
    x = {}; used = {}; st = {}; en = {}; ivs = defaultdict(list); spans = {}
    for w in ws:
        sk = None if w[0] == 'S' else people[w[1]]['skills']
        e0 = 0 if w[0] == 'S' else earliest.get(w[1], 0)
        for i, j in enumerate(J):
            if (sk is None or req(j) <= sk) and j['s'] >= e0:
                x[w, i] = m.NewBoolVar('')
    for i in range(n):
        m.AddExactlyOne([x[w, i] for w in ws if (w, i) in x])
    for w in ws:
        mine = [i for i in range(n) if (w, i) in x]
        u = m.NewBoolVar(''); used[w] = u
        if not mine: m.Add(u == 0); continue
        for i in mine: m.AddImplication(x[w, i], u)
        a = m.NewIntVar(0, 96, ''); b = m.NewIntVar(0, 100, ''); st[w] = a; en[w] = b
        for i in mine:
            m.Add(a * 15 <= J[i]['s']).OnlyEnforceIf(x[w, i]); m.Add(b * 15 >= J[i]['e']).OnlyEnforceIf(x[w, i])
        sp = m.NewIntVar(0, CAPQ, ''); m.Add(sp == b - a).OnlyEnforceIf(u); m.Add(sp == 0).OnlyEnforceIf(u.Not()); m.Add(sp >= MINQ).OnlyEnforceIf(u)
        spans[w] = sp
        if w[0] == 'P' and w[1] in earliest: m.Add(a * 15 >= earliest[w[1]]).OnlyEnforceIf(u)
        if latest and w[0] == 'P' and w[1] in latest: m.Add(b * 15 <= latest[w[1]]).OnlyEnforceIf(u)
        n1 = m.NewBoolVar('')
        m.Add(sp > 24).OnlyEnforceIf(n1); m.Add(sp <= 24).OnlyEnforceIf(n1.Not())
        iv = [m.NewOptionalIntervalVar(J[i]['s'], J[i]['e'] - J[i]['s'], J[i]['e'], x[w, i], '') for i in mine]
        rr = []
        for need in (n1,):
            r = m.NewIntVar(0, 300, '')
            iv.append(m.NewOptionalIntervalVar(r * 5, 60, r * 5 + 60, need, ''))
            m.Add(r * 5 >= a * 15).OnlyEnforceIf(need); m.Add(r * 5 + 60 <= b * 15).OnlyEnforceIf(need)
            yb = []; ya = []
            for i in mine:
                y = m.NewBoolVar(''); m.AddImplication(y, x[w, i]); m.Add(J[i]['e'] <= r * 5).OnlyEnforceIf(y); yb.append(y)
                z = m.NewBoolVar(''); m.AddImplication(z, x[w, i]); m.Add(J[i]['s'] >= r * 5 + 60).OnlyEnforceIf(z); ya.append(z)
            m.AddBoolOr(yb).OnlyEnforceIf(need); m.AddBoolOr(ya).OnlyEnforceIf(need)
            rr.append((r, need))
        m.AddNoOverlap(iv)
        used[w] = u; ivs[w] = rr
    sup = [used[('S', k)] for k in range(NSUP)]
    for k in range(NSUP - 1): m.Add(sup[k] >= sup[k+1])
    m.Minimize(sum(sup) * 100000 + sum(used[('P', p)] for p in extra) * 3000 + sum(spans[w] for w in spans if w[0] == 'S') * 10 + sum(spans[w] for w in spans if w[0] == 'P'))
    sv = cp_model.CpSolver(); sv.parameters.max_time_in_seconds = tlimit; sv.parameters.num_workers = 4
    stt = sv.Solve(m)
    if stt not in (cp_model.OPTIMAL, cp_model.FEASIBLE): return None
    out = {}
    for w in ws:
        js = [J[i] for i in range(n) if (w, i) in x and sv.Value(x[w, i])]
        if not js: continue
        a = sv.Value(st[w]) * 15; b = sv.Value(en[w]) * 15
        rests = sorted((sv.Value(r) * 5, sv.Value(r) * 5 + 60) for r, nd in ivs[w] if sv.Value(nd))
        if len(rests) == 2 and rests[0][1] == rests[1][0]: rests = [(rests[0][0], rests[1][1])]
        brk = sum(e - s for s, e in rests)
        lab = ' → '.join(f"{(v3.codes(j['k'])[0] if j['pos'] == 'ARV' else v3.codes(j['k'])[1])} {'CKIN' if j['pos'] == 'CTR' else j['pos']}" for j in sorted(js, key=lambda j: j['s']))
        out[w] = dict(jobs=sorted(js, key=lambda j: j['s']), start=a, end=b, rests=rests, brk=brk, work=b - a - brk, code=code_of(a, b), label=lab,
                      first=min(j['s'] for j in js), last=max(j['e'] for j in js), rest=None, req=set().union(*[req(j) for j in js]))
    return out, sv.StatusName(stt)

if __name__ == '__main__':
    k = int(sys.argv[1]); a, b, D, H = PERIODS[k]; TL = int(sys.argv[2]) if len(sys.argv) > 2 else 30
    people = train2(load_active('../xl5/in.xlsx'))
    ds = [d for d in days if a <= d <= b]
    jobs = {}; naos = {}
    for d in ds:
        F = frozenset(day_pat[d] - EXCL); jobs[d], naos[d] = naos_split(F)
    import os
    ITER = os.environ.get('ITER', '')
    if ITER:
        prev = pickle.load(open(f'../out/p5{"b" if ITER == "c" else ""}_{k}.pkl', 'rb'))
        need = {d: len(prev['res'][d]) for d in ds}
    else:
        p4 = pickle.load(open(f'../out/p3r1_{k}.pkl' if k < 3 else f'../out/p4_{k}.pkl', 'rb'))
        need = {d: len(p4['sh'][d]) for d in ds}
    skneed = {d: Counter(s for j in jobs[d] for s in req(j) if s in SCARCE) for d in ds}
    plan = dayoff_plan(people, ds, need, skneed, D)
    print('dayoff done', flush=True)
    res = {}; prev_end = {}
    for d in ds:
        earliest = {p: max(0, prev_end[p] - 1440 + 660) for p in plan[d] if p in prev_end}
        t = time.time(); r, stt = assign_day(jobs[d], people, plan[d], earliest, TL)
        res[d] = r
        prev_end = {w[1]: s['end'] for w, s in r.items() if w[0] == 'P'}
        print(d, 'avail', len(plan[d]), 'own', sum(1 for w in r if w[0] == 'P'), 'support', sum(1 for w in r if w[0] == 'S'), stt, round(time.time() - t), flush=True)
    pickle.dump(dict(people=people, a=a, b=b, D=D, H=H, ds=ds, plan=plan, res=res, naos=naos, F={d: frozenset(day_pat[d] - EXCL) for d in ds}), open(f'../out/p5{ITER}_{k}.pkl', 'wb'))
    print('saved', sum(sum(1 for w in res[d] if w[0] == 'S') for d in ds))
