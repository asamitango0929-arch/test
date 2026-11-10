"""49 staff fixed: cut days off / overtime instead of dispatch. Shifts up to OTCAP hours (incl. 1h rest)."""
import sys, math, pickle, time
from collections import Counter
from ortools.sat.python import cp_model
import v3, model
from model import PERIODS, pad, FREE, work_of, INTERVAL, MAXCONSEC
SC, OTCAP, NFIX = sys.argv[1], int(sys.argv[2]), 49
v3.CAP = OTCAP * 60
model.EXCL |= {'B': {'HB382', 'SL392'}, 'H': {'HB382'}}[SC]
def roster_fixed(ds, sh, N, dcap, tlimit=180):
    m = cp_model.CpModel(); codes = {}; dem = {}
    for d in ds:
        dem[d] = Counter(s['code'] for s in sh[d])
        for s in sh[d]: codes[s['code']] = (s['start'], s['end'], s['work'])
    x = {(p, d, c): m.NewBoolVar('') for p in range(N) for d in ds for c in dem[d]}
    for d in ds:
        for c, n in dem[d].items(): m.Add(sum(x[p, d, c] for p in range(N)) == n)
    W = {}; days = []; hrs = []
    for p in range(N):
        for d in ds:
            W[p, d] = sum(x[p, d, c] for c in dem[d]); m.Add(W[p, d] <= 1)
        dp = sum(W[p, d] for d in ds); m.Add(dp <= dcap); days.append(dp)
        hp = sum(codes[c][2] * x[p, d, c] for d in ds for c in dem[d]); hrs.append(hp)
        for i in range(len(ds) - MAXCONSEC):
            m.Add(sum(W[p, ds[i+k]] for k in range(MAXCONSEC + 1)) <= MAXCONSEC)
        for i in range(len(ds) - 1):
            d0, d1 = ds[i], ds[i+1]
            for c2 in dem[d1]:
                bad = [c1 for c1 in dem[d0] if (1440 - codes[c1][1]) + codes[c2][0] < INTERVAL]
                if bad: m.Add(sum(x[p, d0, c1] for c1 in bad) + x[p, d1, c2] <= 1)
        pair = [m.NewBoolVar('') for i in range(len(ds) - 1)]
        for i, v in enumerate(pair):
            m.Add(v + W[p, ds[i]] <= 1); m.Add(v + W[p, ds[i+1]] <= 1)
            if i: m.Add(pair[i-1] + v <= 1)
        m.Add(sum(pair) >= 2)
    mh = m.NewIntVar(0, 400 * 60, ''); md_ = m.NewIntVar(0, 31, '')
    for h in hrs: m.Add(h <= mh)
    for dd in days: m.Add(dd <= md_)
    m.Minimize(md_ * 100000 + mh)
    sv = cp_model.CpSolver(); sv.parameters.max_time_in_seconds = tlimit; sv.parameters.num_workers = 4
    st = sv.Solve(m)
    if st not in (cp_model.OPTIMAL, cp_model.FEASIBLE): return None
    R = []
    for p in range(N):
        sched = {d: c for d in ds for c in dem[d] if sv.Value(x[p, d, c])}
        R.append(dict(no=p + 1, sched=sched, days=len(sched), mins=sum(codes[c][2] for c in sched.values())))
    return dict(R=R, codes=codes, free={d: 0 for d in ds})
out = []
for a, b, D, H in PERIODS:
    ds, sh = model.period_shifts(a, b, model.MINWORK)
    P = sum(len(sh[d]) for d in ds); peak = max(len(sh[d]) for d in ds)
    dcap = max(math.ceil(P / NFIX), 1)
    t = time.time(); r = None
    if peak <= NFIX:
        while dcap <= len(ds) - 6:
            r = roster_fixed(ds, sh, NFIX, dcap)
            print(a, 'dcap', dcap, 'ok' if r else 'no', round(time.time() - t), flush=True)
            if r: break
            dcap += 1
    print(a, 'P', P, 'peak', peak, 'result', 'OK' if r else 'NG', flush=True)
    out.append(dict(a=a, b=b, D=D, H=H, N=NFIX, T=model.MINWORK, ds=ds, sh=sh, peak=peak, P=P, ok=bool(r),
                    naos={d: v3.naos_for(model.Fof(d)) for d in ds}, F={d: model.Fof(d) for d in ds}, **(r or dict(R=[], codes={}, free={}))))
pickle.dump(out, open(f'../out/ot_{SC}_{OTCAP}.pkl', 'wb')); print('saved')
