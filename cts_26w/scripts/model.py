"""Per pay period: pad shifts toward the monthly hours, then build a roster with CP-SAT.
Rules: shift <= 11h incl. 1h rest, min 4h work, 15-min codes, 11h interval, max 5 consecutive days,
at least two separate 2-day-off blocks per person, days worked = monthly 所定勤務日数, hours <= 所定労働時間."""
import math, copy
from collections import Counter, defaultdict
from datetime import date, timedelta as td
from ortools.sat.python import cp_model
import v3
from v3 import shifts_for, naos_for, day_pat, days, code_of, q_down, q_up
EXCL = set()
def Fof(d): return frozenset(day_pat[d] - EXCL)

CAP, REST, MINWORK, INTERVAL, MAXCONSEC = 11*60, 60, 4*60, 11*60, 5
EARLY, LATE = 6*60, 22*60
FREE = dict(start=9*60, end=18*60)   # 予備 (FREE ASGN) day, 8h work
PERIODS = [(date(2026,11,16), date(2026,12,15), 20, 160), (date(2026,12,16), date(2027,1,15), 22, 176),
           (date(2027,1,16), date(2027,2,15), 21, 168), (date(2027,2,16), date(2027,3,15), 19, 152)]

def work_of(span): return span - (REST if span > 360 else 0)

def pad(sh, T):
    """extend one shift so that its work time is at least T (FREE ASGN time), staying within 11h"""
    s, e = sh['start'], sh['end']
    span = e - s
    if work_of(span) >= T: return dict(sh, pad=None, ostart=s, oend=e, orest=sh.get('rest'))
    need = T + (REST if T > 360 else 0)
    need = min(need, CAP)
    s2 = s - (need - span)
    s2 -= s2 % 30                                  # align new start to :00 / :30
    if s2 >= EARLY:
        e2 = max(e, s2 + need)
        if e2 - s2 > CAP: s2 = e2 - CAP
        padseg = (s2, s)
    else:
        s2 = s; e2 = s + need; e2 += (-e2) % 30
        e2 = min(e2, LATE, s2 + CAP)
        padseg = (e, e2)
    out = dict(sh, start=s2, end=e2, pad=padseg, ostart=s, oend=e, orest=sh.get('rest'))
    out['work'] = work_of(e2 - s2); out['code'] = code_of(s2, e2)
    out['brk'] = REST if e2 - s2 > 360 else 0
    if out['brk'] and not sh.get('rest'):
        a, b = padseg                               # rest inside the added time, next to the jobs
        out['rest'] = (b - REST, b) if a < s else (a, a + REST)
        if out['rest'][0] < a: out['rest'] = None
    return out

def period_shifts(a, b, T):
    ds = [d for d in days if a <= d <= b]
    return ds, {d: [pad(s, T) for s in shifts_for(Fof(d))] for d in ds}

def choose(a, b, D, H):
    ds = [d for d in days if a <= d <= b]
    base = {d: shifts_for(Fof(d)) for d in ds}
    P = sum(len(v) for v in base.values())
    N = max(math.ceil(P / D), max(len(v) for v in base.values()))
    best = None
    for T in range(MINWORK, 8*60 + 1, 15):
        _, sh = period_shifts(a, b, T)
        Hs = sum(s['work'] for v in sh.values() for s in v) + (N * D - P) * work_of(FREE['end'] - FREE['start'])
        if Hs <= N * H * 60 * 0.985: best = T
    return N, best

def roster(ds, sh, N, D, H, tlimit=90):
    m = cp_model.CpModel()
    codes = {}                      # code -> (start, end, work)
    dem = {}
    for d in ds:
        c = Counter(s['code'] for s in sh[d]); dem[d] = c
        for s in sh[d]: codes[s['code']] = (s['start'], s['end'], s['work'])
    fcode = 'FREE'
    codes[fcode] = (FREE['start'], FREE['end'], work_of(FREE['end'] - FREE['start']))
    P = sum(sum(c.values()) for c in dem.values())
    x = {}
    for p in range(N):
        for d in ds:
            for c in list(dem[d]) + [fcode]:
                x[p, d, c] = m.NewBoolVar('')
    free = {d: m.NewIntVar(0, N, '') for d in ds}
    for d in ds:
        for c, n in dem[d].items(): m.Add(sum(x[p, d, c] for p in range(N)) == n)
        m.Add(sum(x[p, d, fcode] for p in range(N)) == free[d])
    W = {}
    for p in range(N):
        for d in ds:
            W[p, d] = sum(x[p, d, c] for c in list(dem[d]) + [fcode])
            m.Add(W[p, d] <= 1)
        m.Add(sum(W[p, d] for d in ds) == D)
        hrs = sum(codes[c][2] * x[p, d, c] for d in ds for c in list(dem[d]) + [fcode])
        m.Add(hrs <= H * 60)
        for i in range(len(ds) - MAXCONSEC):
            m.Add(sum(W[p, ds[i+k]] for k in range(MAXCONSEC + 1)) <= MAXCONSEC)
        for i in range(len(ds) - 1):
            d0, d1 = ds[i], ds[i+1]
            for c2 in list(dem[d1]) + [fcode]:
                bad = [c1 for c1 in list(dem[d0]) + [fcode] if (1440 - codes[c1][1]) + codes[c2][0] < INTERVAL]
                if bad: m.Add(sum(x[p, d0, c1] for c1 in bad) + x[p, d1, c2] <= 1)
        pair = [m.NewBoolVar('') for i in range(len(ds) - 1)]
        for i, v in enumerate(pair):
            m.Add(v + W[p, ds[i]] <= 1); m.Add(v + W[p, ds[i+1]] <= 1)
            if i: m.Add(pair[i-1] + v <= 1)
        m.Add(sum(pair) >= 2)
    # spread spare (FREE) days, keep weekly shape
    mx = m.NewIntVar(0, N, '')
    for d in ds: m.Add(free[d] <= mx)
    m.Minimize(mx)
    # symmetry: person p works at least as many early days as p+1 is not needed; keep simple
    sv = cp_model.CpSolver(); sv.parameters.max_time_in_seconds = tlimit; sv.parameters.num_workers = 4
    st = sv.Solve(m)
    if st not in (cp_model.OPTIMAL, cp_model.FEASIBLE): return None
    R = []
    for p in range(N):
        sched = {}
        for d in ds:
            for c in list(dem[d]) + [fcode]:
                if sv.Value(x[p, d, c]): sched[d] = c
        R.append(dict(no=p + 1, sched=sched, days=len(sched), mins=sum(codes[c][2] for c in sched.values())))
    return dict(R=R, codes=codes, free={d: sv.Value(free[d]) for d in ds})

if __name__ == '__main__':
    import sys, time
    for a, b, D, H in PERIODS[int(sys.argv[1]):int(sys.argv[2])]:
        N, T = choose(a, b, D, H)
        ds, sh = period_shifts(a, b, T)
        t = time.time()
        while True:
            r = roster(ds, sh, N, D, H)
            print(a, 'N', N, 'T', T/60, 'ok' if r else 'infeasible', round(time.time() - t))
            if r: break
            N += 1
        hrs = [p['mins'] / 60 for p in r['R']]
        print('  hours min/max', min(hrs), max(hrs), 'free days', sum(r['free'].values()))
