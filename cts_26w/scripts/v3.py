"""CTS 26W: positions per carrier from the CTS line charts, shifts <= 11h (incl. 1h rest), 11h interval, rosters per pay period."""
import math, random
from datetime import date, timedelta as td
from ref import FLIGHTS, tm
from periods import S, E, ops, FL, D

CAP, REST, INTERVAL, MAXCONSEC = 11*60, 60, 11*60, 6
COLOR = {'D7': 'FFF6548E', '7C': 'FFFFC000', 'RF': 'FF92D050', 'SL': 'FFFF5050', 'HB': 'FFFF99CC', '5J': 'FFFFFF00'}
CRNAME = {'D7': 'AirAsia X', '7C': 'Jeju Air', 'RF': 'Aero K', 'SL': 'Thai Lion Air', 'HB': 'Greater Bay', '5J': 'Cebu Pacific'}
# departure template (relative to STD) and arrival template (relative to STA), from the CTS line charts 1DEC-15DEC
DEP = {
 'D7': dict(win=(-180, 5), prep=75, pw=30, ctr=6),
 '7C': dict(win=(-150, 0), prep=60, pw=30, ctr=3),
 'RF': dict(win=(-150, 0), prep=60, pw=30, ctr=3),
 'HB': dict(win=(-180, 0), prep=30, pw=30, ctr=3),
 'SL': dict(win=(-145, 15), prep=50, pw=30, ctr=3),
 '5J': dict(win=(-180, 0), prep=60, pw=40, ctr=3),
}
ARR = {'D7': dict(win=(-40, 80), pw=30, n=3), '7C': dict(win=(-30, 60), pw=30, n=3), 'RF': dict(win=(-30, 60), pw=30, n=3),
       'HB': dict(win=(-40, 70), pw=30, n=3), 'SL': dict(win=(-30, 65), pw=45, n=3), '5J': dict(win=(-30, 60), pw=40, n=3)}
ARR_N = {}   # arrivals: 3 on every flight for OPEN SPOT handling in winter (user rule)          # flights the CTS charts staff with more arrival agents
PSPT = {'7C1503': 2}                         # 7C1504: PSPT CHK x2, STD-60 .. STD-5
CODES = {'RF356': ('RF355', 'RF356')}        # CTS charts: RF355 arrives 16:00, RF356 departs 16:55
PREP_NAME = {'DC': 'DP', 'GI': 'AP', 'CTR': 'CTR準備'}
# user rule: DC prep 60 before CTR OPEN and post work 60 after departure; departure agents at least 45 (up to 60 when free)
PREPOST = {'DC': (60, 60), 'GI': (45, 45), 'CTR': (45, 45), 'AGT': (45, 45)}
AGT_MAX = 60
DC_PREP = {'7CXXX': (60, 30)}   # 7C extra flight (NAOS flight): DC 15:00-19:00 (user rule)
EARLIEST = 6*60
D7_MORNING = {'D7550': (5*60+30, 6*60+30)}   # report times for D7551 (DC/GI/3 CTR, other 3 CTR)
GATE_BEFORE = 60          # CTR agents move to the gate 60 min before STD (26W必要数 line charts: CKIN -> GA/GC)
CTR_STAY = {'D7': 2}      # D7: 2 agents stay at the counter (26W必要数 line charts: CTR STAY)

def codes(k):
    f = FLIGHTS[k]
    if k in CODES: return CODES[k]
    return f['arr'], f['dep']

def flight_jobs(k):
    f = FLIGHTS[k]; cr = f['cr']; arr, dep = codes(k)
    sta, std = tm(f['sta']), tm(f['std'])
    d = DEP[cr]; a, b = std + d['win'][0], std + d['win'][1]
    jobs = []
    for pos, n in (('DC', 1), ('GI', 1), ('CTR', d['ctr'])):
        pre, post = PREPOST[pos]
        if pos == 'DC' and k in DC_PREP: pre, post = DC_PREP[k]
        for i in range(n):
            stay = pos == 'CTR' and i < CTR_STAY.get(cr, 0)
            jobs.append(dict(k=k, cr=cr, pos=pos, code=dep, s=a - pre, e=b + post, main=(a, b), prep=(a - pre, a, PREP_NAME[pos]), pw=(b, b + post),
                             stay=stay, gate=None if (pos != 'CTR' or stay) else std - GATE_BEFORE))
    pre, post = PREPOST['AGT']
    for i in range(PSPT.get(k, 0)):
        m0, m1 = std - 60, std - 5
        jobs.append(dict(k=k, cr=cr, pos='PSPT CHK', code=dep, s=m0 - pre, e=m1 + post, main=(m0, m1), prep=(m0 - pre, m0, '準備'), pw=(m1, m1 + post)))
    if k in D7_MORNING:   # D7551: DC, GI and 3 counter agents report 05:30, the other 3 counter agents 06:30 (user rule)
        ci = 0
        for j in jobs:
            if j['pos'] == 'CTR':
                t = D7_MORNING[k][1] if ci >= 3 else D7_MORNING[k][0]; ci += 1
            elif j['pos'] in ('DC', 'GI'): t = D7_MORNING[k][0]
            else: continue
            j['s'] = t; j['prep'] = (t, j['prep'][1], j['prep'][2])
    else:
        for j in jobs:   # other D7 flights: never before 06:00
            if cr == 'D7' and j['s'] < EARLIEST:
                j['s'] = EARLIEST
                if j['prep']: j['prep'] = (EARLIEST, j['prep'][1], j['prep'][2])
    r = ARR[cr]; n = ARR_N.get(k, r['n'])
    for i in range(n):
        jobs.append(dict(k=k, cr=cr, pos='ARV', code=arr, s=sta + r['win'][0], e=sta + r['win'][1] + r['pw'], main=(sta + r['win'][0], sta + r['win'][1]), prep=None, pw=(sta + r['win'][1], sta + r['win'][1] + r['pw'])))
    return jobs

def q_down(m): return m - m % 15
def q_up(m): return m + (-m) % 15
def code_of(s, e):
    L = 'ABCD'
    return f"{s//60}{L[(s%60)//15]}{e//60}{L[(e%60)//15]}"

def build1(jobs, rnd=None):
    people = []
    for j in sorted(jobs, key=lambda j: (j['s'], -j['e'], rnd.random() if rnd else 0)):
        best = None; bkey = None
        for p in people:
            g = j['s'] - p['last']
            if g < 0: continue
            s0, e0 = q_down(p['first']), q_up(j['e'])
            span = e0 - s0
            if span > CAP: continue
            gmax = max(p['gmax'], g)
            if span > 360 and gmax < REST: continue
            key = (p['last'], rnd.random() if rnd else 0)
            if best is None or key > bkey: best, bkey = p, key
        if best is None:
            best = dict(first=j['s'], last=j['e'], jobs=[], gmax=0); people.append(best)
        else:
            best['gmax'] = max(best['gmax'], j['s'] - best['last'])
        best['jobs'].append(j); best['last'] = j['e']
    for p in people: finish(p)
    people.sort(key=lambda p: (p['start'], p['end']))
    return people

def finish(p):
    p['start'], p['end'] = q_down(p['first']), q_up(p['last'])
    span = p['end'] - p['start']
    p['brk'] = REST if span > 360 else 0
    p['work'] = span - p['brk']
    p['code'] = code_of(p['start'], p['end'])
    p['rest'] = None
    if p['brk']:
        gaps = [(a['e'], b['s']) for a, b in zip(p['jobs'], p['jobs'][1:]) if b['s'] - a['e'] >= REST]
        a, b = max(gaps, key=lambda g: g[1] - g[0])
        p['rest'] = (a, a + REST)

def build(jobs, tries=80):
    sc = lambda ps: (len(ps), sum(p['work'] for p in ps))
    best = build1(jobs)
    for i in range(tries):
        ps = build1(jobs, random.Random(i))
        if sc(ps) < sc(best): best = ps
    return best

days = [S + td(i) for i in range((E - S).days + 1)]
day_pat = {d: frozenset(ops(d)) for d in days}
# 7C extra flight (18:30 dep): NAOS sends 7 people daily, 15:15-19:15 (4h) -> GI, CKIN x3, ARV x3
NAOS_FLIGHT, NAOS_POS, NAOS_SHIFT = '7CXXX', ('GI', 'CTR', 'ARV'), (15*60+15, 19*60+15)
def is_naos(j): return j['k'] == NAOS_FLIGHT and j['pos'] in NAOS_POS
_cache = {}
def shifts_for(F):
    if F not in _cache:
        _cache[F] = build([j for k in F for j in flight_jobs(k) if not is_naos(j)])
    return _cache[F]
def naos_for(F):
    if NAOS_FLIGHT not in F: return []
    a, b = NAOS_SHIFT; out = []
    for j in flight_jobs(NAOS_FLIGHT):
        if is_naos(j):
            assert a <= j['s'] and j['e'] <= b, j
            out.append(dict(first=j['s'], last=j['e'], jobs=[j], start=a, end=b, brk=0, work=b - a, rest=None, code=code_of(a, b), naos=True))
    return out

PAY = [(date(2026,10,16), date(2026,11,15), 21, 168), (date(2026,11,16), date(2026,12,15), 20, 160),
       (date(2026,12,16), date(2027,1,15), 22, 176), (date(2027,1,16), date(2027,2,15), 21, 168),
       (date(2027,2,16), date(2027,3,15), 19, 152), (date(2027,3,16), date(2027,4,15), 21, 168)]

def roster(ds, N, maxdays, maxhours, carry=None):
    """assign every shift of every day to one of N people; returns list of persons or None"""
    P = [dict(no=i+1, days=0, mins=0, consec=0, last=None, lastday=None, sched={}) for i in range(N)]
    if carry:
        for p, c in zip(P, carry): p['last'], p['lastday'], p['consec'] = c
    for d in ds:
        sh = sorted(shifts_for(day_pat[d]), key=lambda s: s['start'])
        used = set()
        for s in sh:
            cand = []
            for p in P:
                if p['no'] in used or p['days'] >= maxdays or p['mins'] + s['work'] > maxhours*60: continue
                worked_y = p['lastday'] == d - td(1)
                if worked_y and (1440 - p['last']) + s['start'] < INTERVAL: continue
                if worked_y and p['consec'] >= MAXCONSEC: continue
                # prefer: fewest days, then fewest minutes, then those who can't take tomorrow's early shift anyway
                cand.append(((p['days'] * 600 + p['mins'] / 60), p['consec'] if worked_y else 0, p['no'], p))
            if not cand: return None
            p = min(cand, key=lambda c: c[:3])[3]
            used.add(p['no'])
            p['consec'] = p['consec'] + 1 if p['lastday'] == d - td(1) else 1
            p['days'] += 1; p['mins'] += s['work']; p['last'] = s['end']; p['lastday'] = d
            p['sched'][d] = s
    return P

def solve_period(a, b, sd, sh):
    ds = [d for d in days if a <= d <= b]
    frac = len(ds) / ((b - a).days + 1)
    md, mh = round(sd * frac), round(sh * frac, 1)
    tot_p = sum(len(shifts_for(day_pat[d])) for d in ds); tot_h = sum(s['work'] for d in ds for s in shifts_for(day_pat[d])) / 60
    lb = max(math.ceil(tot_p / md), math.ceil(tot_h / mh), max(len(shifts_for(day_pat[d])) for d in ds))
    N = lb
    while True:
        R = roster(ds, N, md, mh)
        if R: break
        N += 1
    return dict(a=a, b=b, ds=ds, frac=frac, md=md, mh=mh, P=tot_p, H=tot_h, lb=lb, N=N, R=R, sd=sd, sh=sh)

if __name__ == '__main__':
    import sys
    for a, b, sd, sh in PAY:
        r = solve_period(a, b, sd, sh)
        print(a, b, 'days', len(r['ds']), 'maxd', r['md'], 'maxh', r['mh'], 'P', r['P'], 'H', round(r['H']), 'LB', r['lb'], 'N', r['N'])

# ---------- exact day shifts (CP-SAT): fewest people, then least paid time ----------
def build_cp(jobs, tlimit=60):
    from ortools.sat.python import cp_model
    greedy = build(jobs)
    K = len(greedy)
    J = sorted(jobs, key=lambda j: (j['s'], j['e']))
    n = len(J)
    m = cp_model.CpModel()
    x = [[m.NewBoolVar('') for _ in range(K)] for _ in range(n)]
    used = [m.NewBoolVar('') for _ in range(K)]
    for i in range(n): m.AddExactlyOne(x[i])
    for k in range(K - 1): m.Add(used[k] >= used[k+1])
    # incompatible pairs (overlap or too far apart for one shift)
    for i in range(n):
        for i2 in range(i + 1, n):
            a, b = J[i], J[i2]
            if a['s'] < b['e'] and b['s'] < a['e'] or q_up(max(a['e'], b['e'])) - q_down(min(a['s'], b['s'])) > CAP:
                for k in range(K): m.AddBoolOr([x[i][k].Not(), x[i2][k].Not()])
    spans = []; RS = []
    for k in range(K):
        for i in range(n): m.AddImplication(x[i][k], used[k])
        st = m.NewIntVar(0, 96, ''); en = m.NewIntVar(0, 96, '')   # quarter hours
        for i in range(n):
            m.Add(st * 15 <= J[i]['s']).OnlyEnforceIf(x[i][k]); m.Add(en * 15 >= J[i]['e']).OnlyEnforceIf(x[i][k])
        span = m.NewIntVar(0, CAP, ''); m.Add(span == (en - st) * 15).OnlyEnforceIf(used[k]); m.Add(span == 0).OnlyEnforceIf(used[k].Not())
        needr = m.NewBoolVar(''); m.Add(span > 360).OnlyEnforceIf(needr); m.Add(span <= 360).OnlyEnforceIf(needr.Not())
        rs = m.NewIntVar(0, 96*3, '')    # rest start in 5-min units
        rest = m.NewOptionalIntervalVar(rs * 5, REST, rs * 5 + REST, needr, '')
        m.Add(rs * 5 >= st * 15).OnlyEnforceIf(needr); m.Add(rs * 5 + REST <= en * 15).OnlyEnforceIf(needr)
        ivs = [m.NewOptionalIntervalVar(J[i]['s'], J[i]['e'] - J[i]['s'], J[i]['e'], x[i][k], '') for i in range(n)]
        m.AddNoOverlap(ivs + [rest])
        # the rest sits between two jobs of the shift (never at its start or end)
        ybef, yaft = [], []
        for i in range(n):
            y = m.NewBoolVar(''); m.AddImplication(y, x[i][k]); m.Add(J[i]['e'] <= rs * 5).OnlyEnforceIf(y); ybef.append(y)
            z = m.NewBoolVar(''); m.AddImplication(z, x[i][k]); m.Add(J[i]['s'] >= rs * 5 + REST).OnlyEnforceIf(z); yaft.append(z)
        m.AddBoolOr(ybef).OnlyEnforceIf(needr); m.AddBoolOr(yaft).OnlyEnforceIf(needr)
        spans.append(span); RS.append((rs, needr))
    m.Minimize(sum(used) * 10000 + sum(spans))
    sv = cp_model.CpSolver(); sv.parameters.max_time_in_seconds = tlimit; sv.parameters.num_workers = 4
    stt = sv.Solve(m)
    if stt not in (cp_model.OPTIMAL, cp_model.FEASIBLE): return greedy
    people = []
    for k in range(K):
        js = [J[i] for i in range(n) if sv.Value(x[i][k])]
        if not js: continue
        js.sort(key=lambda j: j['s'])
        p = dict(first=js[0]['s'], last=max(j['e'] for j in js), jobs=js, gmax=0)
        finish_cp(p)
        if p['brk']:
            r0 = sv.Value(RS[k][0]) * 5; p['rest'] = (r0, r0 + REST)
        people.append(p)
    people.sort(key=lambda p: (p['start'], p['end']))
    return people if len(people) <= len(greedy) else greedy

def finish_cp(p):
    p['start'], p['end'] = q_down(p['first']), q_up(p['last'])
    span = p['end'] - p['start']
    p['brk'] = REST if span > 360 else 0
    p['work'] = span - p['brk']; p['code'] = code_of(p['start'], p['end']); p['rest'] = None
    if p['brk']:
        # widest free window inside the shift (between jobs, or at the shift edges after rounding)
        pts = [(p['start'], p['jobs'][0]['s'])] + [(a['e'], b['s']) for a, b in zip(p['jobs'], p['jobs'][1:])] + [(p['last'], p['end'])]
        mid = [g for g in pts[1:-1] if g[1] - g[0] >= REST]
        if mid:
            g = max(mid, key=lambda g: g[1] - g[0]); p['rest'] = (g[0], g[0] + REST)
