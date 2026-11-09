import json, sys
from datetime import timedelta as td
from ref import FLIGHTS, tm
hm = lambda m: f'{m//60:02d}:{m%60:02d}'
from periods import FL, D, S, E
import v3
from v3 import flight_jobs, shifts_for, day_pat, days, DEP, ARR, ARR_N, PSPT, AGT_MAX, COLOR, CRNAME, codes
WJ = ['', '月', '火', '水', '木', '金', '土', '日']
NEWF = {k for k, f in FLIGHTS.items() if f.get('new')}
def ctr_open(k): f = FLIGHTS[k]; return tm(f['std']) + DEP[f['cr']]['win'][0]
def order_flights(F):
    by = {}
    for k in F: by.setdefault(FLIGHTS[k]['cr'], []).append(k)
    for v in by.values(): v.sort(key=lambda k: (ctr_open(k), tm(FLIGHTS[k]['std'])))
    return [k for c in sorted(by, key=lambda c: ctr_open(by[c][0])) for k in by[c]]
def segments(p):
    segs = []; jobs = p['jobs']
    for i, j in enumerate(jobs):
        lo = p['start'] if i == 0 else jobs[i-1]['e']
        hi = p['end'] if i == len(jobs) - 1 else jobs[i+1]['s']
        if p['rest']:
            if p['rest'][1] <= j['s']: lo = max(lo, p['rest'][1])
            if p['rest'][0] >= j['e']: hi = min(hi, p['rest'][0])
        stretch = j['pos'] in ('GI', 'CTR', 'PSPT CHK')
        if segs: lo = max(lo, segs[-1][1])
        cr = j['cr']
        if j['prep']:
            a0, a1, nm = j['prep']
            if stretch and lo < a0: a0 = max(lo, a1 - AGT_MAX)
            segs.append([a0, a1, nm, 'prep', cr])
        m0, m1 = j['main']
        if j['pos'] == 'CTR' and j.get('gate'):
            segs.append([m0, j['gate'], f"{j['code']} CTR", 'main', cr]); segs.append([j['gate'], m1, f"{j['code']} GATE", 'gate', cr])
        else:
            lab = j['code'] + (' CTR STAY' if j.get('stay') else ' ' + j['pos'])
            segs.append([m0, m1, lab, 'arv' if j['pos'] == 'ARV' else ('gate' if j['pos'] == 'PSPT CHK' else 'main'), cr])
        b0, b1 = j['pw']
        if stretch and hi > b1: b1 = min(hi, b0 + AGT_MAX)
        segs.append([b0, b1, 'PW', 'pw', cr])
    # first segment starts the shift (15-min rounding slack shown as prep)
    if segs and segs[0][0] > p['start']: segs[0][0] = p['start'] if segs[0][3] == 'prep' else segs[0][0]
    if p['rest']: segs.append([p['rest'][0], p['rest'][1], 'REST', 'rest', None])
    return sorted(segs)
def ranges(ds):
    out = []; a = b = ds[0]
    for d in ds[1:]:
        if d == b + td(7): b = d
        else: out.append((a, b)); a = b = d
    out.append((a, b)); return out
pats = {}
for d in days: pats.setdefault(day_pat[d], []).append(d)
plist = sorted(pats.items(), key=lambda x: x[1][0])
pid = {F: f"P{i+1:02d}" for i, (F, _) in enumerate(plist)}
def counts(k):
    f = FLIGHTS[k]; cr = f['cr']; n = DEP[cr]['ctr']; st = v3.CTR_STAY.get(cr, 0); ps = PSPT.get(k, 0)
    return dict(dc=1, gi=1, ctr=n, stay=st, pspt=ps, gate=1 + n - st + ps, arv=ARR_N.get(k, ARR[cr]['n']))
out = dict(colors=COLOR, carriers=CRNAME, patterns=[], periods=[])
for F, ds in plist:
    rows = shifts_for(F)
    fl = []
    for k in order_flights(F):
        f = FLIGHTS[k]; cr = f['cr']; arr, dep = codes(k); std, sta = tm(f['std']), tm(f['sta'])
        w = DEP[cr]['win']; a = ARR[cr]['win']
        fl.append(dict(cr=cr, arr=arr, dep=dep, sta=f['sta'], std=f['std'], ac=f['ac'], dest=f['dest'], ctr=hm(std + w[0]),
                       depBar=[std + w[0], std + w[1]], gateAt=std - v3.GATE_BEFORE, arrBar=[sta + a[0], sta + a[1]], new=k in NEWF, n=counts(k)))
    out['patterns'].append(dict(id=pid[F], wds=sorted({d.isoweekday() for d in ds}), days=len(ds),
        dates=[[a.isoformat(), b.isoformat()] for a, b in ranges(ds)], P=len(rows), H=round(sum(r['work'] for r in rows) / 60, 1),
        flights=fl, rows=[dict(code=r['code'], s=r['start'], e=r['end'], work=r['work'], segs=segments(r)) for r in rows]))
bps = sorted({D(a) for _, ps in FL for a, b, _ in ps} | {D(b) + td(1) for _, ps in FL for a, b, _ in ps} | {S})
bps = [d for d in bps if S <= d <= E] + [E + td(1)]
for a, b in zip(bps[:-1], bps[1:]):
    b = b - td(1); seg = [a + td(i) for i in range((b - a).days + 1)]
    wk = {}
    for w in range(1, 8):
        dd = [d for d in seg if d.isoweekday() == w]
        if dd: wk[w] = pid[day_pat[dd[0]]]
    out['periods'].append(dict(a=a.isoformat(), b=b.isoformat(), days=len(seg), wk=wk))
fl_all = []
for k in order_flights(set(FLIGHTS)):
    f = FLIGHTS[k]; arr, dep = codes(k)
    fl_all.append(dict(cr=f['cr'], arr=arr, dep=dep, sta=f['sta'], std=f['std'], ctr=hm(ctr_open(k)), new=k in NEWF, n=counts(k)))
out['flights'] = fl_all
json.dump(out, open(sys.argv[1], 'w'), ensure_ascii=False, separators=(',', ':'))
print(len(out['patterns']), sum(len(p['rows']) for p in out['patterns']))
