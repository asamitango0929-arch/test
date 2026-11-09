import pickle, time, sys
import model
from model import *
SC = sys.argv[1]
model.EXCL |= {'A': set(), 'B': {'HB382', 'SL392'}, 'S': {'SL392'}, 'H': {'HB382'}}[SC]
out = []
for a, b, D, H in PERIODS:
    N, T = choose(a, b, D, H)
    ds, sh = period_shifts(a, b, T)
    t = time.time()
    while True:
        r = roster(ds, sh, N, D, H, tlimit=150)
        print(a, 'N', N, 'T', T / 60, 'ok' if r else 'no', round(time.time() - t), flush=True)
        if r: break
        N += 1
    out.append(dict(a=a, b=b, D=D, H=H, N=N, T=T, ds=ds, sh=sh, naos={d: naos_for(Fof(d)) for d in ds}, F={d: Fof(d) for d in ds}, **r))
pickle.dump(out, open(f'../out/solved_{SC}.pkl', 'wb'))
print('saved')
