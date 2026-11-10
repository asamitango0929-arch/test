import pickle, sys, json
from collections import Counter, defaultdict
from skillq import SK
def analyze(path):
    D = pickle.load(open(path, 'rb')); people = D['people']; res = []
    for o in D['out']:
        if 'U' not in o: res.append(dict(a=str(o['a']), status='none')); continue
        lack = Counter(); lack_day = defaultdict(int)
        for (d, t), n in o['U'].items():
            code, rq = t
            for s in rq: lack[s] += n
            lack_day[str(d)] += n
        maxday = max(lack_day.values()) if lack_day else 0
        days_ = [p['days'] for p in o['R']]; hrs = [p['mins'] / 60 for p in o['R']]
        res.append(dict(a=str(o['a']), shifts=sum(len(o['sh'][d]) for d in o['ds']), uncovered=sum(o['U'].values()), max_unc_day=maxday,
                        lack={f"{c} {s}": n for (c, s), n in lack.most_common()}, days=(min(days_), max(days_)), hrs=(round(min(hrs)), round(max(hrs))), D=o['D'], H=o['H']))
    return people, res
def candidates(people, cr, s):
    out = []
    for p in people:
        if (cr, s) in p['skills']: continue
        score = 0; why = []
        if (cr, s) in p['train']: score += 10; why.append(f'{cr} {s} 訓練中(△)')
        if s in ('DC', 'GI'):
            if (cr, 'CKIN') in p['skills'] and (cr, 'GATE') in p['skills']: score += 4; why.append(f'{cr} CKIN/GATE 有資格')
            other = [c for (c, x) in p['skills'] if x == s]
            if other: score += 3; why.append(f'{s} 有資格（{"・".join(sorted(other))}）')
        else:
            other = [c for (c, x) in p['skills'] if x == s]
            if other: score += 2; why.append(f'{s} 他社有資格（{"・".join(sorted(other))}）')
        if score: out.append((score, p['name'], p['company'], '、'.join(why)))
    return sorted(out, reverse=True)
if __name__ == '__main__':
    for f in sys.argv[1:]:
        people, res = analyze(f)
        print(f); [print(' ', r) for r in res]
