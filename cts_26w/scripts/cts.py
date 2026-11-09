import re, openpyxl, datetime
from collections import defaultdict, Counter
def c2m(c): return 240 + (c-5)*5
def hm(m): return f"{m//60:02d}:{m%60:02d}"
FLT = re.compile(r'(D7|7C|RF|SL|HB|5J|STA)\s*(\d{3,5})')
def norm(cr, n):
    if cr == 'STA': cr = 'SL'
    if cr == 'D7': n = n[-3:]  # 9553/79553 -> 553
    return f'{cr}{n}'
def parse(ws):
    sp = {(m.min_row, m.min_col): m.max_col for m in ws.merged_cells.ranges}
    title = ws['D1'].value
    flights = {}
    staff = []
    for r in range(3, ws.max_row+1):
        b = ws.cell(r, 2).value
        cells = []
        for c in range(3, ws.max_column+1):
            v = ws.cell(r, c).value
            if v is None or c < 5: continue
            if not isinstance(v, str): continue
            cells.append((c2m(c), c2m(sp.get((r, c), c)+1), v.replace('\n', ' ').strip()))
        if r < 27:
            for s, e, v in cells:
                m = FLT.search(v); t = re.search(r'(ST[AD])\s*(\d{1,2}:\d{2})', v)
                if m and t: flights[norm(*m.groups())] = dict(kind=t.group(1), time=t.group(2), s=s, e=e)
            continue
        if isinstance(b, str) and re.match(r'^\d+[A-D]\d+[A-D]$', b.strip()):
            staff.append(dict(code=b.strip(), tasks=cells))
    return title, flights, staff
def shift_of(code):
    m = re.match(r'^(\d+)([A-D])(\d+)([A-D])$', code)
    q = {'A': 0, 'B': 15, 'C': 30, 'D': 45}
    return int(m[1])*60+q[m[2]], int(m[3])*60+q[m[4]]
def role_of(v):
    m = FLT.search(v)
    if not m: return None, v
    rest = (v[:m.start()] + ' ' + v[m.end():]).strip()
    rest = re.sub(r'\s+', ' ', rest)
    return norm(*m.groups()), (rest or 'CTR')
if __name__ == '__main__':
    wb = openpyxl.load_workbook('../xl3/in.xlsx')
    tot = defaultdict(Counter); codes = Counter(); others = Counter()
    for ws in wb:
        title, fl, st = parse(ws)
        day = defaultdict(Counter)
        for p in st:
            codes[p['code']] += 1
            seen = set()
            for s, e, v in p['tasks']:
                k, role = role_of(v)
                if k is None: others[v] += 1; continue
                if (k, role) in seen: continue
                seen.add((k, role)); day[k][role] += 1
        print(ws.title, title, 'staff', len(st), {k: dict(c) for k, c in sorted(day.items())})
    print(codes.most_common()); print(others.most_common(40))
