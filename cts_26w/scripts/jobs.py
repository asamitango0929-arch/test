import openpyxl
from ref import *
NEWF={'7CXXX':'7C1503','RFX0910':'RF352','RFX1830':'RF356'}
def deltas(k, src):
    r=dict(FLIGHTS[src]); r.update(REFTIME.get(src,{}))
    return tm(FLIGHTS[k]['std'])-tm(r['std']), tm(FLIGHTS[k]['sta'])-tm(r['sta'])
def extract(path):
    wb=openpyxl.load_workbook(path)
    _,staff=parse_sheet(wb['1'])
    J={}
    for st in staff:
        cur=None
        for t in st['tasks']:
            k=t['k']
            if k is None: cur=None; continue
            if cur and cur['k']==k and cur['e']==t['s'] and cur['arr']==t['arr']:
                cur['tasks'].append(t); cur['e']=t['e']
            else:
                cur=dict(k=k,s=t['s'],e=t['e'],arr=t['arr'],tasks=[t]); J.setdefault(k,[]).append(cur)
    def mv(job,k,dd,da):
        d=da if job['arr'] else dd
        return dict(k=k,s=job['s']+d,e=job['e']+d,arr=job['arr'],tasks=[dict(t,k=k,s=t['s']+d,e=t['e']+d) for t in job['tasks']])
    out={}
    for k,js in J.items():
        dd,da=deltas(k,k); out[k]=[mv(j,k,dd,da) for j in js]
    for k,src in NEWF.items():
        dd,da=deltas(k,src); out[k]=[mv(j,k,dd,da) for j in J[src]]
    return out
if __name__=='__main__':
    J=extract('../xl/in.xlsx')
    for k,js in J.items():
        print(k,len(js),[('/'.join(t['role'] for t in j['tasks']),hm(j['s']),hm(j['e'])) for j in js])
