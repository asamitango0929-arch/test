import re, openpyxl, datetime
X0=21  # column of 04:00 ; 5 min per column
def c2m(col): return 240+(col-X0)*5
def m2c(m): return X0+(m-240)//5
def hm(m): return f"{m//60:02d}{m%60:02d}"
def tm(s): h,m=s.split(':'); return int(h)*60+int(m)

FLIGHTS={  # key: (label, carrier, type, WN, dest, freq, sta, std, ctr, fill, depcode, arrcode)
 'D7550':dict(lab='D7550/D7551',cr='D7',ac='A330',wn='Wide',dest='KUL',sta='08:40',std='09:55',ctr='06:55',fill='FF0070C0',arr='D7550',dep='D7551'),
 'D7552':dict(lab='D7552/D7553',cr='D7',ac='A330',wn='Wide',dest='KUL',sta='17:00',std='18:15',ctr='15:15',fill='FF0070C0',arr='D7552',dep='D7553'),
 '7C1501':dict(lab='7C1501/7C1502',cr='7C',ac='B737',wn='Narrow',dest='ICN',sta='10:00',std='11:00',ctr='08:30',fill='FFFFC000',arr='7C1501',dep='7C1502'),
 '7C1551':dict(lab='7C1551/7C1552',cr='7C',ac='B737',wn='Narrow',dest='PUS',sta='13:40',std='14:40',ctr='12:10',fill='FFFFC000',arr='7C1551',dep='7C1552'),
 '7C1503':dict(lab='7C1503/7C1504',cr='7C',ac='B737',wn='Narrow',dest='ICN',sta='15:00',std='15:55',ctr='13:25',fill='FFFFC000',arr='7C1503',dep='7C1504'),
 '7CXXX':dict(lab='7CXXXX/7CXXXX【増便】',cr='7C',ac='B737',wn='Narrow',dest='XXX',sta='17:30',std='18:30',ctr='16:00',fill='FFFFC000',arr='7C増便',dep='7C増便',new=True),
 'RF352':dict(lab='RF352/RF351',cr='RF',ac='A320',wn='Narrow',dest='CJJ',sta='10:30',std='11:30',ctr='09:00',fill='FF92D050',arr='RF352',dep='RF351'),
 'RF356':dict(lab='RF356/RF355',cr='RF',ac='A320',wn='Narrow',dest='CJJ',sta='16:00',std='16:55',ctr='14:25',fill='FF92D050',arr='RF356',dep='RF355'),
 'RFX0910':dict(lab='RFXXX/RFXXX【増便】',cr='RF',ac='A320',wn='Narrow',dest='XXX',sta='09:10',std='10:00',ctr='07:30',fill='FF92D050',arr='RF増便',dep='RF増便',new=True),
 'RFX1830':dict(lab='RFXXX/RFXXX【増便】',cr='RF',ac='A320',wn='Narrow',dest='XXX',sta='18:30',std='19:30',ctr='17:00',fill='FF92D050',arr='RF増便',dep='RF増便',new=True),
 'HB382':dict(lab='HB382/HB383',cr='HB',ac='B787',wn='Narrow',dest='HKG',sta='15:00',std='16:15',ctr='13:45',fill='FFFF99CC',arr='HB382',dep='HB383'),
 'SL392':dict(lab='SL392/SL393',cr='SL',ac='B737',wn='Narrow',dest='KHH/DMK',sta='13:10',std='14:10',ctr='11:40',fill='FFC00000',arr='SL392',dep='SL393'),
 '5J5128':dict(lab='5J5128/5J5129',cr='5J',ac='A320',wn='Narrow',dest='MNL',sta='16:50',std='17:50',ctr='14:50',fill='FFFFFF00',arr='5J5128',dep='5J5129'),
}
# reference (sheet) times for flights whose times differ from the Operating Period sheet
REFTIME={'HB382':dict(sta='15:00',std='16:00',ctr='13:30')}
CODEMAP=[(r'D7\s*9?55[01]','D7550'),(r'D7\s*9?55[23]','D7552'),(r'7C\s*(1501|1502|5102|5101)','7C1501'),(r'7C\s*155[12]','7C1551'),
 (r'7C\s*(1503|1504|5104|5103)','7C1503'),(r'RF\s*35[12]','RF352'),(r'RF\s*35[56]','RF356'),(r'HB\s*88[01]','HB382'),(r'HB\s*38[23]','HB382'),(r'SL\s*39[23]','SL392'),(r'5J\s*512[89]','5J5128')]
NEUTRAL=re.compile(r'^(PRE|PW|R)\b')
ARRROLE=re.compile(r'\b(AS|AR|BUS)\b')
def code_of(text):
    for pat,k in CODEMAP:
        if re.search(pat,text): return k
    return None
def window(k,src=None):
    f=dict(FLIGHTS[k]); f.update(REFTIME.get(k,{}) if src=='ref' else {})
    return tm(f['ctr']),tm(f['std']),tm(f['sta'])
COLOR2CR={'FF0070C0':'D7','FFFFC000':'7C','FF92D050':'RF','FFFF99CC':'HB','FFC00000':'SL','FFFFFF00':'5J'}

def merged_span(ws):
    sp={}
    for m in ws.merged_cells.ranges: sp[(m.min_row,m.min_col)]=m.max_col
    return sp

def parse_sheet(ws):
    sp=merged_span(ws)
    flights=[];staff=[];hdr=None
    for r in range(2,ws.max_row+1):
        if ws.cell(r,3).value=='Shift Code': hdr=r; break
        h=ws.cell(r,8).value
        if isinstance(h,str) and '/' in h: flights.append(code_of(h))
    for r in range(hdr+1,ws.max_row+1):
        if not isinstance(ws.cell(r,2).value,int): 
            if ws.cell(r,3).value=='AVERAGE': break
            continue
        tasks=[]
        for c in range(X0-3,ws.max_column+1):
            v=ws.cell(r,c).value
            if v is None or c<X0: continue
            fill=ws.cell(r,c).fill.fgColor.rgb if ws.cell(r,c).fill.fill_type else None
            end=sp.get((r,c),c)
            tasks.append(dict(s=c2m(c),e=c2m(end+1),text=str(v).strip(),fill=fill if fill in COLOR2CR else None))
        def tv(x):
            if isinstance(x,datetime.time): return x.hour*60+x.minute
            if isinstance(x,str) and ':' in x: return tm(x)
            return None
        staff.append(dict(no=ws.cell(r,2).value,start=tv(ws.cell(r,11).value),end=tv(ws.cell(r,12).value),brk=tv(ws.cell(r,13).value),tasks=tasks))
    # assign flights to tasks
    for st in staff:
        prev=None
        for t in st['tasks']:
            role=t['text'].split('\n')[0]
            k=code_of(t['text'])
            if k is None and not NEUTRAL.match(role):
                if t['fill']:
                    cr=COLOR2CR[t['fill']]
                    cands=[f for f in flights if f and FLIGHTS[f]['cr']==cr]
                    if prev and prev.get('k') in cands and prev['e']==t['s']:
                        k=prev['k']
                    else:
                        fit=[f for f in cands if window(f,'ref')[0]<=t['s'] and t['e']<=window(f,'ref')[1]+5]
                        k=min(fit or cands,key=lambda f:abs(window(f,'ref')[1]-t['e'])) if cands else None
                elif prev and prev.get('k') and prev['e']==t['s']:
                    k=prev['k']
            t['k']=k
            t['role']=re.sub(r'^\S*\d\S*\s+','',role) if code_of(role) else role
            t['arr']=bool(k and ARRROLE.search(role))
            prev=t
    return [f for f in flights if f],staff

if __name__=='__main__':
    wb=openpyxl.load_workbook('../xl/in.xlsx')
    for n in '1234567':
        fl,st=parse_sheet(wb[n])
        unk=[(s['no'],t['text'].replace('\n',' ')) for s in st for t in s['tasks'] if t['k'] is None and not NEUTRAL.match(t['text'])]
        print(n,fl,len(st),'unmapped:',unk)
    fl,st=parse_sheet(wb['1'])
    for s in st[:60]: print(s['no'],hm(s['start']),hm(s['end']),[(t['k'],t['role'],hm(t['s']),hm(t['e'])) for t in s['tasks']])
