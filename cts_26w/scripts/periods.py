from datetime import date, timedelta as td
S,E=date(2026,10,25),date(2027,3,27)
A=[1,2,3,4,5,6,7]
def D(s):
    m={'JAN':1,'FEB':2,'MAR':3,'OCT':10,'NOV':11,'DEC':12}
    return date(2000+int(s[5:7]),m[s[2:5]],int(s[:2]))
FL=[
 ('D7550',[('25OCT26','27MAR27',A)]),
 ('D7552',[('14DEC26','12FEB27',[1,5,7])]),
 ('7C1501',[('01DEC26','27MAR27',A)]),
 ('7C1551',[('01DEC26','01MAR27',[1,3,4,5,7])]),
 ('7C1503',[('25OCT26','02MAR27',A)]),
 ('7CXXX',[('24DEC26','27MAR27',A)]),
 ('RF352',[('25OCT26','27MAR27',A)]),
 ('RF356',[('27NOV26','02MAR27',A)]),
 ('RFX0910',[('13DEC26','19JAN27',[2,4,7]),('26DEC26','27FEB27',[6]),('26JAN27','23FEB27',[2])]),
 ('RFX1830',[('20DEC26','31DEC26',[2,4,7])]),
 ('HB382',[('25OCT26','27MAR27',A)]),
 ('SL392',[('25OCT26','27MAR27',[1,3,5])]),
 ('5J5128',[('07DEC26','27FEB27',A)]),
]
def ops(d):
    w=d.isoweekday()
    return tuple(n for n,ps in FL if any(D(a)<=d<=D(b) and w in ds for a,b,ds in ps))
