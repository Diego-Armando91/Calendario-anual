import json, os, datetime as dt, numpy as np
from paths import *
import config as C
B=json.load(open(os.path.join(WORK,'base.json')))
people=B['people']; seq=B['seq']; fl=B['fl']; summ=B['summ']
D=365; P=len(people); start=C.YEAR_START
date=[start+dt.timedelta(i) for i in range(D)]
wd=[d.weekday() for d in date]
def wk(d): return (d+4)//7          # semanas lunes-domingo; 0 = vie 1 a dom 3 ene
NW=wk(D-1)+1
IDX={n.split(' (')[0]:i for i,n in enumerate(people)}
vac=np.array([[fl[p][i]=='theme1' for i in range(D)] for p in people])
CLOSED=C.CLOSED; SPECIAL=C.SPECIAL
closed=np.array([d in CLOSED for d in date])
avail=(~vac)&(~closed[None,:])
SATN={d for d in range(D) if wd[d]==5 and any(a<=date[d]<=b for a,b in C.SAT_NIGHT_PERIODS)}
JA=np.array([{k.split(' (')[0]:v for k,v in summ.items()}[n]['ja'] for n in people],float)
def load_roles():
    R=json.load(open(os.path.join(WORK,'roles.json')))
    role=np.empty((P,D),dtype='U1')
    for i in range(P):
        for d in range(D): role[i,d]=R['roles'][str(i)][wk(d)]
    return role
def save_solution(x,name):
    import csv
    with open(os.path.join(WORK,name),'w',newline='') as f:
        w=csv.writer(f); w.writerow(['persona']+[d.isoformat() for d in date])
        for i,n in enumerate(people): w.writerow([n]+[int(v) for v in x[i]])
def load_solution(name):
    import csv
    rows=list(csv.reader(open(os.path.join(WORK,name))))[1:]
    return np.array([[int(v) for v in r[1:]] for r in rows])
