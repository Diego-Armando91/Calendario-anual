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
# Fines de semana anteriores a las vacaciones: libres (se tratan como vacaciones). Ver CLAUDE.md, regla 1b.
_LIB=os.path.join(DATA,'libres_antes_vacaciones.json')
LIBRES=np.zeros_like(vac)
if getattr(C,'LIBRES_ANTES_VACACIONES',True) and os.path.exists(_LIB):
    for _n,_d in json.load(open(_LIB)):
        LIBRES[people.index(_n),(dt.date.fromisoformat(_d)-C.YEAR_START).days]=True
    vac=vac|LIBRES
# Días no disponibles por motivos personales (p. ej. fines de semana con hijos): data/no_disponible.json
_ND=os.path.join(DATA,'no_disponible.json')
if os.path.exists(_ND):
    for _n,_d in json.load(open(_ND)):
        _i=people.index(_n); _k=(dt.date.fromisoformat(_d)-C.YEAR_START).days
        LIBRES[_i,_k]=True; vac[_i,_k]=True
VAC_REAL=vac&~LIBRES   # vacaciones de verdad (sin findes libres ni días no disponibles)
CLOSED=C.CLOSED; SPECIAL=C.SPECIAL
closed=np.array([d in CLOSED for d in date])
avail=(~vac)&(~closed[None,:])
# Días rojos (festivos no laborables / cierres): solo ahí se puede romper la regla de noches mínimas
_RJ=os.path.join(DATA,'rojos.json')   # rojos revisados por Diego (solo festivos no laborables y cierres reales)
if os.path.exists(_RJ):
    ROJO=np.zeros((P,D),bool)
    for _n,_d in json.load(open(_RJ)): ROJO[people.index(_n),(dt.date.fromisoformat(_d)-C.YEAR_START).days]=True
    ROJO|=closed[None,:]
else:
    ROJO=np.array([[fl[p][i]=='FFFF0000' for i in range(D)] for p in people])|closed[None,:]
ROJO_PINTAR=ROJO.copy()   # lo que se pinta en rojo en Excel y calendarios
for _x in getattr(C,'NIGHT_OFF',()): ROJO[:,(_x-C.YEAR_START).days]=True
SATN={d for d in range(D) if wd[d]==5 and any(a<=date[d]<=b for a,b in C.SAT_NIGHT_PERIODS)}
JA=np.array([{k.split(' (')[0]:v for k,v in summ.items()}[n]['ja'] for n in people],float)
def load_roles():
    R=json.load(open(os.path.join(WORK,'roles.json')))
    role=np.empty((P,D),dtype='U1')
    for i in range(P):
        for d in range(D): role[i,d]=R['roles'][str(i)][wk(d)]
    for _mon,_m in getattr(C,'ROLE_OVERRIDE',{}).items():
        _w=wk((_mon-C.YEAR_START).days)
        for _n,_r in _m.items():
            for d in range(D):
                if wk(d)==_w: role[IDX[_n],d]=_r
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
