"""Paso 2. Asigna el turno de cada persona por semana (lunes-domingo).
Fijos: config.FIX. Parejas (config.PAIRS): alternan su turno en bloques de ~5 semanas.
Elige la longitud de los bloques (4-6, preferido 5) para igualar las noches/tardes disponibles
entre los dos de cada pareja según sus vacaciones y suavizar la oferta semanal por turno.
Guarda work/roles.json."""
import json, numpy as np
from common import *
def week_av(i):
    w=np.zeros(NW)
    for d in range(D): w[wk(d)]+=avail[i,d]
    return w
WAV=[week_av(i) for i in range(P)]
def seqs():
    out=[]
    def rec(cur,tot):
        if tot>=52: out.append(tuple(cur)); return
        for L in C.BLOCK_LENGTHS: rec(cur+[L],tot+L)
    rec([],0); return out
ALL=seqs()
def role_from(lengths,ar,br):
    ra=np.empty(NW,dtype='U1'); rb=np.empty(NW,dtype='U1'); w=1
    ra[0]=ar; rb[0]=br            # vie 1 - dom 3 ene: igual que el primer bloque (continuidad con 2026)
    for k,L in enumerate(lengths):
        for _ in range(L):
            if w>=NW: break
            ra[w]=ar if k%2==0 else br; rb[w]=br if k%2==0 else ar; w+=1
    return ra,rb
def supply(rolesw):
    s={c:np.zeros(NW) for c in 'MTN'}
    for i,r in rolesw.items():
        for c in 'MTN': s[c]+=np.where(r==c,WAV[i],0)
    return s
roles={}
for n,r in C.FIX.items(): roles[IDX[n]]=np.array([r]*NW)
for a,ar,b,br in C.PAIRS:
    ra,rb=role_from((5,)*11,ar,br); roles[IDX[a]]=ra; roles[IDX[b]]=rb
def score(L,a,ar,b,br,other):
    ra,rb=role_from(L,ar,br); cur=dict(other); cur[IDX[a]]=ra; cur[IDX[b]]=rb
    s=supply(cur); var=sum(((v[1:52]-v[1:52].mean())**2).mean() for v in s.values())
    NA=WAV[IDX[a]][ra==br].sum(); NB=WAV[IDX[b]][rb==br].sum(); MA=WAV[IDX[a]][ra==ar].sum(); MB=WAV[IDX[b]][rb==ar].sum()
    return abs(NA-NB)+0.5*abs(MA-MB)+0.6*var+4*sum(1 for x in L if x!=5)
cfg={}
for it in range(2):
    for a,ar,b,br in C.PAIRS:
        other={k:v for k,v in roles.items() if k not in (IDX[a],IDX[b])}
        best=min(ALL,key=lambda L:score(L,a,ar,b,br,other))
        ra,rb=role_from(best,ar,br); roles[IDX[a]]=ra; roles[IDX[b]]=rb; cfg[f'{a}/{b}']=list(best)
json.dump({'roles':{str(i):list(map(str,r)) for i,r in roles.items()},'bloques':cfg},open(os.path.join(WORK,'roles.json'),'w'),ensure_ascii=False,indent=0)
for k,v in cfg.items(): print(k,v)
