"""Paso 3 (alternativa recomendada). Mismo modelo que model.py, pero en CP-SAT (OR-Tools),
resuelto para el año entero de una vez, con varios hilos y LNS interno.
  python solve_cpsat.py --inicio propuesta_v5.csv --minutos 30 --salida propuesta_v6.csv
Diferencias respecto a model.py:
  - El descanso de 1 día se penaliza también cuando el día libre es forzado (p. ej. sábado sin sábado noche),
    para que el patrón viernes noche + sábado libre + domingo noche cueste de verdad (en el MILP no se penalizaba).
  - Penaliza las tandas de noche de 6 o más (regla 13: tandas de 4-5).
  - El mínimo semanal (regla 12) es duro si se pasa --minimo-duro."""
import argparse, time, numpy as np
from ortools.sat.python import cp_model
import model as M
from common import *
W=dict(C.W)
W.setdefault('night_run6', 20)
ap=argparse.ArgumentParser(); ap.add_argument('--inicio',default=None); ap.add_argument('--minutos',type=float,default=20)
ap.add_argument('--salida',default='propuesta_cpsat.csv'); ap.add_argument('--hilos',type=int,default=8)
ap.add_argument('--minimo-duro',action='store_true'); ap.add_argument('--limites-duros',action='store_true'); ap.add_argument('--libre',default=None,help='desde,hasta,turnos (p.ej. 2027-11-29,2027-12-31,MT): solo se reoptimiza eso; el resto queda fijo al --inicio'); ap.add_argument('--semilla',type=int,default=1); ap.add_argument('--mantener-equidad',default=None,help='CSV de referencia: cada persona queda a ±EQUIDAD_TOL de sus sábados, viernes, sábados y domingos noche'); ap.add_argument('--coste-cambio',type=int,default=400)
a=ap.parse_args()
role=M.role; A2=M.avail2; ISDAY=M.ISDAY
m=cp_model.CpModel(); obj=[]
# Reoptimización parcial (--libre): las celdas fuera de la ventana quedan fijas y las reglas duras
# solo se imponen donde toca alguna celda libre (lo fijo se respeta tal cual, aunque incumpla algo).
FREE=np.ones((P,D),bool)
if a.libre:
    _f,_t,_g=a.libre.split(','); f0=(dt.date.fromisoformat(_f)-start).days; f1=(dt.date.fromisoformat(_t)-start).days
    FREE[:]=False
    for i in range(P):
        for d in range(f0,f1+1):
            if role[i,d] in _g: FREE[i,d]=True
def T(i,ds): return any(FREE[i,d] for d in ds if 0<=d<D)
x=[[m.NewBoolVar('') if A2[i,d] else m.NewConstant(0) for d in range(D)] for i in range(P)]
def B(): return m.NewBoolVar('')
# ---------- duras
for i in range(P):
    for d in range(D):
        if M.FORCED[i,d]: m.Add(x[i][d]==1)
for i in range(P):
    for t in range(D-6):
        if T(i,range(t,t+7)): m.Add(sum(x[i][t+k] for k in range(7))<=6)
    for d in range(1,D):
        p,q=role[i,d-1],role[i,d]
        if p!=q and ((p=='N' and q in 'MT') or (p=='T' and q=='M')) and T(i,(d-1,d)): m.Add(x[i][d-1]+x[i][d]<=1)
    for d in range(D-1):
        if wd[d]==5 and ISDAY[i,d] and ISDAY[i,d+1] and A2[i,d] and A2[i,d+1] and T(i,(d,d+1)): m.Add(x[i][d]==x[i][d+1])
    if T(i,range(D)):
        hrs=sum(int(M.HRS[i,d])*x[i][d] for d in range(D) if A2[i,d])
        if a.libre:   # celdas fijas trabajadas aunque el modelo no las permita (p. ej. en vacaciones): cuentan horas
            x0_=load_solution(a.inicio); hrs=hrs+int(sum(M.HRS[i,d] for d in range(D) if not A2[i,d] and x0_[i,d] and not FREE[i,d]))
        m.Add(hrs>=int(JA[i])); m.Add(hrs<=int(JA[i])+C.HOURS_TOL)
# ---------- noches: al menos MIN_NIGHT_RUN seguidas (solo se rompe junto a un día rojo o al principio/fin del año)
NIGHT=lambda i,d: x[i][d] if (0<=d<D and role[i,d]=='N') else 0
for i in range(P):
    for d in range(D):
        if role[i,d]!='N' or not A2[i,d]: continue
        if d==0 or d==D-1 or ROJO[i,d-1] or ROJO[i,d+1]: continue
        if not T(i,(d-1,d,d+1)): continue
        m.Add(x[i][d]<=NIGHT(i,d-1)+NIGHT(i,d+1))      # nada de noches sueltas
        if C.MIN_NIGHT_RUN>=3 and d+2<D and role[i,d+1]=='N' and A2[i,d+1]:
            # tanda de exactamente 2 noches: penalizada (mejor 3 o más)
            t2=B(); m.Add(t2>=x[i][d]+x[i][d+1]-NIGHT(i,d-1)-NIGHT(i,d+2)-1); obj.append(W.get('night_run2',40)*t2)
# ---------- fines de semana marcados con 1 sola persona de mañana y de tarde (duro)
for _dd in getattr(C,'WEEKEND_ONE',()):
    d=(_dd-start).days
    for c in 'MT':
        ps=[x[i][d] for i in range(P) if role[i,d]==c and A2[i,d]]
        if ps: m.Add(sum(ps)<=1)
# ---------- cobertura
cl={'M':0,'T':1,'N':2}
for c in 'MTN':
    k=cl[c]
    for d in range(D):
        ps=[x[i][d] for i in range(P) if role[i,d]==c and A2[i,d]]
        s=sum(ps) if ps else 0
        t=int(M.TGT[k,d]); cap=int(M.CAP[k,d])
        if t>0:
            u=m.NewIntVar(0,t,''); m.Add(u>=t-s); u2=m.NewIntVar(0,t,''); m.Add(u2>=u-1)
            obj+= [int(M.W1[k,d])*u, int(M.W1[k,d])*u2]
        o=m.NewIntVar(0,P,''); m.Add(o>=s-cap); obj.append(int(M.WO[k,d])*o)
# ---------- por persona
Z={i:[] for i in range(P)}; S6={i:[] for i in range(P)}
for i in range(P):
    for d in range(1,D-1):
        if vac[i,d]: continue
        if not (A2[i,d-1] and A2[i,d+1]): continue
        z=B(); m.Add(z>=x[i][d-1]+x[i][d+1]-x[i][d]-1)
        w=W['single_rest_fri_sun_night'] if (wd[d]==5 and d not in SATN and role[i,d]=='N') else W['single_rest']
        obj.append(w*z); Z[i].append(z)
        if A2[i,d]:
            iw=B(); m.Add(iw>=x[i][d]-x[i][d-1]-x[i][d+1]); obj.append(W['isolated_work']*iw)
    for t in range(D-5):
        if all(A2[i,t+k] for k in range(6)):
            s6=B(); m.Add(sum(x[i][t+k] for k in range(6))-5<=s6); obj.append(W['six_days']*s6); S6[i].append(s6)
            if all(role[i,t+k]=='N' for k in range(6)): obj.append(W['night_run6']*s6)
    for d in range(D):
        if A2[i,d]:
            st=B(); m.Add(st>=x[i][d]-(x[i][d-1] if d>0 else 0))
            obj.append((W['run_start_night'] if role[i,d]=='N' else W['run_start_day'])*st)
    for w in range(NW):
        mn=int(M.WEEKMIN[i,w])
        if mn>0:
            s=sum(x[i][d] for d in range(D) if wk(d)==w and A2[i,d])
            if a.minimo_duro and T(i,[d for d in range(D) if wk(d)==w]): m.Add(s>=mn)
            else:
                v=m.NewIntVar(0,7,''); m.Add(v>=mn-s); obj.append(W['week_min']*v)
# ---------- límites por persona (rachas de 6 y descansos de 1 día)
for i in range(P):
    for lst,lim in ((S6[i],C.MAX_RACHAS6),(Z[i],C.MAX_DESCANSOS_1)):
        if a.limites_duros: m.Add(sum(lst)<=lim)
        else:
            e=m.NewIntVar(0,len(lst),''); m.Add(e>=sum(lst)-lim); obj.append(W.get('limit_excess',400)*e)
# ---------- fines de semana seguidos (sábado o domingo trabajado)
WEND={}
for i in range(P):
    ks=[]
    for d in range(D):
        if wd[d]!=5: continue
        ds=[k for k in (d,d+1) if k<D and A2[i,k]]
        if not ds: ks.append(0); continue
        w=B()
        for k in ds: m.Add(w>=x[i][k])
        m.Add(w<=sum(x[i][k] for k in ds)); ks.append(w)
    WEND[i]=ks
    lim=getattr(C,'MAX_FINDES_SEGUIDOS',None)
    if lim:
        for k in range(len(ks)-lim):
            win=ks[k:k+lim+1]
            if all(isinstance(v,int) for v in win): continue
            if people[i] in C.G_NIGHT:
                t=B(); m.Add(sum(win)-lim<=t); obj.append(W.get('three_weekends_night',150)*t)
            elif T(i,[d for d in range(D) if wd[d] in (5,6)]):
                m.Add(sum(win)<=lim)
# ---------- mantener la equidad de la versión de partida
if a.mantener_equidad:
    x0e=load_solution(a.mantener_equidad)
    def cnt(i,kind,v):
        out=[]
        for d in range(D):
            if not A2[i,d] and v is not x0e: continue
            r=role[i,d]; w_=wd[d]
            ok=(kind=='SAB' and w_==5 and r in 'MT') or (kind=='NV' and w_==4 and r=='N') or (kind=='NS' and w_==5 and r=='N') or (kind=='DN' and w_==6 and r=='N')
            if ok: out.append(v[i][d] if v is x else int(v[i,d]))
        return out
    for i in range(P):
        for kind in ('SAB','NV','NS','DN'):
            c0=sum(cnt(i,kind,x0e)); vs=cnt(i,kind,x)
            if c0==0 and not vs: continue
            lo,hi=C.EQUIDAD_EXC.get((people[i],kind),(-C.EQUIDAD_TOL,C.EQUIDAD_TOL))
            e=sum(vs); m.Add(e>=max(0,c0+lo)); m.Add(e<=c0+hi)
# ---------- equidad
def spread(exprs,weight):
    mx=m.NewIntVar(0,400,''); mn=m.NewIntVar(0,400,'')
    for e in exprs: m.Add(mx>=e); m.Add(mn<=e)
    obj.append(weight*(mx-mn))
for G in M.GROUPS:
    ex=[]
    for i in G:
        ent=[]
        for d in range(D):
            if wd[d]==5 and A2[i,d]: ent.append(x[i][d])
            elif wd[d]==6 and date[d-1] in CLOSED and A2[i,d] and ISDAY[i,d]: ent.append(x[i][d])
        ex.append(sum(ent))
    spread(ex,W['fair_weekends'])
spread([sum(x[i][d] for d in range(D) if role[i,d]=='N' and A2[i,d]) for i in M.G2],W['fair_rot_nights'])
for i in range(P):
    wes=[]
    for d in range(D):
        if wd[d]==5:
            ent=[x[i][d]] if A2[i,d] else []
            if not A2[i,d] and d+1<D and A2[i,d+1] and ISDAY[i,d+1]: ent=[x[i][d+1]]
            wes.append((date[d].month,ent))
    for k in range(len(wes)-1):
        e1,e2=wes[k][1],wes[k+1][1]
        if e1 and e2:
            v=B(); m.Add(sum(e1+e2)-1<=v); obj.append(W['consecutive_weekends']*v)
    for mth in range(1,13):
        ent=[e for mm,es in wes if mm==mth for e in es]
        if len(ent)>2:
            v=m.NewIntVar(0,5,''); m.Add(sum(ent)-2<=v); obj.append(W['month_3_weekends']*v)
m.Minimize(sum(obj))
if a.libre:
    x0=load_solution(a.inicio)
    for i in range(P):          # cambiar lo mínimo respecto al punto de partida
        for d in range(D):
            if FREE[i,d] and A2[i,d]: obj.append(W.get('cambio',int(getattr(a,'coste_cambio',400)))*(x[i][d] if x0[i,d]==0 else 1-x[i][d]))
    m.Minimize(sum(obj))
    nfix=0
    for i in range(P):
        for d in range(D):
            if A2[i,d] and not FREE[i,d]: m.Add(x[i][d]==int(x0[i,d])); nfix+=1
    print('celdas fijas',nfix,flush=True)
if a.inicio:
    x0=load_solution(a.inicio)
    for i in range(P):
        for d in range(D):
            if A2[i,d]: m.AddHint(x[i][d],int(x0[i,d]))
s=cp_model.CpSolver(); s.parameters.max_time_in_seconds=a.minutos*60; s.parameters.num_workers=a.hilos
s.parameters.random_seed=a.semilla; s.parameters.log_search_progress=False
class CB(cp_model.CpSolverSolutionCallback):
    def __init__(s_): super().__init__(); s_.t0=time.time()
    def on_solution_callback(s_):
        print(f'  {round(time.time()-s_.t0)} s  objetivo {round(s_.ObjectiveValue())}  cota {round(s_.BestObjectiveBound())}',flush=True)
        sol=np.array([[int(s_.Value(x[i][d])) for d in range(D)] for i in range(P)]); save_solution(sol,a.salida)
st=s.Solve(m,CB())
print('FIN',s.StatusName(st),'objetivo',s.ObjectiveValue(),'cota',s.BestObjectiveBound(),'->',a.salida)
