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
ap.add_argument('--minimo-duro',action='store_true'); ap.add_argument('--semilla',type=int,default=1)
a=ap.parse_args()
role=M.role; A2=M.avail2; ISDAY=M.ISDAY
m=cp_model.CpModel(); obj=[]
x=[[m.NewBoolVar('') if A2[i,d] else m.NewConstant(0) for d in range(D)] for i in range(P)]
def B(): return m.NewBoolVar('')
# ---------- duras
for i in range(P):
    for t in range(D-6): m.Add(sum(x[i][t+k] for k in range(7))<=6)
    for d in range(1,D):
        p,q=role[i,d-1],role[i,d]
        if p!=q and ((p=='N' and q in 'MT') or (p=='T' and q=='M')): m.Add(x[i][d-1]+x[i][d]<=1)
    for d in range(D-1):
        if wd[d]==5 and ISDAY[i,d] and ISDAY[i,d+1] and A2[i,d] and A2[i,d+1]: m.Add(x[i][d]==x[i][d+1])
    hrs=sum(int(M.HRS[i,d])*x[i][d] for d in range(D) if A2[i,d])
    m.Add(hrs>=int(JA[i])); m.Add(hrs<=int(JA[i])+C.HOURS_TOL)
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
for i in range(P):
    for d in range(1,D-1):
        if vac[i,d]: continue
        if not (A2[i,d-1] and A2[i,d+1]): continue
        z=B(); m.Add(z>=x[i][d-1]+x[i][d+1]-x[i][d]-1)
        w=W['single_rest_fri_sun_night'] if (wd[d]==5 and d not in SATN and role[i,d]=='N') else W['single_rest']
        obj.append(w*z)
        if A2[i,d]:
            iw=B(); m.Add(iw>=x[i][d]-x[i][d-1]-x[i][d+1]); obj.append(W['isolated_work']*iw)
    for t in range(D-5):
        if all(A2[i,t+k] for k in range(6)):
            s6=B(); m.Add(sum(x[i][t+k] for k in range(6))-5<=s6); obj.append(W['six_days']*s6)
            if all(role[i,t+k]=='N' for k in range(6)): obj.append(W['night_run6']*s6)
    for d in range(D):
        if A2[i,d]:
            st=B(); m.Add(st>=x[i][d]-(x[i][d-1] if d>0 else 0))
            obj.append((W['run_start_night'] if role[i,d]=='N' else W['run_start_day'])*st)
    for w in range(NW):
        mn=int(M.WEEKMIN[i,w])
        if mn>0:
            s=sum(x[i][d] for d in range(D) if wk(d)==w and A2[i,d])
            if a.minimo_duro: m.Add(s>=mn)
            else:
                v=m.NewIntVar(0,7,''); m.Add(v>=mn-s); obj.append(W['week_min']*v)
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
