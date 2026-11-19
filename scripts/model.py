"""Modelo MILP (scipy.optimize.milp / HiGHS). Variables x[i,d]=1 si la persona i trabaja el día d
(en el turno de su semana, fijado por roles.json). Todo lo demás son variables de holgura penalizadas.
Duro: vacaciones y cierres, máx. 6 días seguidos, sin pasar de noche a mañana/tarde ni de tarde a mañana
al día siguiente, sábado=domingo en turnos de día, horas anuales entre JA y JA+HOURS_TOL.
Blando (pesos en config.W): cobertura, descansos de 1 día, rachas de 6, inicios de tanda, equidad,
fines de semana seguidos / tercer finde del mes, mínimo de días por semana."""
import numpy as np, scipy.sparse as sp, time
from scipy.optimize import milp, LinearConstraint, Bounds
from common import *
W=C.W
role=load_roles(); ISDAY=(role!='N')
# ---- demanda (personas por turno y día)
CUR={}
CLm={'M':'M','MS':'M','MD':'M','C':'M','T':'T','TS':'T','TD':'T','N':'N','NV':'N','NS':'N'}
for d in range(D):
    c={'M':0,'T':0,'N':0}
    for n in people:
        v=seq[n][d]
        if v: c[CLm[v]]+=1
    CUR[d]=c
OVER={(x-start).days for x in C.OVERRIDE_DATES}
cl={'M':0,'T':1,'N':2}
TGT=np.zeros((3,D),int); CAP=np.zeros((3,D),int); W1=np.zeros((3,D)); WO=np.zeros((3,D))
for d in range(D):
    w=wd[d]
    for c in 'MTN':
        k=cl[c]
        if d in OVER: t=CUR[d][c]; cap=t+1; a=W['under_weekday']; o=W['over_weekday']
        elif w<=3: t=C.TGT_WEEKDAY; cap=C.CAP_WEEKDAY; a=W['under_weekday']; o=W['over_weekday']
        elif w==4:
            if c=='N': t=C.FRI_NIGHT_WITH_SAT if (d+1) in SATN else C.FRI_NIGHT_NO_SAT; cap=t+1; a=W['under_fri_night']; o=W['over_weekday']
            else: t=C.TGT_WEEKDAY; cap=C.CAP_WEEKDAY; a=W['under_weekday']; o=W['over_weekday']
        else:
            if c=='N': t=(C.SAT_NIGHT if d in SATN else 0) if w==5 else C.SUN_NIGHT
            else: t=1 if date[d]<C.WEEKEND_TWO_FROM else 2
            cap=t; a=W['under_weekend']; o=W['over_weekend']
        TGT[k,d]=t;CAP[k,d]=cap;W1[k,d]=a;WO[k,d]=o
avail2=avail.copy()
for d in range(D):
    if date[d] in SPECIAL: avail2[:,d]&=(role[:,d]=='N')
    if wd[d]==5 and d not in SATN: avail2[:,d]&=(role[:,d]!='N')
HRS=np.full((P,D),8.0)
for i in range(P):
    for d in range(D):
        if role[i,d]=='N' and (wd[d]==6 or date[d] in SPECIAL): HRS[i,d]=7.0
GROUPS=[[IDX[n] for n in g] for g in (C.G_DAY,C.G_ROT,C.G_NIGHT)]
G2=GROUPS[1]
# mínimo de días por semana
HOLWEEK={wk(d) for d in range(D) if date[d] in CLOSED or date[d] in SPECIAL}
WEEKMIN=np.zeros((P,NW),int)
for i in range(P):
    for w in range(NW):
        days=[d for d in range(D) if wk(d)==w]
        a=sum(1 for d in days if not vac[i,d] and not closed[d])
        if w in HOLWEEK: m=min(C.WEEK_MIN_HOLIDAY,int(a*4/7)) if a<7 else C.WEEK_MIN_HOLIDAY
        else: m=C.WEEK_MIN_NORMAL if (a==7 and len(days)==7) else int(a*4/7)
        WEEKMIN[i,w]=m
def build():
    nx=P*D; ix=lambda i,d:i*D+d
    bz=nx; bs=bz+P*D; bw=bs+P*D; bst=bw+P*D; bc=bst+P*D; bf=bc+3*D*3; bwk=bf+8; NWE=60; bmo=bwk+P*NWE; bwm=bmo+P*12; n=bwm+P*NW
    c=np.zeros(n); lb=np.zeros(n); ub=np.ones(n); integ=np.zeros(n); integ[:nx]=1
    for i in range(P):
        for d in range(D):
            if not avail2[i,d]: ub[ix(i,d)]=0
    ub[bs:bs+P*D]=10; ub[bc:bf]=10; ub[bf:bwk]=400; ub[bwm:n]=7
    c[bz:bz+P*D]=W['single_rest']; c[bs:bs+P*D]=W['six_days']; c[bw:bw+P*D]=W['isolated_work']
    for i in range(P):
        for d in range(D):
            c[bst+i*D+d]=W['run_start_night'] if role[i,d]=='N' else W['run_start_day']
            if wd[d]==5 and d not in SATN and role[i,d]=='N': c[bz+i*D+d]=W['single_rest_fri_sun_night']
    rows=[];cols=[];vals=[];rl=[];ru=[];r=[0]; hrows=[]
    def add(ent,lo,hi):
        for j,v in ent: rows.append(r[0]); cols.append(j); vals.append(v)
        rl.append(lo); ru.append(hi); r[0]+=1
    INF=1e9
    for k,cc in enumerate('MTN'):
        for d in range(D):
            ps=[i for i in range(P) if role[i,d]==cc and avail2[i,d]]
            u1=bc+(k*D+d)*3; u2=u1+1; o=u1+2
            c[u1]=W1[k,d]; c[u2]=W1[k,d]; c[o]=WO[k,d]
            add([(ix(i,d),1) for i in ps]+[(u1,1)],TGT[k,d],INF)
            add([(u1,1),(u2,-1)],-INF,1)
            add([(ix(i,d),1) for i in ps]+[(o,-1)],-INF,CAP[k,d])
    for i in range(P):
        for d in range(1,D-1):
            if avail2[i,d] and (avail2[i,d-1] or avail2[i,d+1]):
                add([(bz+i*D+d,1),(ix(i,d-1),-1),(ix(i,d+1),-1),(ix(i,d),1)],-1,INF)
            if avail2[i,d]:
                add([(bw+i*D+d,1),(ix(i,d),-1),(ix(i,d-1),1),(ix(i,d+1),1)],0,INF)
        for t in range(D-6): add([(ix(i,t+k),1) for k in range(7)],-INF,6)
        for d in range(D):
            if avail2[i,d]: add([(bst+i*D+d,1),(ix(i,d),-1)]+([(ix(i,d-1),1)] if d>0 else []),0,INF)
        for t in range(D-5): add([(ix(i,t+k),1) for k in range(6)]+[(bs+i*D+t,-1)],-INF,5)
        for d in range(1,D):
            a,b=role[i,d-1],role[i,d]
            if a!=b and ((a=='N' and b in 'MT') or (a=='T' and b=='M')): add([(ix(i,d-1),1),(ix(i,d),1)],-INF,1)
        for d in range(D-1):
            if wd[d]==5 and ISDAY[i,d] and ISDAY[i,d+1] and avail2[i,d] and avail2[i,d+1]: add([(ix(i,d),1),(ix(i,d+1),-1)],0,0)
        hrows.append(r[0]); add([(ix(i,d),HRS[i,d]) for d in range(D) if avail2[i,d]],JA[i],JA[i]+C.HOURS_TOL)
        for w in range(NW):
            if WEEKMIN[i,w]>0:
                v=bwm+i*NW+w; c[v]=W['week_min']
                add([(ix(i,d),1) for d in range(D) if wk(d)==w and avail2[i,d]]+[(v,1)],WEEKMIN[i,w],INF)
    for g,G in enumerate(GROUPS):
        mx=bf+2*g; mn=mx+1; c[mx]=W['fair_weekends']; c[mn]=-W['fair_weekends']
        for i in G:
            ent=[]
            for d in range(D):
                if wd[d]==5 and avail2[i,d]: ent.append((ix(i,d),1))
                elif wd[d]==6 and date[d-1] in CLOSED and avail2[i,d] and ISDAY[i,d]: ent.append((ix(i,d),1))
            add(ent+[(mx,-1)],-INF,0); add(ent+[(mn,-1)],0,INF)
    mx=bf+6; mn=bf+7; c[mx]=W['fair_rot_nights']; c[mn]=-W['fair_rot_nights']
    for i in G2:
        ent=[(ix(i,d),1) for d in range(D) if role[i,d]=='N' and avail2[i,d]]
        add(ent+[(mx,-1)],-INF,0); add(ent+[(mn,-1)],0,INF)
    for i in range(P):
        wes=[]
        for d in range(D):
            if wd[d]==5:
                ent=[(ix(i,d),1)] if avail2[i,d] else []
                if not avail2[i,d] and d+1<D and avail2[i,d+1] and ISDAY[i,d+1]: ent=[(ix(i,d+1),1)]
                wes.append((date[d].month,ent))
        for k in range(len(wes)-1):
            e1,e2=wes[k][1],wes[k+1][1]
            if e1 and e2:
                v=bwk+i*NWE+k; c[v]=W['consecutive_weekends']; add(e1+e2+[(v,-1)],-INF,1)
        for mth in range(1,13):
            ent=[e for mm,es in wes if mm==mth for e in es]
            if ent:
                v=bmo+i*12+mth-1; c[v]=W['month_3_weekends']; add(ent+[(v,-1)],-INF,2)
    A=sp.csr_matrix((vals,(rows,cols)),shape=(r[0],n))
    return dict(hrows=hrows,c=c,A=A,rl=np.array(rl),ru=np.array(ru),lb=lb,ub=ub,integ=integ,nx=nx,n=n)
def run_milp(pb,lb,ub,rl=None,ru=None,tl=30,gap=0.005):
    return milp(pb['c'],constraints=LinearConstraint(pb['A'],pb['rl'] if rl is None else rl,pb['ru'] if ru is None else ru),
                integrality=pb['integ'],bounds=Bounds(lb,ub),options=dict(time_limit=tl,disp=False,mip_rel_gap=gap))
def fix_outside(pb,x,free_mask):
    """Fija x fuera de free_mask (P x D bool) a la solución actual."""
    nx=pb['nx']; lb=pb['lb'].copy(); ub=pb['ub'].copy()
    idx=np.where((~free_mask).ravel() & (pb['ub'][:nx]>0))[0]
    lb[idx]=x.ravel()[idx]; ub[idx]=x.ravel()[idx]
    return lb,ub
