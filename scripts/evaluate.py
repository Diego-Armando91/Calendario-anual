"""Paso 4. Evalúa una propuesta frente al cuadrante actual.
  python evaluate.py propuesta_v5.csv"""
import sys, numpy as np, collections
from common import *
role=load_roles()
CLm={'M':'M','MS':'M','MD':'M','C':'M','T':'T','TS':'T','TD':'T','N':'N','NV':'N','NS':'N'}
def base_matrix():
    X=np.zeros((P,D),int); C=np.full((P,D),'',dtype=object)
    for i,n in enumerate(people):
        for d in range(D):
            v=seq[n][d]
            if v: X[i,d]=1; C[i,d]=CLm[v]
    return X,C
def sol_matrix(x):
    C=np.full((P,D),'',dtype=object)
    for i in range(P):
        for d in range(D):
            if x[i,d]: C[i,d]=role[i,d]
    return x,C
def metrics(X,C):
    vac_=vac
    per={}
    for i,n in enumerate(people):
        s=X[i]; single=0; run6=0; run_len=collections.Counter()
        d=0
        while d<D:
            if s[d]==0 and not vac_[i,d]:
                j=d
                while j<D and s[j]==0 and not vac_[i,j]: j+=1
                if d>0 and j<D and s[d-1]==1 and s[j]==1 and (j-d)==1: single+=1
                d=j
            else: d+=1
        d=0
        while d<D:
            if s[d]==1:
                j=d
                while j<D and s[j]==1: j+=1
                run_len[j-d]+=1; d=j
            else: d+=1
        # semanas lun-dom completas con 6 días
        w6=0; d0=3  # lunes 4 ene = índice 3
        while d0+6<D and d0+6<=D-1:
            seg=range(d0,d0+7)
            if not any(vac_[i,k] for k in seg) and sum(s[k] for k in seg)>=6: w6+=1
            d0+=7
        # fines de semana (sábado trabajado como día/noche S)
        wk=sum(1 for d in range(D) if wd[d]==5 and s[d]==1)
        per[n]=dict(dias=int(s.sum()),simples=single,rachas6=sum(v for k,v in run_len.items() if k>=6),sem6=w6,sab=wk)
    cov=collections.defaultdict(lambda: collections.defaultdict(list))
    for d in range(D):
        for c in 'MTN':
            cov[c][wd[d]].append(sum(1 for i in range(P) if X[i,d] and C[i,d]==c))
    return per,cov
def fri_dist(cov): return collections.Counter(cov['N'][4])
def night_runs(X,C):
    L=collections.Counter()
    for i in range(P):
        d=0
        while d<D:
            if X[i,d] and C[i,d]=='N':
                j=d
                while j<D and X[i,j] and C[i,j]=='N': j+=1
                L[j-d]+=1; d=j
            else: d+=1
    return L
def mixed_weeks(X,C):
    m=0
    for i in range(P):
        for w in range(NW):
            s={C[i,d] for d in range(D) if wk(d)==w and X[i,d]}
            if len(s)>1: m+=1
    return m
def weekend_stats(X):
    res={}
    for i,n in enumerate(people):
        wes=[]
        for d in range(D):
            if wd[d]==5:
                w=X[i,d] or (closed[d] and d+1<D and X[i,d+1] and role[i,d+1]!='N')
                wes.append((date[d].month,int(bool(w))))
        cons=sum(1 for k in range(len(wes)-1) if wes[k][1] and wes[k+1][1])
        mon=collections.Counter(); 
        for m,w in wes: mon[m]+=w
        res[n]=dict(total=sum(w for _,w in wes),seguidos=cons,meses3=sum(1 for v in mon.values() if v>=3),maxmes=max(mon.values()))
    return res

def week_below_min(x):
    import model as M
    return [(people[i],w,int(sum(x[i,d] for d in range(D) if wk(d)==w)),int(M.WEEKMIN[i,w])) for i in range(P) for w in range(NW) if sum(x[i,d] for d in range(D) if wk(d)==w)<M.WEEKMIN[i,w]]
def hours(x):
    return {n:int(sum((7 if (role[i,d]=='N' and (wd[d]==6 or date[d] in SPECIAL)) else 8) for d in range(D) if x[i,d])-JA[i]) for i,n in enumerate(people)}
def hard_violations(x):
    bad=[]
    for i in range(P):
        for t in range(D-6):
            if x[i,t:t+7].sum()>6: bad.append((people[i],'7 días seguidos',str(date[t])))
        for d in range(1,D):
            a,b=role[i,d-1],role[i,d]
            if x[i,d-1] and x[i,d] and ((a=='N' and b in 'MT') or (a=='T' and b=='M')): bad.append((people[i],'cambio de turno sin descanso',str(date[d])))
        for d in range(D):
            if x[i,d] and (vac[i,d] or closed[d]): bad.append((people[i],'trabaja en vacaciones/cierre',str(date[d])))
            if x[i,d] and role[i,d]=='N' and date[d] in C.NIGHT_OFF: bad.append((people[i],'noche antes de festivo',str(date[d])))
            if x[i,d] and role[i,d]!='N' and date[d] in C.DAY_OFF: bad.append((people[i],'mañana/tarde en festivo',str(date[d])))
            if date[d] in C.FIXED and role[i,d] in C.FIXED[date[d]] and bool(x[i,d])!=(people[i] in C.FIXED[date[d]][role[i,d]]): bad.append((people[i],'asignación fija',str(date[d])))
            if x[i,d] and role[i,d]=='N' and wd[d]==5 and d not in SATN: bad.append((people[i],'sábado noche fuera de periodo',str(date[d])))
    return bad
def report(x):
    X,Cc=sol_matrix(x); per,cov=metrics(X,Cc); ws=weekend_stats(X)
    Xb,Cb=base_matrix(); perb,covb=metrics(Xb,Cb); wsb=weekend_stats(Xb)
    S=lambda d,k: sum(v[k] for v in d.values())
    print('                         actual  propuesta')
    for lab,k in (('Descansos de 1 día','simples'),('Rachas >=6 días','rachas6'),('Semanas L-D con 6 días','sem6')): print(f'{lab:26s}{S(perb,k):6d}{S(per,k):10d}')
    print(f"{'Fines de semana seguidos':26s}{S(wsb,'seguidos'):6d}{S(ws,'seguidos'):10d}")
    print(f"{'Meses con 3+ findes':26s}{S(wsb,'meses3'):6d}{S(ws,'meses3'):10d}")
    print(f"{'Semanas turno mezclado':26s}{mixed_weeks(Xb,Cb):6d}{mixed_weeks(X,Cc):10d}")
    print('Tandas de noche (longitud:nº) actual',sorted(night_runs(Xb,Cb).items()),' propuesta',sorted(night_runs(X,Cc).items()))
    print('Viernes noche (personas:nº)',sorted(collections.Counter(cov['N'][4]).items()),' Domingo noche',sorted(collections.Counter(cov['N'][6]).items()),' Sábado noche',sorted(collections.Counter(cov['N'][5]).items()))
    vf=sum(1 for i in range(P) for d in range(1,D-1) if wd[d]==5 and x[i,d]==0 and not vac[i,d] and x[i,d-1] and x[i,d+1] and role[i,d-1]=='N' and role[i,d+1]=='N')
    print('Viernes noche + sábado libre + domingo noche:',vf)
    wb=week_below_min(x); print('Semanas por debajo del mínimo semanal:',len(wb), wb[:10])
    print('Horas sobre JA:',hours(x))
    print('Máx. por persona -> descansos de 1 día:',max(v['simples'] for v in per.values()),' rachas de 6:',max(v['rachas6'] for v in per.values()),
          ' | personas con >%d descansos de 1 día: %s' % (C.MAX_DESCANSOS_1,[n for n,v in per.items() if v['simples']>C.MAX_DESCANSOS_1]),
          ' | con >%d rachas de 6: %s' % (C.MAX_RACHAS6,[n for n,v in per.items() if v['rachas6']>C.MAX_RACHAS6]))
    print('Violaciones duras:',len(hard_violations(x)), hard_violations(x)[:5])
    for n in people: print('  ',n,per[n],ws[n])
if __name__=='__main__':
    report(load_solution(sys.argv[1]))
