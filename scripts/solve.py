"""Paso 3. Genera la propuesta.
  python solve.py --inicio rolling --minutos 25 --salida propuesta_v5.csv
  python solve.py --inicio propuesta_v4.csv --minutos 15 --salida propuesta_v5.csv   (refinar una existente)
Fases: (a) solución inicial por horizonte rodante (ventanas de 6 semanas, horas a ritmo proporcional);
       (b) reparación del grupo de noche con todo el año libre (opcional, --reparar-noche);
       (c) LNS: se libera un grupo (noche / mañana / tarde) durante 9 semanas al azar, el resto fijo, y se
           acepta si mejora el objetivo. Es lo que más mejora la calidad: dale tiempo."""
import argparse, time, numpy as np
import model as M
from common import *
ap=argparse.ArgumentParser(); ap.add_argument('--inicio',default='rolling'); ap.add_argument('--minutos',type=float,default=20)
ap.add_argument('--salida',default='propuesta_nueva.csv'); ap.add_argument('--reparar-noche',action='store_true'); ap.add_argument('--semilla',type=int,default=1)
a=ap.parse_args()
t0=time.time(); pb=M.build(); nx=pb['nx']; print('modelo',pb['n'],'variables',flush=True)
def rolling():
    AV=M.avail2.astype(float); cum=np.cumsum(AV,axis=1); tot=AV.sum(1)
    x=np.zeros((P,D),int); Wn=42; step=21; s=0
    while s<D:
        e=min(D,s+Wn); last=(e==D)
        free=np.zeros((P,D),bool); free[:,s:e]=True
        lb,ub=M.fix_outside(pb,x,free); rl=pb['rl'].copy(); ru=pb['ru'].copy()
        if not last:
            for k,row in enumerate(pb['hrows']):
                t=JA[k]*cum[k,e-1]/tot[k]; rl[row]=t-24; ru[row]=t+24
        r=M.run_milp(pb,lb,ub,rl,ru,tl=60,gap=0.01)
        if r.x is None:   # última ventana: ampliar hacia atrás
            s=max(0,s-14); x[:,s:]=0; print('  reintento desde',s,flush=True); continue
        xn=np.rint(r.x[:nx]).reshape(P,D).astype(int); keep=e if last else s+step
        x[:,s:keep]=xn[:,s:keep]; print('  ventana',s,e,'obj',round(r.fun),flush=True)
        if last: break
        s+=step
    return x
x=rolling() if a.inicio=='rolling' else load_solution(a.inicio)
save_solution(x,a.salida)
G={'N':[IDX[n] for n in ('ARMANDO','JUANMA V','RUBEN S','ESTHER','DAVID P','DIEGO ARMANDO','LORENA')],
   'M':[IDX[n] for n in ('CHEMA','JUANMA','BEA','DENIA','ESTHER','DAVID P','DIEGO ARMANDO','LORENA')],
   'T':[IDX[n] for n in ('MICHEL','ALFREDO','IRENE','SOFIA','BEA','DENIA')]}
if a.reparar_noche:
    free=np.zeros((P,D),bool); free[G['N'],:]=True
    lb,ub=M.fix_outside(pb,x,free); r=M.run_milp(pb,lb,ub,tl=420,gap=0.01)
    if r.x is not None: x=np.rint(r.x[:nx]).reshape(P,D).astype(int); print('reparación noche',round(r.fun),flush=True)
free=np.zeros((P,D),bool); lb,ub=M.fix_outside(pb,x,free); r=M.run_milp(pb,lb,ub,tl=60)
if r.x is None: raise SystemExit('La solución de partida no cumple las restricciones duras (¿horas?). Usa --inicio rolling.')
best=r.fun; print('objetivo inicial',round(best),flush=True)
rng=np.random.default_rng(a.semilla); it=0; order=['N','M','T']
while time.time()-t0<a.minutos*60:
    g=order[it%3]; it+=1; s=int(rng.integers(-20,D-20))
    free=np.zeros((P,D),bool); free[np.ix_(G[g],range(max(0,s),min(D,s+63)))]=True
    lb,ub=M.fix_outside(pb,x,free); r=M.run_milp(pb,lb,ub,tl=25)
    if r.x is not None and r.fun<best-1e-6:
        best=r.fun; x=np.rint(r.x[:nx]).reshape(P,D).astype(int); save_solution(x,a.salida)
        print(f'  it {it} grupo {g} desde {s}: objetivo {round(best)}  ({round(time.time()-t0)} s)',flush=True)
save_solution(x,a.salida); print('FIN objetivo',round(best),'->',a.salida)
