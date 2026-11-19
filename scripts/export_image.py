"""Paso 6. Imagen comparativa: python export_image.py v5"""
import sys, os, pickle, collections
from paths import *
TAG=sys.argv[1] if len(sys.argv)>1 else 'nueva'
import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt, numpy as np
M=pickle.load(open(os.path.join(WORK,f'metricas_{TAG}.pkl'),'rb'))
per,perb,cov,covb,ws,wsb,nr,nrb=M['per'],M['perb'],M['cov'],M['covb'],M['ws'],M['wsb'],M['nr'],M['nrb']
A='#9AA5B1'; Bc='#2E7D5B'
def bars(a,labels,v0,v1,title,off=1.0):
    x=np.arange(len(labels)); b0=a.bar(x-0.2,v0,0.38,color=A,label='Actual'); b1=a.bar(x+0.2,v1,0.38,color=Bc,label='Propuesta '+TAG)
    for b in list(b0)+list(b1): a.text(b.get_x()+b.get_width()/2,b.get_height()+off,int(b.get_height()),ha='center',fontsize=15,fontweight='bold')
    a.set_xticks(x); a.set_xticklabels(labels,fontsize=12); a.set_ylim(0,max(max(v0),max(v1))*1.2+1); a.legend(fontsize=12,frameon=False)
    a.set_title(title,fontsize=17,fontweight='bold')
    for s in ('top','right'): a.spines[s].set_visible(False)
fig,ax=plt.subplots(2,2,figsize=(17,12)); fig.patch.set_facecolor('white')
fig.suptitle('Cuadrante 2027: actual vs propuesta '+TAG,fontsize=28,fontweight='bold',y=0.985)
S=lambda d,k: sum(v[k] for v in d.values())
bars(ax[0,0],['Descansos de\n1 solo día','Semanas L–D\ncon 6 días','Rachas de 6\ndías o más','Semanas con\nturno mezclado'],
     [S(perb,'simples'),S(perb,'sem6'),S(perb,'rachas6'),M['mixb']],[S(per,'simples'),S(per,'sem6'),S(per,'rachas6'),M['mix']],'Descansos y semanas (15 personas)',3)
ks=[1,2,3,4,5,6]
bars(ax[0,1],['1 noche','2','3','4','5','6'],[nrb[k] for k in ks],[nr[k] for k in ks],'Noches seguidas: nº de tandas según su longitud',2)
cb=collections.Counter(covb['N'][4]); cp=collections.Counter(cov['N'][4]); ks=[0,1,2,3,4]
bars(ax[1,0],['0 personas','1 persona','2 personas','3 personas','4 personas'],[cb[k] for k in ks],[cp[k] for k in ks],'Viernes noche: nº de viernes según personas',0.6)
bars(ax[1,1],['Fines de semana\nseguidos','Meses con 3 o más\nfines de semana','Domingos noche\ncon 2 personas'],
     [S(wsb,'seguidos'),S(wsb,'meses3'),collections.Counter(covb['N'][6])[2]],[S(ws,'seguidos'),S(ws,'meses3'),collections.Counter(cov['N'][6])[2]],'Fines de semana y domingo noche',1)
fig.text(0.5,0.01,'Mismas vacaciones y horas anuales · semanas lunes–domingo enteras del mismo turno · rotaciones en bloques de 5 semanas',ha='center',fontsize=12,color='#444')
plt.tight_layout(rect=(0,0.03,1,0.96)); plt.savefig(os.path.join(OUT,f'Propuesta_{TAG}_vs_Actual_2027.png'),dpi=140,facecolor='white')
