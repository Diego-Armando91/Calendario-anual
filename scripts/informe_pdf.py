"""Informe PDF de un cuadrante editado a mano (lee los códigos de la hoja 'Propuesta anual').
  python informe_pdf.py ../data/Propuesta_Cuadrante_2027_v13_Diego.xlsx propuesta_v13.csv "v13 editada"
Todo se calcula a partir de los códigos de cada celda (no de la rotación), así que vale aunque se cambien turnos a mano."""
import sys, os, collections, datetime as dt, openpyxl, numpy as np
import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt
from common import *
import model as M
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image, KeepTogether

SRC=sys.argv[1]; BASE=sys.argv[2] if len(sys.argv)>2 else None; TAG=sys.argv[3] if len(sys.argv)>3 else 'editada'
CL={'M':'M','MS':'M','MD':'M','C':'C','T':'T','TS':'T','TD':'T','N':'N','NV':'N','NS':'N'}
DIAS=['L','M','X','J','V','S','D']; MESES=['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic']
f=lambda d: f"{date[d].day} {MESES[date[d].month-1]}"
nice=lambda p: p.title().replace('Diego Armando','Diego').replace('Ruben S','Rubén S').replace('Sofia','Sofía')

# ---------- lectura
ws=openpyxl.load_workbook(SRC)['Propuesta anual']
rows={str(ws.cell(r,1).value).strip().upper():r for r in range(1,ws.max_row+1) if str(ws.cell(r,1).value).strip().upper() in people}
K=np.full((P,D),'',dtype=object); V=np.zeros((P,D),bool)
for i,p in enumerate(people):
    for d in range(D):
        v=ws.cell(rows[p],4+d).value; v=v.strip().upper() if isinstance(v,str) else ''
        if v=='V': V[i,d]=True
        elif v in CL: K[i,d]=v
S=np.vectorize(lambda v: CL.get(v,''))(K)           # clase de turno M/C/T/N
W_=(S!='')                                            # trabaja
DAY=np.isin(S,['M','C','T'])
J=IDX['JUANMA']
HOL={(date[d]) for d in range(D) if closed[d]}
def hrs(i): return sum(7 if (S[i,d]=='N' and (wd[d]==6 or date[d] in SPECIAL)) else 8 for d in range(D) if W_[i,d])

# ---------- métricas por persona
VER0=(C.VERANO[0]-start).days; VER1=(C.VERANO[1]-start).days
sats=[d for d in range(D) if wd[d]==5]
per={}; viol=[]
for i,p in enumerate(people):
    r=dict()
    r['dias']=int(W_[i].sum()); r['horas']=hrs(i); r['ja']=int(JA[i]); r['dif']=r['horas']-r['ja']
    for c in 'MCTN': r[c]=int((S[i]==c).sum())
    r['vac']=int(V[i].sum())
    fr=[d for d in range(D) if wd[d]==4 and W_[i,d]]
    r['vie']=len(fr); r['vie_por']={c:[d for d in fr if S[i,d]==c] for c in 'MCTN'}
    r['sab']=sum(1 for d in sats if W_[i,d]); r['dom']=sum(1 for d in range(D) if wd[d]==6 and W_[i,d])
    r['sab_MT']=sum(1 for d in sats if DAY[i,d]); r['NS']=sum(1 for d in sats if S[i,d]=='N')
    r['DN']=sum(1 for d in range(D) if wd[d]==6 and S[i,d]=='N'); r['NV']=len(r['vie_por']['N'])
    wk_=[bool(W_[i,d] or (d+1<D and W_[i,d+1])) for d in sats]
    r['findes']=sum(wk_); run=mx=0; pares=0
    for k,v in enumerate(wk_):
        run=run+1 if v else 0; mx=max(mx,run)
        if k and v and wk_[k-1]: pares+=1
    r['max_findes']=mx; r['pares']=pares
    mon=collections.Counter(date[d].month for k,d in enumerate(sats) if wk_[k]); r['meses3']=sum(1 for v in mon.values() if v>=3)
    # fines de semana "de mañana/tarde o sábado noche" (sin domingos noche) y verano
    wk2=[bool(DAY[i,d] or (d+1<D and DAY[i,d+1]) or S[i,d]=='N') for d in sats]
    r['findes_sin_DN']=sum(wk2)
    vs=[k for k,d in enumerate(sats) if VER0<=d<VER1]
    r['verano']=sum(wk2[k] for k in vs); r['verano_seg']=sum(1 for a,b in zip(vs,vs[1:]) if wk2[a] and wk2[b])
    # descansos de 1 día, rachas
    s1=[d for d in range(1,D-1) if not W_[i,d] and not V[i,d] and W_[i,d-1] and W_[i,d+1]]
    r['desc1']=len(s1); r['desc1_f']=s1
    runs=[]; d=0
    while d<D:
        if W_[i,d]:
            j=d
            while j<D and W_[i,j]: j+=1
            runs.append((d,j-d)); d=j
        else: d+=1
    r['rachas6']=sum(1 for _,L in runs if L>=6); r['max_racha']=max(L for _,L in runs)
    for d0,L in runs:
        if L>=7: viol.append((p,'%d días seguidos'%L,f"{f(d0)} – {f(d0+L-1)}"))
    nr=[]; d=0
    while d<D:
        if S[i,d]=='N':
            j=d
            while j<D and S[i,j]=='N': j+=1
            nr.append(j-d); d=j
        else: d+=1
    r['tandas_noche']=collections.Counter(nr)
    # semanas
    r['sem6']=0; r['bajo_min']=[]; r['mezcla']=[]
    for w in range(NW):
        dd=[d for d in range(D) if wk(d)==w]; n=int(W_[i,dd].sum())
        if len(dd)==7 and not V[i,dd].any() and n>=6: r['sem6']+=1
        if n<M.WEEKMIN[i,w] and not V[i,dd].all(): r['bajo_min'].append((w,n,int(M.WEEKMIN[i,w])))
        cls={('D' if S[i,d] in 'MC' else S[i,d]) for d in dd if W_[i,d]}
        if len(cls)>1: r['mezcla'].append(w)
    # reglas duras
    for d in range(1,D):
        a,b=S[i,d-1],S[i,d]
        if a=='N' and b in ('M','C','T'): viol.append((p,'de noche a %s sin descanso'%b,f(d)))
        if a=='T' and b in ('M','C'): viol.append((p,'de tarde a mañana sin descanso',f(d)))
    for d in range(D):
        if W_[i,d] and (V[i,d] or closed[d]): viol.append((p,'trabaja en vacaciones/cierre',f(d)))
        if S[i,d]=='N' and wd[d]==5 and d not in SATN: viol.append((p,'sábado noche fuera de periodo',f(d)))
        if S[i,d]=='N' and 0<d<D-1 and S[i,d-1]!='N' and S[i,d+1]!='N' and not ROJO[i,d-1] and not ROJO[i,d+1]: viol.append((p,'noche suelta',f(d)))
        if wd[d]==5 and d in SATN and S[i,d]=='N' and d+1<D and S[i,d+1]!='N' and not ROJO[i,d+1] and not V[i,d+1]: viol.append((p,'sábado noche sin domingo noche',f(d)))
    if not (0<=r['dif']<=C.HOURS_TOL): viol.append((p,'horas fuera de JA..JA+8',f"{r['dif']:+d} h"))
    per[p]=r

# ---------- cobertura
cov=dict(m7=[],t=[],vie=collections.Counter(),vie_bajo=[],dom=collections.Counter(),dom1=[],sab0=[],findes_bajo=[])
for d in range(D):
    m7=int(((S[:,d]=='M')).sum()); t=int((S[:,d]=='T').sum()); n=int((S[:,d]=='N').sum())
    if wd[d]<5 and not closed[d] and date[d] not in C.DAY_OFF and date[d] not in SPECIAL:
        if m7<2: cov['m7'].append((d,m7))
        if t<2: cov['t'].append((d,t))
    if wd[d]==4 and not closed[d]:
        cov['vie'][n]+=1
        if n<M.TGT[2,d]: cov['vie_bajo'].append((d,n,int(M.TGT[2,d])))
    if wd[d]==6 and not closed[d]:
        cov['dom'][n]+=1
        if n<2: cov['dom1'].append((d,n))
    if d in SATN and n==0: cov['sab0'].append(d)
    if wd[d]>=5:
        for k,c in ((0,'M'),(1,'T')):
            x=int(np.isin(S[:,d],[c] if c=='T' else ['M','C']).sum())
            if x<M.TGT[k,d]: cov['findes_bajo'].append((d,c,x,int(M.TGT[k,d])))
# cambios frente a la base
cambios=collections.Counter()
if BASE:
    xb=load_solution(BASE)
    for i,p in enumerate(people):
        for d in range(D):
            if bool(xb[i,d])!=bool(W_[i,d]) or (W_[i,d] and S[i,d] not in (M.role[i,d],'C' if M.role[i,d]=='M' else '')): cambios[p]+=1

# ---------- PDF
OUTD=os.path.join(OUT,f'Informe_Cuadrante_2027_{TAG.replace(" ","_")}.pdf'); IMG=os.path.join(WORK,'_inf'); os.makedirs(IMG,exist_ok=True)
ss=getSampleStyleSheet()
H1=ParagraphStyle('h1',parent=ss['Heading1'],fontSize=16,spaceAfter=6,textColor=colors.HexColor('#1F3864'))
H2=ParagraphStyle('h2',parent=ss['Heading2'],fontSize=12.5,spaceBefore=8,spaceAfter=4,textColor=colors.HexColor('#2E5597'))
B=ParagraphStyle('b',parent=ss['BodyText'],fontSize=9.2,leading=12.5)
SM=ParagraphStyle('sm',parent=B,fontSize=7.6,leading=9.6)
def tbl(data,widths=None,fs=7.8,hl=None,zebra=True):
    t=Table(data,colWidths=widths,repeatRows=1)
    st=[('FONT',(0,0),(-1,-1),'Helvetica',fs),('FONT',(0,0),(-1,0),'Helvetica-Bold',fs),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#DCE6F1')),
        ('GRID',(0,0),(-1,-1),0.25,colors.HexColor('#B0B0B0')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('ALIGN',(1,1),(-1,-1),'CENTER'),
        ('TOPPADDING',(0,0),(-1,-1),1.5),('BOTTOMPADDING',(0,0),(-1,-1),1.5)]
    if zebra:
        for r in range(1,len(data)):
            if r%2==0: st.append(('BACKGROUND',(0,r),(-1,r),colors.HexColor('#F5F7FA')))
    for (c,r,col) in (hl or []): st.append(('BACKGROUND',(c,r),(c,r),colors.HexColor(col)))
    t.setStyle(TableStyle(st)); return t
RED='#F8CBAD'; GRN='#C6EFCE'
story=[]
story.append(Paragraph(f'Análisis del cuadrante 2027 — {TAG}',H1))
story.append(Paragraph(f'Fichero analizado: <i>{os.path.basename(SRC)}</i>. Todo se calcula leyendo los códigos de cada celda de la hoja «Propuesta anual» '
    '(M/MS/MD mañana, C central, T/TS/TD tarde, N/NV/NS noche, V vacaciones). Domingo noche y noches especiales cuentan 7 h.',B))
nv=len(viol)
res=[f'<b>Reglas duras incumplidas:</b> {nv}.',
     f'<b>Horas fuera de JA…JA+8:</b> {sum(1 for p in people if not 0<=per[p]["dif"]<=C.HOURS_TOL)} personas.',
     f'<b>Semanas por debajo del mínimo:</b> {sum(len(per[p]["bajo_min"]) for p in people)}.',
     f'<b>Días entre semana con menos de 2 a las 7:00</b> (sin contar la central): {len(cov["m7"])}; <b>tarde con menos de 2:</b> {len(cov["t"])}.',
     f'<b>Viernes noche por debajo del objetivo:</b> {len(cov["vie_bajo"])}; <b>domingos noche con 1:</b> {len(cov["dom1"])}; <b>sábados noche sin nadie:</b> {len(cov["sab0"])}.',
     f'<b>Fines de semana de mañana/tarde por debajo de lo previsto:</b> {len(cov["findes_bajo"])} días.',
     f'<b>Descansos de 1 día:</b> {sum(per[p]["desc1"] for p in people)} en total (máx. {max(per[p]["desc1"] for p in people)} por persona). '
     f'<b>Rachas de 6 días:</b> {sum(per[p]["rachas6"] for p in people)}. <b>Fines de semana seguidos (pares):</b> {sum(per[p]["pares"] for p in people)}.']
if BASE: res.append(f'<b>Celdas cambiadas a mano respecto a la propuesta generada:</b> {sum(cambios.values())} '+'('+', '.join(f'{nice(p)} {n}' for p,n in cambios.most_common() if n)+').')
story.append(Paragraph('Resumen',H2))
for l in res: story.append(Paragraph('• '+l,B))
# reglas duras
story.append(Paragraph('Reglas duras incumplidas',H2))
if viol: story.append(tbl([['Persona','Qué pasa','Cuándo']]+[[nice(p),q,w] for p,q,w in viol],[4*cm,7*cm,5*cm]))
else: story.append(Paragraph('Ninguna. ✔',B))
# tabla general
story.append(PageBreak()); story.append(Paragraph('Resumen por persona',H2))
hdr=['Persona','Días','Horas','JA','Dif','Mañ.','Cent.','Tarde','Noche','Vac.','Sáb.','Dom.','Findes','Findes\nseg.\nmáx','Pares\nseg.','Desc.\n1 día','Rach.\n6','Sem.\n6 d.','Bajo\nmín.']
data=[hdr]; hl=[]
for k,p in enumerate(people,1):
    r=per[p]; data.append([nice(p),r['dias'],r['horas'],r['ja'],f"{r['dif']:+d}",r['M'],r['C'],r['T'],r['N'],r['vac'],r['sab'],r['dom'],r['findes'],r['max_findes'],r['pares'],r['desc1'],r['rachas6'],r['sem6'],len(r['bajo_min'])])
    if not 0<=r['dif']<=C.HOURS_TOL: hl.append((4,k,RED))
    if r['desc1']>C.MAX_DESCANSOS_1: hl.append((15,k,RED))
    if r['rachas6']>C.MAX_RACHAS6: hl.append((16,k,RED))
    if r['max_findes']>C.MAX_FINDES_SEGUIDOS and p not in C.G_NIGHT: hl.append((13,k,RED))
    if r['bajo_min']: hl.append((18,k,RED))
story.append(tbl(data,[2.5*cm]+[0.86*cm]*18,fs=6.4,hl=hl))
story.append(Paragraph(f'En rojo lo que no cumple: horas fuera de JA…JA+{C.HOURS_TOL}, más de {C.MAX_DESCANSOS_1} descansos de 1 día, más de {C.MAX_RACHAS6} rachas de 6, '
    f'más de {C.MAX_FINDES_SEGUIDOS} fines de semana seguidos (día y rotativos) o semanas bajo el mínimo. «Findes» = sábado o domingo trabajado en cualquier turno.',SM))
# noches
story.append(Paragraph('Noches de fin de semana y viernes',H2))
data=[['Persona','Viernes\nnoche','Sábado\nnoche','Domingo\nnoche','Tandas de noche (longitud: nº)']]
for p in people:
    r=per[p]
    if r['N']: data.append([nice(p),r['NV'],r['NS'],r['DN'],', '.join(f'{L}:{n}' for L,n in sorted(r['tandas_noche'].items()))])
story.append(tbl(data,[3*cm,1.8*cm,1.8*cm,1.8*cm,9*cm]))
# viernes
story.append(PageBreak()); story.append(Paragraph('Viernes trabajados por persona y turno',H2))
data=[['Persona','Total','Mañana','Central','Tarde','Noche (NV)']]
for p in people:
    r=per[p]; data.append([nice(p),r['vie']]+[len(r['vie_por'][c]) for c in 'MCTN'])
story.append(tbl(data,[3.2*cm]+[2*cm]*5))
story.append(Spacer(1,6)); story.append(Paragraph('Detalle de fechas',H2))
for p in people:
    r=per[p]; parts=[]
    for c,lab in (('M','Mañana'),('C','Central'),('T','Tarde'),('N','Noche')):
        if r['vie_por'][c]: parts.append(f"<b>{lab} ({len(r['vie_por'][c])}):</b> "+', '.join(f(d) for d in r['vie_por'][c]))
    story.append(Paragraph(f"<b>{nice(p)}</b> — {r['vie']} viernes. "+' · '.join(parts),SM))
# cobertura
story.append(PageBreak()); story.append(Paragraph('Cobertura',H2))
cl=[f"<b>Viernes noche</b> (personas: nº de viernes): "+', '.join(f'{k}: {v}' for k,v in sorted(cov['vie'].items()))+'. Por debajo del objetivo: '+(', '.join(f"{f(d)} ({n} de {t})" for d,n,t in cov['vie_bajo']) or 'ninguno')+'.',
    f"<b>Domingo noche</b> (personas: nº): "+', '.join(f'{k}: {v}' for k,v in sorted(cov['dom'].items()))+'. Con 1 persona: '+(', '.join(f(d) for d,_ in cov['dom1']) or 'ninguno')+'.',
    f"<b>Sábados noche sin nadie</b> (dentro del periodo): "+(', '.join(f(d) for d in cov['sab0']) or 'ninguno')+'.',
    f"<b>Entre semana con menos de 2 a las 7:00</b> (sin la central): "+(', '.join(f"{DIAS[wd[d]]} {f(d)} ({n})" for d,n in cov['m7']) or 'ninguno')+'.',
    f"<b>Tarde entre semana con menos de 2:</b> "+(', '.join(f"{DIAS[wd[d]]} {f(d)} ({n})" for d,n in cov['t']) or 'ninguno')+'.',
    f"<b>Fines de semana de mañana/tarde por debajo:</b> "+(', '.join(f"{DIAS[wd[d]]} {f(d)} {c} ({x} de {t})" for d,c,x,t in cov['findes_bajo']) or 'ninguno')+'.']
for l in cl: story.append(Paragraph('• '+l,B))
bm=[(p,w,n,m) for p in people for (w,n,m) in per[p]['bajo_min']]
if bm:
    def wr(w):
        dd=[d for d in range(D) if wk(d)==w]; return f"{f(dd[0])} – {f(dd[-1])}"
    story.append(Paragraph('<b>Semanas por debajo del mínimo:</b> '+'; '.join(f"{nice(p)} {wr(w)} ({n} de {m})" for p,w,n,m in bm)+'.',B))
mz=[(p,w) for p in people for w in per[p]['mezcla']]
if mz:
    def wr2(w):
        dd=[d for d in range(D) if wk(d)==w]; return f"{f(dd[0])}"
    story.append(Paragraph('<b>Semanas con dos turnos distintos</b> (mañana/tarde/noche mezclados en la misma semana): '+'; '.join(f"{nice(p)} sem. del {wr2(w)}" for p,w in mz)+'.',B))

# ---------- comparativas
def bars(names,series,labels,fn,title):
    x=np.arange(len(names)); wdt=0.8/len(series)
    fig,ax=plt.subplots(figsize=(9.5,3.6))
    pal=['#4472C4','#ED7D31','#A5A5A5','#FFC000','#5B9BD5','#70AD47']
    for k,(s,l) in enumerate(zip(series,labels)):
        b=ax.bar(x+k*wdt-0.4+wdt/2,s,wdt,label=l,color=pal[k%len(pal)])
        for rect,v in zip(b,s): ax.text(rect.get_x()+rect.get_width()/2,rect.get_height()+0.2,str(v),ha='center',fontsize=7.5)
    ax.set_xticks(x); ax.set_xticklabels(names); ax.legend(fontsize=7.5,ncol=len(series),loc='upper center',bbox_to_anchor=(0.5,-0.1),frameon=False)
    ax.spines[['top','right']].set_visible(False); ax.set_title(title,fontsize=10,pad=6)
    plt.tight_layout(); p=os.path.join(IMG,fn); plt.savefig(p,dpi=170); plt.close(); return p
def comp_table(group):
    keys=[('Días trabajados','dias'),('Horas (dif. JA)',None),('Vacaciones (días V)','vac'),('Mañana','M'),('Tarde','T'),('Noche','N'),
          ('Viernes trabajados','vie'),('  viernes de mañana',lambda r: len(r['vie_por']['M'])),('  viernes de tarde',lambda r: len(r['vie_por']['T'])),('  viernes noche (NV)','NV'),
          ('Sábados (cualquier turno)','sab'),('Domingos (cualquier turno)','dom'),('Sábados de mañana/tarde','sab_MT'),('Sábado noche','NS'),('Domingo noche','DN'),
          ('Fines de semana trabajados','findes'),('Findes sin contar domingo noche','findes_sin_DN'),('Findes en verano (29 may–29 ago)','verano'),('Findes seguidos (pares)','pares'),
          ('Máx. findes seguidos','max_findes'),('Descansos de 1 día','desc1'),('Rachas de 6 días','rachas6')]
    data=[['']+[nice(p) for p in group]+['Diferencia\nmáx–mín']]; hl=[]
    for k,(lab,key) in enumerate(keys,1):
        if key is None: vals=[per[p]['dif'] for p in group]; row=[f"{per[p]['horas']} ({v:+d})" for p,v in zip(group,vals)]
        else:
            vals=[(key(per[p]) if callable(key) else per[p][key]) for p in group]; row=vals
        sp=max(vals)-min(vals); data.append([lab]+list(row)+[sp])
        if key not in (None,'vac','dias') and sp>=3: hl.append((len(group)+1,k,RED))
    return tbl(data,[5.2*cm]+[2.3*cm]*len(group)+[2.2*cm],fs=7.8,hl=hl)
def razones(group):
    out=[]
    for p in group:
        i=IDX[p]; r=per[p]
        # semanas de cada turno disponibles (sin vacaciones)
        wk_by=collections.Counter()
        for w in range(NW):
            dd=[d for d in range(D) if wk(d)==w and not V[i,d] and not closed[d]]
            cls=collections.Counter(S[i,d] for d in dd if W_[i,d])
            if cls: wk_by[cls.most_common(1)[0][0]]+=1
        vacm=collections.Counter(date[d].month for d in range(D) if V[i,d])
        meses=', '.join(f"{MESES[m-1]} ({n})" for m,n in sorted(vacm.items()) if n>=3)
        out.append(f"<b>{nice(p)}:</b> {wk_by.get('M',0)+wk_by.get('C',0)} semanas de mañana, {wk_by.get('T',0)} de tarde y {wk_by.get('N',0)} de noche trabajadas; "
                   f"{r['vac']} días de vacaciones/libres (sobre todo en {meses or '—'}).")
    return out
for title,group,intro in (('Comparativa: Bea y Denia (rotan mañana ↔ tarde)',['BEA','DENIA'],
        'Bea y Denia forman pareja: cuando una está de mañana, la otra está de tarde. Lo lógico es que acaben con números parecidos; '
        'las diferencias suelen venir de cuándo cae cada una de vacaciones (el bloque de la otra se alarga) y de los cambios de semana para reforzar la mañana.'),
     ('Comparativa: Diego, Esther, David y Lorena (rotan mañana ↔ noche)',['DIEGO ARMANDO','ESTHER','DAVID P','LORENA'],
        'Los cuatro alternan mañana y noche por parejas (Diego–Lorena y Esther–David). Deberían quedar igualados en noches, viernes noche, sábados noche y fines de semana.')):
    story.append(PageBreak()); story.append(Paragraph(title,H1)); story.append(Paragraph(intro,B)); story.append(Spacer(1,4))
    story.append(comp_table(group))
    names=[nice(p) for p in group]
    p1=bars(names,[[len(per[p]['vie_por']['M']) for p in group],[len(per[p]['vie_por']['T']) for p in group],[per[p]['NV'] for p in group]],['Viernes mañana','Viernes tarde','Viernes noche'],f'v_{group[0]}.png','Viernes por turno')
    p2=bars(names,[[per[p]['sab_MT'] for p in group],[per[p]['NS'] for p in group],[per[p]['DN'] for p in group],[per[p]['verano'] for p in group]],['Sábados mañana/tarde','Sábados noche','Domingos noche','Findes en verano'],f'f_{group[0]}.png','Fines de semana')
    story.append(Spacer(1,4)); story.append(Image(p1,width=17*cm,height=6.4*cm)); story.append(Image(p2,width=17*cm,height=6.4*cm))
    story.append(Paragraph('Posibles motivos de las diferencias',H2))
    for l in razones(group): story.append(Paragraph('• '+l,B))
    # diferencias destacadas automáticas
    notes=[]
    def spread(key,lab):
        vals={p:(key(per[p]) if callable(key) else per[p][key]) for p in group}
        mx=max(vals,key=vals.get); mn=min(vals,key=vals.get)
        if vals[mx]-vals[mn]>=3: notes.append(f"{lab}: {nice(mx)} {vals[mx]} frente a {nice(mn)} {vals[mn]} (diferencia {vals[mx]-vals[mn]}).")
    for key,lab in (('N','Noches'),('NV','Viernes noche'),('NS','Sábados noche'),('DN','Domingos noche'),('vie','Viernes trabajados'),('findes','Fines de semana'),('sab_MT','Sábados de mañana/tarde'),('verano','Findes en verano'),('desc1','Descansos de 1 día')):
        spread(key,lab)
    if notes:
        story.append(Paragraph('Diferencias más marcadas (3 o más):',B))
        for n in notes: story.append(Paragraph('– '+n,B))
    story.append(Paragraph('Cómo leerlo: quien tiene más semanas de noche hace más viernes y domingos noche; quien tiene más semanas de mañana hace más sábados de mañana. '
        'Las vacaciones largas en verano reducen los fines de semana de verano, y los cambios de semana hechos a mano desplazan noches o mañanas de una persona a su pareja.',SM))
doc=SimpleDocTemplate(OUTD,pagesize=A4,leftMargin=1.4*cm,rightMargin=1.4*cm,topMargin=1.3*cm,bottomMargin=1.3*cm,title=f'Análisis cuadrante 2027 {TAG}')
def foot(c,d_):
    c.setFont('Helvetica',7); c.setFillColor(colors.grey); c.drawRightString(A4[0]-1.4*cm,0.8*cm,f'Cuadrante 2027 · {TAG} · pág. {d_.page}')
doc.build(story,onFirstPage=foot,onLaterPages=foot)
import json
json.dump({p:{k:(v if not isinstance(v,(dict,collections.Counter)) else {str(a):(b if isinstance(b,int) else len(b)) for a,b in v.items()}) for k,v in r.items() if k not in('desc1_f','bajo_min','mezcla')} for p,r in per.items()},open(os.path.join(WORK,'_informe.json'),'w'),default=str,ensure_ascii=False)
print('ok',OUTD,'violaciones',len(viol))
