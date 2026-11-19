"""Paso 7. Calendario de una persona (imagen + Excel), con la leyenda pactada con Diego:
  python calendario_persona.py propuesta_v5.csv "DIEGO ARMANDO" v5
Leyenda: verde mañana (borde amarillo sábado/domingo), lila tarde, naranja noche (NV viernes),
naranja claro sábado noche, gris oscuro vacaciones, rojo festivo no laborable/cierre,
recuadro rojo festivo laborable (29 ene, 5 mar, 23 abr, 1 may, 15 ago)."""
import sys, os, json, datetime as dt, calendar, openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from collections import Counter
import common as CM
FROM_XLSX=sys.argv[1].lower().endswith('.xlsx')   # calendario desde un Excel con hoja 'Rotación anual' (p. ej. el revisado por el encargado)
who=sys.argv[2]; TAG=sys.argv[3] if len(sys.argv)>3 else ''
if not FROM_XLSX: role=CM.load_roles(); x=CM.load_solution(sys.argv[1])
PREFIX='Calendario 2027' if FROM_XLSX else 'Propuesta 2027'
i0=CM.IDX[who]; code_lbl=[k for k in CM.summ][i0].split('(')[1].strip(')')
nice=who.title().replace('Diego Armando','Diego')
SPECIAL=CM.SPECIAL
FEST_LAB={dt.date(2027,1,29),dt.date(2027,3,5),dt.date(2027,4,23),dt.date(2027,5,1),dt.date(2027,8,15)}
KIND={'M':'M','MS':'MW','MD':'MW','C':'M','T':'T','TS':'TW','TD':'TW','N':'N','NV':'N','NS':'NS'}
D={}; hrs=0
if FROM_XLSX:
    _wb=openpyxl.load_workbook(sys.argv[1]); _ws=_wb['Rotación anual']
    _r=[r for r in range(4,19) if str(_ws.cell(r,1).value).strip().upper()==who][0]
    for d in range(CM.D):
        assert _ws.cell(2,4+d).value=='LMXJVSD'[CM.wd[d]]
        c=_ws.cell(_r,4+d); v=c.value.strip().upper() if isinstance(c.value,str) and c.value.strip() else None
        f=c.fill; fc=None
        if f is not None and f.fill_type=='solid': fc=f.fgColor.rgb if f.fgColor.type=='rgb' else f'theme{f.fgColor.theme}'
        dd=CM.date[d]; w=CM.wd[d]
        if v in KIND: hrs+=7 if (v in ('N','NV','NS') and (w==6 or dd in SPECIAL)) else 8
        else: v=None
        if v: k=KIND[v]
        elif fc=='theme1': k='vac'
        elif fc=='FFFF0000' or CM.closed[d]: k='rojo'
        elif fc=='FF00B0F0': k='azul'
        else: k=None
        D[dd]=dict(code=v,kind=k,fest_lab=dd in FEST_LAB and k!='rojo')
for d in ([] if FROM_XLSX else range(CM.D)):
    dd=CM.date[d]; w=CM.wd[d]; v=None
    if x[i0,d]:
        r=role[i0,d]
        v={'M':'MS' if w==5 else 'MD' if w==6 else ('C' if who=='JUANMA' else 'M'),'T':'TS' if w==5 else 'TD' if w==6 else 'T','N':'NV' if w==4 else 'NS' if w==5 else 'N'}[r]
        hrs+=7 if (r=='N' and (w==6 or dd in SPECIAL)) else 8
    if CM.vac[i0,d]: k='vac'
    elif v: k=KIND[v]
    elif CM.ROJO_PINTAR[i0,d]: k='rojo'
    else: k=None
    D[dd]=dict(code=v,kind=k,fest_lab=dd in FEST_LAB and k!='rojo')
P=dict(title=f'{who} ({code_lbl})',tt=f'{nice} ({code_lbl})'+(f' · {TAG}' if TAG else ''),hours=hrs,sab=sum(1 for dd,i in D.items() if dd.weekday()==5 and i['code']),ja=int(CM.JA[i0]),out=f'{nice.replace(" ","_")}_{TAG or "propuesta"}')
n=sum(1 for i in D.values() if i['code']); sun=sum(1 for d,i in D.items() if i['code']=='N' and d.weekday()==6)
spec=sum(1 for d in SPECIAL if D[d]['code'] in('N','NV','NS'))
print('turnos',n,'horas',P['hours'])
COL={'M':'A9D18E','MW':'A9D18E','T':'C9A7E0','TW':'C9A7E0','N':'E97132','NS':'F4B183','vac':'404040','azul':'00B0F0','rojo':'FF0000',None:'F2F2F2'}
NAME={'M':'Mañana','MW':'Mañana sábado/domingo','T':'Tarde','TW':'Tarde sábado/domingo','N':'Noche','NS':'Noche sábado','vac':'Vacaciones','azul':'Día azul (sin turno)','rojo':'Festivo no laborable',None:'Libre'}
LEGTXT={'M':'Mañana (M)','MW':'Mañana sábado/domingo (MS, MD)','T':'Tarde (T)','TW':'Tarde sábado/domingo (TS, TD)','N':'Noche (N, NV: viernes noche)','NS':'Noche sábado (NS)','vac':'Vacaciones','azul':'Día marcado en azul (sin turno)','rojo':'Día festivo no laborable','fl':'Festivo laborable'}
used=Counter(i['kind'] for i in D.values())
order=[k for k in ('M','MW','T','TW','N','NS','vac','azul','rojo') if used[k]]+['fl']
MESES=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
DIAS=['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo']
YEL,RED='FFFF00','C00000'
wb=openpyxl.Workbook(); ws=wb.active; ws.title='Calendario 2027'
ws['A1']=f'{PREFIX.upper()} — {P["title"]}'; ws['A1'].font=Font(bold=True,size=16)
ws['A2']='Propuesta de cuadrante 2027 (optimizada): mismas vacaciones que el cuadrante original, misma leyenda que los calendarios anteriores.'; ws['A2'].font=Font(italic=True,size=9)
def box(c,s): return Border(left=s,right=s,top=s,bottom=s)
for i,k in enumerate(order):
    r=4+i; ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=3); ws.cell(r,4,LEGTXT[k])
    for cc in (1,2,3):
        c=ws.cell(r,cc)
        if k=='fl': s=Side(style='thick',color=RED); c.border=Border(left=s if cc==1 else None,right=s if cc==3 else None,top=s,bottom=s)
        else:
            c.fill=PatternFill('solid',fgColor=COL[k])
            if k in('MW','TW'): s=Side(style='medium',color=YEL); c.border=Border(left=s if cc==1 else None,right=s if cc==3 else None,top=s,bottom=s)
thin=Side(style='thin',color='BBBBBB'); R0=4+len(order)+1
for m in range(1,13):
    br=(m-1)//3; bc=(m-1)%3; r0=R0+br*10; c0=1+bc*8
    t=ws.cell(r0,c0,MESES[m-1].upper()); t.font=Font(bold=True,size=12); ws.merge_cells(start_row=r0,start_column=c0,end_row=r0,end_column=c0+6)
    for i,l in enumerate('LMXJVSD'):
        c=ws.cell(r0+1,c0+i,l); c.font=Font(bold=True,color='C00000' if i>=5 else '000000'); c.alignment=Alignment(horizontal='center')
    for ri,week in enumerate(calendar.monthcalendar(2027,m)):
        for ci,day in enumerate(week):
            if not day: continue
            i=D[dt.date(2027,m,day)]; c=ws.cell(r0+2+ri,c0+ci,day); c.alignment=Alignment(horizontal='center')
            c.fill=PatternFill('solid',fgColor=COL[i['kind']]); c.font=Font(bold=True,color='FFFFFF' if i['kind']=='vac' else '000000')
            if i['kind'] in('MW','TW'): c.border=box(c,Side(style='medium',color=YEL))
            elif i['fest_lab']: c.border=box(c,Side(style='thick',color=RED))
            else: c.border=box(c,thin)
for c in range(1,24): ws.column_dimensions[get_column_letter(c)].width=4.6 if (c%8)!=0 else 1.6
wd=wb.create_sheet('Detalle'); wd.append(['Fecha','Día','Código','Turno','Mes','Observación'])
for c in wd[1]: c.font=Font(bold=True)
for d in sorted(D):
    i=D[d]; wd.append([d,DIAS[d.weekday()],i['code'] or '',NAME[i['kind']],d.month,'Festivo laborable' if i['fest_lab'] else ''])
    wd.cell(wd.max_row,1).number_format='dd/mm/yyyy'; wd.cell(wd.max_row,4).fill=PatternFill('solid',fgColor=COL[i['kind']])
    if i['kind']=='vac': wd.cell(wd.max_row,4).font=Font(color='FFFFFF')
for col,w in zip('ABCDEF',(12,12,8,26,6,20)): wd.column_dimensions[col].width=w
wd.freeze_panes='A2'
wr=wb.create_sheet('Resumen'); wr['A1']=f'RESUMEN 2027 — {P["title"]}'; wr['A1'].font=Font(bold=True,size=14)
turn=[k for k in ('M','MW','T','TW','N','NS') if used[k]]
hdr=['Mes']+[NAME[k] for k in turn]+['Días de turno','Vacaciones (días)','Días azules','Festivos no laborables','Festivos laborables']
for i,h in enumerate(hdr): c=wr.cell(3,1+i,h); c.font=Font(bold=True); c.alignment=Alignment(wrap_text=True,horizontal='center')
nt=len(turn)
for m in range(1,13):
    r=3+m; wr.cell(r,1,MESES[m-1])
    for j,k in enumerate(turn): wr.cell(r,2+j,f'=COUNTIFS(Detalle!$E:$E,{m},Detalle!$D:$D,"{NAME[k]}")')
    wr.cell(r,2+nt,f'=SUM(B{r}:{get_column_letter(1+nt)}{r})')
    for j,nm in enumerate(['Vacaciones','Día azul (sin turno)','Festivo no laborable']): wr.cell(r,3+nt+j,f'=COUNTIFS(Detalle!$E:$E,{m},Detalle!$D:$D,"{nm}")')
    wr.cell(r,6+nt,f'=COUNTIFS(Detalle!$E:$E,{m},Detalle!$F:$F,"Festivo laborable")')
wr.cell(16,1,'TOTAL AÑO').font=Font(bold=True)
for c in range(2,7+nt): wr.cell(16,c,f'=SUM({get_column_letter(c)}4:{get_column_letter(c)}15)').font=Font(bold=True)
wr['A18']=f'{PREFIX} ({P["title"]}): {P["sab"]} sábados trabajados, {P["hours"]:,} horas previstas, jornada anual {P["ja"]:,} h.'.replace(',','.')
wr.column_dimensions['A'].width=14
for c in range(2,8+nt): wr.column_dimensions[get_column_letter(c)].width=16
wb.save(os.path.join(CM.OUT,f'Calendario_Laboral_2027_{P["out"]}.xlsx'))
fig=plt.figure(figsize=(22,18.5)); gs=fig.add_gridspec(4,3,hspace=0.35,wspace=0.12,top=0.92,bottom=0.15,left=0.03,right=0.97)
for m in range(1,13):
    ax=fig.add_subplot(gs[(m-1)//3,(m-1)%3]); mat=calendar.monthcalendar(2027,m)
    ax.set_xlim(0,7); ax.set_ylim(0,len(mat)+1.1); ax.invert_yaxis(); ax.axis('off'); ax.set_title(MESES[m-1],fontsize=20,fontweight='bold',pad=6)
    for c,l in enumerate('LMXJVSD'): ax.text(c+.5,.5,l,ha='center',va='center',fontsize=12,fontweight='bold',color='#B00020' if c>=5 else '#333')
    for r,week in enumerate(mat):
        for c,day in enumerate(week):
            if not day: continue
            i=D[dt.date(2027,m,day)]; y=r+1
            ax.add_patch(Rectangle((c,y),1,1,facecolor='#'+COL[i['kind']],edgecolor='#BBBBBB',linewidth=.6))
            tc='white' if i['kind']=='vac' else '#111'
            ax.text(c+.08,y+.2,str(day),ha='left',va='center',fontsize=12,fontweight='bold',color=tc)
            if i['code']: ax.text(c+.5,y+.68,i['code'],ha='center',va='center',fontsize=10.5,color=tc)
            if i['kind'] in('MW','TW'): ax.add_patch(Rectangle((c+.04,y+.04),.92,.92,facecolor='none',edgecolor='#'+YEL,linewidth=3.5))
            elif i['fest_lab']: ax.add_patch(Rectangle((c+.04,y+.04),.92,.92,facecolor='none',edgecolor='#'+RED,linewidth=3.5))
fig.suptitle(f'{PREFIX} — {P["tt"]}',fontsize=28,fontweight='bold',y=0.975)
h=[]
for k in order:
    if k=='fl': h.append(Rectangle((0,0),1,1,facecolor='white',edgecolor='#'+RED,linewidth=3,label=LEGTXT[k]))
    elif k in('MW','TW'): h.append(Rectangle((0,0),1,1,facecolor='#'+COL[k],edgecolor='#'+YEL,linewidth=3,label=LEGTXT[k]))
    else: h.append(Rectangle((0,0),1,1,facecolor='#'+COL[k],edgecolor='#999',label=LEGTXT[k]))
fig.legend(handles=h,loc='lower center',ncol=4,fontsize=13,frameon=False,bbox_to_anchor=(0.5,0.035))
work=sum(used[k] for k in turn)
fig.text(0.5,0.012,f'{work} días de turno · {used["vac"]} días de vacaciones · {used["rojo"]} festivos y cierres · {P["hours"]:,} h '.replace(',','.') + ('(revisado)' if FROM_XLSX else 'previstas (propuesta)'),ha='center',fontsize=12,color='#444')
plt.savefig(os.path.join(CM.OUT,f'Calendario_2027_{P["out"]}.png'),dpi=130,facecolor='white')
print(used)
