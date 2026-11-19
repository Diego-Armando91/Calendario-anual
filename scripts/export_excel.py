"""Paso 5. Excel EDITABLE de la propuesta: python export_excel.py propuesta_v9.csv v9
Las hojas de resumen son fórmulas que leen la hoja 'Propuesta anual': si cambias un código allí,
se recalculan solas (horas, JA, viernes/domingo noche, cobertura, descansos de 1 día, rachas, semanas).
Códigos: M MS MD C T TS TD N NV NS, y V = vacaciones / día libre fijo. Celda vacía = descanso.
Hojas: Propuesta anual, Resumen personas, Semanas, Cobertura, Viernes noche, Domingos noche, Rotaciones, Criterios.
Para leer de vuelta un Excel editado: python leer_excel.py <excel> propuesta_vX.csv salida.csv"""
import sys, json, numpy as np, datetime as dt
from evaluate import *
import model as M
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
TAG=sys.argv[2] if len(sys.argv)>2 else 'nueva'
roles_w={int(k):v for k,v in json.load(open(os.path.join(WORK,'roles.json')))['roles'].items()}
x=load_solution(sys.argv[1]); X,C_=sol_matrix(x); Xb,Cb=base_matrix()
DIAS=['L','M','X','J','V','S','D']; DIASF=['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo']
codes_lbl=[k.split('(')[1].strip(')') for k in summ]
def code_of(i,d):
    r=role[i,d]; w=wd[d]
    if r=='M': return 'C' if (people[i]=='JUANMA' and w<5) else ('MS' if w==5 else 'MD' if w==6 else 'M')
    if r=='T': return 'TS' if w==5 else 'TD' if w==6 else 'T'
    return 'NV' if w==4 else 'NS' if w==5 else 'N'
FILL={'M':'A9D18E','MS':'A9D18E','MD':'A9D18E','C':'FFE699','T':'C9A7E0','TS':'C9A7E0','TD':'C9A7E0','N':'E97132','NV':'E97132','NS':'F4B183','V':'404040'}
YEL=Side(style='medium',color='FFFF00'); thin=Side(style='thin',color='DDDDDD'); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
RED=PatternFill('solid',fgColor='FFC7CE'); GREEN=PatternFill('solid',fgColor='C6EFCE'); GREY=PatternFill('solid',fgColor='F2F2F2')
WORK_CODES='{"M","MS","MD","C","T","TS","TD","N","NV","NS"}'
PA="'Propuesta anual'"
R0=5; R1=R0+P-1; C0=4; C1=C0+D-1            # filas de personas y columnas de días
HROW=R1+2; WROW=R1+3                          # filas auxiliares: horas de la noche y semana
col=lambda d: L(C0+d)
def rng(r,d0=0,d1=D-1,abs_=True):
    s='$' if abs_ else ''
    return f"{PA}!${col(d0)}{s}{r}:${col(d1)}{s}{r}"
def W(r,d0,d1): return f"ISNUMBER(MATCH({PA}!{col(d0)}{r}:{col(d1)}{r},{WORK_CODES},0))"
BLOCK=f"{PA}!$D${R0}:${col(D-1)}${R1}"
NAMES=f"{PA}!$A${R0}:$A${R1}"
wb=openpyxl.Workbook()
# ======================= Propuesta anual
ws=wb.active; ws.title='Propuesta anual'
ws['A1']=f'PROPUESTA DE CUADRANTE 2027 (versión {TAG}) — EDITABLE'; ws['A1'].font=Font(bold=True,size=14)
ws['A2']=('Escribe el código en la celda y todo se recalcula: M/MS/MD mañana · C central · T/TS/TD tarde · N noche · NV viernes noche · NS sábado noche · '
          'V vacaciones o día libre fijo · vacío = descanso. Rojo = cierre/festivo no laborable. Los colores se ponen solos.')
ws['A2'].font=Font(italic=True,size=8)
ws.cell(3,3,'JA').font=Font(bold=True,size=8)
for d in range(D):
    c=ws.cell(3,C0+d,date[d]); c.number_format='d-mmm'; c.alignment=Alignment(text_rotation=90,horizontal='center'); c.font=Font(size=7,bold=True)
    c2=ws.cell(4,C0+d,DIAS[wd[d]]); c2.alignment=Alignment(horizontal='center'); c2.font=Font(size=8,bold=True,color='C00000' if wd[d]>=5 else '000000')
    ws.cell(HROW,C0+d,7 if (wd[d]==6 or date[d] in SPECIAL) else 8).font=Font(size=6,color='808080')
    ws.cell(WROW,C0+d,wk(d)).font=Font(size=6,color='808080')
ws.cell(HROW,1,'h noche (aux.)').font=Font(size=7,color='808080'); ws.cell(WROW,1,'semana (aux.)').font=Font(size=7,color='808080')
ws.row_dimensions[3].height=38
for i,n in enumerate(people):
    r=R0+i; ws.cell(r,1,n).font=Font(bold=True,size=9); ws.cell(r,2,codes_lbl[i]).font=Font(size=9); ws.cell(r,3,int(JA[i])).font=Font(size=7,color='808080')
    for d in range(D):
        c=ws.cell(r,C0+d); c.font=Font(size=7); c.alignment=Alignment(horizontal='center'); c.border=BOX
        if X[i,d]: c.value=code_of(i,d)
        elif vac[i,d]: c.value='V'
        elif closed[d] or fl[n][d]=='FFFF0000': c.fill=PatternFill('solid',fgColor='FF0000')
ws.column_dimensions['A'].width=16; ws.column_dimensions['B'].width=5; ws.column_dimensions['C'].width=5
for d in range(D): ws.column_dimensions[col(d)].width=3.4
ws.freeze_panes=ws.cell(R0,C0)
area=f'D{R0}:{col(D-1)}{R1}'
for k,f in FILL.items():
    kw=dict(fill=PatternFill('solid',fgColor=f,bgColor=f))
    if k in ('MS','MD','TS','TD'): kw['border']=Border(left=YEL,right=YEL,top=YEL,bottom=YEL)
    if k=='V': kw['font']=Font(color='808080')
    ws.conditional_formatting.add(area,CellIsRule(operator='equal',formula=[f'"{k}"'],**kw))
dv=DataValidation(type='list',formula1='"M,MS,MD,C,T,TS,TD,N,NV,NS,V"',allow_blank=True,showErrorMessage=True,errorStyle='warning',
                  errorTitle='Código no reconocido',error='Usa M, MS, MD, C, T, TS, TD, N, NV, NS o V (vacío = descanso).')
ws.add_data_validation(dv); dv.add(area)
# ======================= Resumen personas
wp=wb.create_sheet('Resumen personas')
hdr=['Persona','Código','Días trabajados','Horas','JA','Diferencia (h)','Estado JA','Vacaciones / libres (V)','Sábados trabajados',
     'Viernes noche (NV)','Sábado noche (NS)','Domingo noche','Descansos de 1 día','Rachas de 6 días','7 días seguidos (no permitido)',
     'Semanas bajo el mínimo']
wp.append(hdr)
for c in wp[1]: c.font=Font(bold=True); c.alignment=Alignment(wrap_text=True,horizontal='center',vertical='center')
NW_=NW
for i in range(P):
    r=R0+i; rr=2+i; Rr=rng(r)
    wp.cell(rr,1,f'={PA}!A{r}'); wp.cell(rr,2,f'={PA}!B{r}')
    wp.cell(rr,3,f'=SUMPRODUCT(--ISNUMBER(MATCH({Rr},{WORK_CODES},0)))')
    wp.cell(rr,4,f'=8*C{rr}-SUMPRODUCT(ISNUMBER(MATCH({Rr},{{"N","NV","NS"}},0))*(8-{rng(HROW)}))')
    wp.cell(rr,5,f'={PA}!C{r}')
    wp.cell(rr,6,f'=D{rr}-E{rr}')
    wp.cell(rr,7,f'=IF(F{rr}<0,"FALTAN "&-F{rr}&" h",IF(F{rr}>{C.HOURS_TOL},"SOBRAN "&(F{rr}-{C.HOURS_TOL})&" h","OK"))')
    wp.cell(rr,8,f'=COUNTIF({Rr},"V")')
    wp.cell(rr,9,f'=SUMPRODUCT(({rng(4)}="S")*ISNUMBER(MATCH({Rr},{WORK_CODES},0)))')
    wp.cell(rr,10,f'=COUNTIF({Rr},"NV")'); wp.cell(rr,11,f'=COUNTIF({Rr},"NS")')
    wp.cell(rr,12,f'=SUMPRODUCT(({rng(4)}="D")*({Rr}="N"))')
    wp.cell(rr,13,f'=SUMPRODUCT({W(r,0,D-3)}*({PA}!{col(1)}{r}:{col(D-2)}{r}="")*{W(r,2,D-1)})')
    # rachas de 6: 6 días trabajados con el día anterior libre (la columna C tiene la JA, que no es un código)
    six='*'.join(W(r,k,D-6+k) for k in range(6)); prev=f'(1-ISNUMBER(MATCH({PA}!C{r}:{col(D-7)}{r},{WORK_CODES},0)))'
    wp.cell(rr,14,f'=SUMPRODUCT({prev}*{six})')
    seven='*'.join(W(r,k,D-7+k) for k in range(7))
    wp.cell(rr,15,f'=SUMPRODUCT({seven})')
    wp.cell(rr,16,f"=SUMPRODUCT(--('Semanas'!C{rr}:{L(2+NW_)}{rr}<'Mínimos'!C{rr}:{L(2+NW_)}{rr}))")
last=1+P; tr=last+1
wp.cell(tr,1,'TOTAL').font=Font(bold=True)
for c in (3,4,5,8,9,10,11,12,13,14,15,16): wp.cell(tr,c,f'=SUM({L(c)}2:{L(c)}{last})').font=Font(bold=True)
wp.conditional_formatting.add(f'G2:G{last}',FormulaRule(formula=['G2<>"OK"'],fill=RED)); wp.conditional_formatting.add(f'G2:G{last}',FormulaRule(formula=['G2="OK"'],fill=GREEN))
wp.conditional_formatting.add(f'M2:M{last}',CellIsRule(operator='greaterThan',formula=[str(C.MAX_DESCANSOS_1)],fill=RED))
wp.conditional_formatting.add(f'N2:N{last}',CellIsRule(operator='greaterThan',formula=[str(C.MAX_RACHAS6)],fill=RED))
wp.conditional_formatting.add(f'O2:P{last}',CellIsRule(operator='greaterThan',formula=['0'],fill=RED))
wp['A'+str(tr+2)]=(f'Límites: horas entre JA y JA+{C.HOURS_TOL} · máx. {C.MAX_DESCANSOS_1} descansos de 1 día · máx. {C.MAX_RACHAS6} rachas de 6 días · '
                  'nunca 7 días seguidos · ninguna semana bajo el mínimo (hoja Semanas). En rojo lo que no cumple.')
wp['A'+str(tr+2)].font=Font(italic=True,size=9)
wp.column_dimensions['A'].width=16
for c in range(2,17): wp.column_dimensions[L(c)].width=12
wp.row_dimensions[1].height=48; wp.freeze_panes='C2'
# ======================= Semanas (días trabajados por semana) + Mínimos
wsm=wb.create_sheet('Semanas'); wmin=wb.create_sheet('Mínimos')
for sh in (wsm,wmin):
    sh.cell(1,1,'Persona').font=Font(bold=True); sh.cell(1,2,'Código').font=Font(bold=True)
    for w in range(NW):
        dd=[d for d in range(D) if wk(d)==w]; c=sh.cell(1,3+w,date[dd[0]]); c.number_format='d-mmm'
        c.font=Font(bold=True,size=8); c.alignment=Alignment(text_rotation=90,horizontal='center'); sh.column_dimensions[L(3+w)].width=4
    sh.row_dimensions[1].height=40; sh.column_dimensions['A'].width=16; sh.freeze_panes='C2'
for i in range(P):
    r=R0+i; rr=2+i
    for sh in (wsm,wmin): sh.cell(rr,1,f'={PA}!A{r}'); sh.cell(rr,2,f'={PA}!B{r}')
    for w in range(NW):
        wsm.cell(rr,3+w,f'=SUMPRODUCT(({rng(WROW)}={w})*ISNUMBER(MATCH({rng(r)},{WORK_CODES},0)))').alignment=Alignment(horizontal='center')
        wmin.cell(rr,3+w,int(M.WEEKMIN[i,w])).alignment=Alignment(horizontal='center')
wsm.conditional_formatting.add(f'C2:{L(2+NW)}{1+P}',FormulaRule(formula=["C2<'Mínimos'!C2"],fill=RED))
wsm.cell(P+3,1,'Días trabajados por semana (lunes a domingo; la primera columna es vie 1 - dom 3 ene). En rojo: por debajo del mínimo de la hoja Mínimos.').font=Font(italic=True,size=9)
wmin.cell(P+3,1,'Mínimo de días por semana (4 normal, 3 con festivo, proporcional con vacaciones). Si pintas vacaciones nuevas, ajusta aquí el mínimo a mano.').font=Font(italic=True,size=9)
# ======================= Cobertura
wc=wb.create_sheet('Cobertura')
wc.append(['Fecha','Día','Mañana','Tarde','Noche','Objetivo mañana','Objetivo tarde','Objetivo noche','Mañana (actual)','Tarde (actual)','Noche (actual)'])
for c in wc[1]: c.font=Font(bold=True); c.alignment=Alignment(wrap_text=True,horizontal='center')
GRP={'M':'{"M","MS","MD","C"}','T':'{"T","TS","TD"}','N':'{"N","NV","NS"}'}
for d in range(D):
    rr=2+d; colref=f'INDEX({BLOCK},0,{d+1})'
    wc.cell(rr,1,date[d]).number_format='dd/mm/yyyy'; wc.cell(rr,2,DIASF[wd[d]])
    for k,c in enumerate('MTN'):
        wc.cell(rr,3+k,f'=SUMPRODUCT(--ISNUMBER(MATCH({colref},{GRP[c]},0)))')
        wc.cell(rr,6+k,int(M.TGT[k,d]))
        wc.cell(rr,9+k,int(sum(1 for i in range(P) if Xb[i,d] and Cb[i,d]==c)))
for k in range(3):
    a=L(3+k); b=L(6+k); wc.conditional_formatting.add(f'{a}2:{a}{D+1}',FormulaRule(formula=[f'{a}2<{b}2'],fill=RED))
for c_,w_ in zip('ABCDEFGHIJK',(12,11,9,9,9,10,10,10,10,10,10)): wc.column_dimensions[c_].width=w_
wc.freeze_panes='C2'
# ======================= Viernes noche / Domingos noche
def night_sheet(title,wday,code):
    sh=wb.create_sheet(title)
    sh.append(['Fecha','Personas de noche','Objetivo','Quién','Personas (actual)'])
    for c in sh[1]: c.font=Font(bold=True)
    rr=1
    for d in range(D):
        if wd[d]!=wday: continue
        rr+=1; colref=f'INDEX({BLOCK},0,{d+1})'
        sh.cell(rr,1,date[d]).number_format='dd/mm/yyyy'
        sh.cell(rr,2,f'=COUNTIF({colref},"{code}")'); sh.cell(rr,3,int(M.TGT[2,d]))
        ref=f'D{rr}'; sh[ref]=ArrayFormula(ref,f'=_xlfn.TEXTJOIN(", ",TRUE,IF({colref}="{code}",{NAMES},""))')
        sh.cell(rr,5,int(sum(1 for i in range(P) if Xb[i,d] and Cb[i,d]=='N')))
    sh.conditional_formatting.add(f'B2:B{rr}',FormulaRule(formula=['B2<C2'],fill=RED))
    sh.column_dimensions['A'].width=12; sh.column_dimensions['B'].width=16; sh.column_dimensions['C'].width=10; sh.column_dimensions['D'].width=55; sh.column_dimensions['E'].width=16
    # por persona
    sh['G1']=f'{title} por persona'; sh['G1'].font=Font(bold=True)
    for c_,t in zip('GHI',('Persona','Total','Actual')): sh[f'{c_}2']=t; sh[f'{c_}2'].font=Font(bold=True)
    k=0
    for i in range(P):
        if not any(role[i,d]=='N' for d in range(D)): continue
        r=R0+i; row=3+k; k+=1
        sh.cell(row,7,f'={PA}!A{r}')
        sh.cell(row,8,f'=SUMPRODUCT(({rng(4)}="{DIAS[wday]}")*({rng(r)}="{code}"))')
        sh.cell(row,9,int(sum(1 for d in range(D) if wd[d]==wday and Xb[i,d] and Cb[i,d]=='N')))
    sh.column_dimensions['G'].width=16
    sh.cell(rr+2,1,'Objetivo: viernes 3 (2 si no hay sábado noche o si alguien de noche tiene vacaciones esa semana); domingo 2 (1 con vacaciones en noche). En rojo, por debajo.').font=Font(italic=True,size=9)
night_sheet('Viernes noche',4,'NV'); night_sheet('Domingos noche',6,'N')
# ======================= Rotaciones (estático)
wr=wb.create_sheet('Rotaciones')
rot=['ESTHER','DAVID P','DIEGO ARMANDO','LORENA','BEA','DENIA']
wr.append(['Semana','Desde (lun)','Hasta (dom)']+[r_.title() for r_ in rot])
for c in wr[1]: c.font=Font(bold=True)
for w in range(NW):
    dd=[d for d in range(D) if wk(d)==w]
    row=[w,date[dd[0]],date[dd[-1]]]+[str(roles_w[IDX[r_]][w]) for r_ in rot]
    wr.append(row); wr.cell(wr.max_row,2).number_format='dd/mm/yyyy'; wr.cell(wr.max_row,3).number_format='dd/mm/yyyy'
    for k in range(6):
        v=row[3+k]; c=wr.cell(wr.max_row,4+k); c.alignment=Alignment(horizontal='center'); c.fill=PatternFill('solid',fgColor={'M':'A9D18E','T':'C9A7E0','N':'E97132'}.get(v,'FFFFFF'))
for c_,w_ in zip('ABCDEFGHI',(8,12,12,10,10,14,10,8,8)): wr.column_dimensions[c_].width=w_
wr.freeze_panes='D2'
# ======================= Criterios
wk_=wb.create_sheet('Criterios')
for l in ['CRITERIOS DE LA PROPUESTA','',
 'Rotación: Esther/David y Diego/Lorena alternan mañana-noche; Bea/Denia alternan mañana-tarde, en bloques de unas 5 semanas. Cada semana de lunes a domingo es entera del mismo turno.',
 'Fijos: Chema y Juanma (C) mañana; Michel, Alfredo, Irene y Sofía tarde; Armando, Juanma V y Rubén noche.',
 'Vacaciones: las del cuadrante original. El fin de semana anterior a unas vacaciones es libre (a Diego solo los que marcó él). Juanma no trabaja los fines de semana con su hijo.',
 'Cierres (25-26 mar, 24-25 y 31 dic): nadie trabaja. Noches especiales 00:00-07:00 (1 ene, 6 ene, 27 mar, 12 oct): solo noche, 7 h. La noche anterior a un festivo o cierre no se trabaja (5 ene, 24 mar, 11 oct, 23 dic); el 6 ene y el 12 oct no hay mañana ni tarde. 2-3 ene: solo Chema (mañana) y Michel (tarde).',
 'Cobertura objetivo: mañana 3, tarde 3, noche 3 de lunes a jueves. Viernes noche 3 (2 si no hay sábado noche o si alguien de noche tiene vacaciones esa semana). Domingo noche 2 (1 con vacaciones en noche). Sábado noche solo del 27 feb al 29 may y del 17 jul al 4 dic.',
 'Fines de semana de mañana y tarde: 1 persona hasta el 3 de septiembre y 2 desde el 4; quien trabaja el sábado trabaja el domingo.',
 f'Por persona y año: máximo {C.MAX_DESCANSOS_1} descansos de 1 solo día y {C.MAX_RACHAS6} rachas de 6 días seguidos (crucen o no de semana); nunca 7 días seguidos.',
 'Mínimo semanal: 4 días (3 con festivo, proporcional con vacaciones). Lo normal es 5, pero la jornada anual no da para 5 todas las semanas.',
 'Noches: tandas de al menos 2 noches seguidas (mejor 3-5); una noche suelta solo junto a un día rojo. Se evita viernes noche + sábado libre + domingo noche.',
 'Fines de semana: alternos, como máximo 2 al mes; dos seguidos de vez en cuando. Horas: cada persona entre su JA y JA + 8 h (domingo noche y noches especiales = 7 h).',
 '','CÓMO EDITAR: cambia los códigos en la hoja Propuesta anual (o borra la celda para dejar descanso). El resto de hojas se recalcula solo; lo que no cumple sale en rojo.']:
    wk_.append([l])
wk_.column_dimensions['A'].width=170; wk_['A1'].font=Font(bold=True,size=13)
# ======================= Cambios respecto a otra versión (opcional: 3er argumento = CSV anterior)
if len(sys.argv)>3:
    xa=load_solution(sys.argv[3]); wch=wb.create_sheet('Cambios',1)
    wch.append([f'Cambios respecto a la versión anterior ({sys.argv[3]}). Celdas de la hoja Propuesta anual que han cambiado.']); wch['A1'].font=Font(bold=True,size=12)
    wch.append(['Persona','Fecha','Día','Antes','Ahora']);
    for c in wch[2]: c.font=Font(bold=True)
    for i in range(P):
        for d in range(D):
            if xa[i,d]!=x[i,d]:
                wch.append([people[i].title(),date[d],DIASF[wd[d]],code_of(i,d) if xa[i,d] else 'descanso',code_of(i,d) if x[i,d] else 'descanso'])
                wch.cell(wch.max_row,2).number_format='dd/mm/yyyy'
    for c_,w_ in zip('ABCDE',(16,12,11,10,10)): wch.column_dimensions[c_].width=w_
out=os.path.join(OUT,f'Propuesta_Cuadrante_2027_{TAG}.xlsx'); wb.save(out); print('ok',out)
