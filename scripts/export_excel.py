"""Paso 5. Excel de la propuesta: python export_excel.py propuesta_v5.csv v5
Hojas: Propuesta anual (formato Rotación anual), Rotaciones, Resumen personas, Cobertura, Viernes noche, Criterios."""
import sys, json, pickle, numpy as np, collections, datetime as dt
from evaluate import *
TAG=sys.argv[2] if len(sys.argv)>2 else 'nueva'
roles_w={int(k):v for k,v in json.load(open(os.path.join(WORK,'roles.json')))['roles'].items()}
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
x=load_solution(sys.argv[1]); X,C=sol_matrix(x); per,cov=metrics(X,C); WS=weekend_stats(X)
Xb,Cb=base_matrix(); perb,covb=metrics(Xb,Cb)
DIAS=['L','M','X','J','V','S','D']; DIASF=['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo']
CODE={}
for i in range(P):
    for d in range(D):
        if X[i,d]:
            r=role[i,d]; w=wd[d]
            if r=='M': CODE[i,d]=('C' if (people[i]=='JUANMA' and w<5) else ('MS' if w==5 else 'MD' if w==6 else 'M'))
            elif r=='T': CODE[i,d]='TS' if w==5 else 'TD' if w==6 else 'T'
            else: CODE[i,d]='NV' if w==4 else 'NS' if w==5 else 'N'
COL={'M':'A9D18E','MS':'A9D18E','MD':'A9D18E','C':'FFE699','T':'C9A7E0','TS':'C9A7E0','TD':'C9A7E0','N':'E97132','NV':'E97132','NS':'F4B183'}
YEL=Side(style='medium',color='FFFF00'); thin=Side(style='thin',color='DDDDDD')
codes_names={n:ws_ for n,ws_ in zip(people,[k.split('(')[1].strip(')') for k in summ])}
wb=openpyxl.Workbook(); ws=wb.active; ws.title='Propuesta anual'
ws['A1']='PROPUESTA DE CUADRANTE 2027 (versión '+TAG+')'; ws['A1'].font=Font(bold=True,size=14)
HR=3
for j,d in enumerate(date):
    c=ws.cell(HR,4+j,d); c.number_format='d-mmm'; c.alignment=Alignment(text_rotation=90,horizontal='center'); c.font=Font(size=7,bold=True)
    c2=ws.cell(HR+1,4+j,DIAS[wd[j]]); c2.alignment=Alignment(horizontal='center'); c2.font=Font(size=8,bold=True,color='C00000' if wd[j]>=5 else '000000')
ws.row_dimensions[HR].height=38
for i,n in enumerate(people):
    r=HR+2+i; ws.cell(r,1,n).font=Font(bold=True,size=9); ws.cell(r,2,summ_code:=[k for k in summ][i].split('(')[1].strip(')')).font=Font(size=9)
    for d in range(D):
        c=ws.cell(r,4+d); c.font=Font(size=7); c.alignment=Alignment(horizontal='center'); c.border=Border(left=thin,right=thin,top=thin,bottom=thin)
        if (i,d) in CODE:
            k=CODE[i,d]; c.value=k; c.fill=PatternFill('solid',fgColor=COL[k])
            if k in('MS','MD','TS','TD'): c.border=Border(left=YEL,right=YEL,top=YEL,bottom=YEL)
        elif vac[i,d]: c.fill=PatternFill('solid',fgColor='404040')
        elif closed[d]: c.fill=PatternFill('solid',fgColor='FF0000')
ws.column_dimensions['A'].width=16; ws.column_dimensions['B'].width=5; ws.column_dimensions['C'].width=2
for d in range(D): ws.column_dimensions[get_column_letter(4+d)].width=3.4
ws.freeze_panes=ws.cell(HR+2,4)
# leyenda
ws['A2']='Verde M · amarillo C (central) · lila T · naranja N/NV · naranja claro NS · bordes amarillos = fin de semana · gris oscuro = vacaciones (se respetan las del cuadrante original) · rojo = cierre'; ws['A2'].font=Font(italic=True,size=8)
# --- Rotaciones
wr=wb.create_sheet('Rotaciones')
rot=['ESTHER','DAVID P','DIEGO ARMANDO','LORENA','BEA','DENIA']
wr.append(['Semana','Desde (lun)','Hasta (dom)']+[r.title() for r in rot]); 
for c in wr[1]: c.font=Font(bold=True)
for w in range(NW):
    dd=[d for d in range(D) if wk(d)==w]; d0=dd[0]; d1=dd[-1]
    row=[w,date[d0],date[d1]]+[str(roles_w[IDX[r]][w]) or '-' for r in rot]
    wr.append(row); wr.cell(wr.max_row,2).number_format='dd/mm/yyyy'; wr.cell(wr.max_row,3).number_format='dd/mm/yyyy'
    for k in range(6):
        v=row[3+k]; c=wr.cell(wr.max_row,4+k); c.alignment=Alignment(horizontal='center'); c.fill=PatternFill('solid',fgColor={'M':'A9D18E','T':'C9A7E0','N':'E97132'}.get(v,'FFFFFF'))
for col,w in zip('ABCDEFGHI',(8,12,12,10,10,14,10,8,8)): wr.column_dimensions[col].width=w
wr.freeze_panes='D2'
# --- Resumen personas
JAv={k.split(' (')[0]:v['ja'] for k,v in summ.items()}
wp=wb.create_sheet('Resumen personas')
hdr=['Persona','Días','Horas','JA','Ajuste (días)','Sábados trabajados','Fines de semana seguidos','Meses con 3+ fines de semana','Viernes noche','Domingo noche','Descansos de 1 día','Rachas ≥6 días','Semanas L–D de 6 días','Descansos de 1 día (actual)','Rachas ≥6 días (actual)','Semanas de 6 días (actual)']
wp.append(hdr)
for c in wp[1]: c.font=Font(bold=True); c.alignment=Alignment(wrap_text=True,horizontal='center')
for i,n in enumerate(people):
    h=sum((7 if (role[i,d]=='N' and (wd[d]==6 or date[d] in SPECIAL)) else 8) for d in range(D) if X[i,d])
    fv=sum(1 for d in range(D) if X[i,d] and role[i,d]=='N' and wd[d]==4); dn=sum(1 for d in range(D) if X[i,d] and role[i,d]=='N' and wd[d]==6)
    wp.append([n.title(),int(X[i].sum()),h,JAv[n],round((h-JAv[n])/8,2),per[n]['sab'],WS[n]['seguidos'],WS[n]['meses3'],fv,dn,per[n]['simples'],per[n]['rachas6'],per[n]['sem6'],perb[n]['simples'],perb[n]['rachas6'],perb[n]['sem6']])
r=wp.max_row+1
wp.cell(r,1,'TOTAL').font=Font(bold=True)
for c in (2,3,4,6,7,8,9,10,11,12,13,14,15,16): wp.cell(r,c,f'=SUM({get_column_letter(c)}2:{get_column_letter(c)}{r-1})').font=Font(bold=True)
wp.column_dimensions['A'].width=16
for c in range(2,17): wp.column_dimensions[get_column_letter(c)].width=13
wp.row_dimensions[1].height=48
# --- Cobertura
wc=wb.create_sheet('Cobertura')
wc.append(['Fecha','Día','Mañana','Tarde','Noche','Mañana (actual)','Tarde (actual)','Noche (actual)'])
for c in wc[1]: c.font=Font(bold=True)
for d in range(D):
    wc.append([date[d],DIASF[wd[d]]]+[sum(1 for i in range(P) if X[i,d] and C[i,d]==c) for c in 'MTN']+[sum(1 for i in range(P) if Xb[i,d] and Cb[i,d]==c) for c in 'MTN'])
    wc.cell(wc.max_row,1).number_format='dd/mm/yyyy'
for col,w in zip('ABCDEFGH',(12,11,9,9,9,15,14,14)): wc.column_dimensions[col].width=w
wc.freeze_panes='A2'
# --- Viernes noche
wv=wb.create_sheet('Viernes noche')
wv.append(['Fecha','Personas de noche','Quién (propuesta)','Personas (actual)'])
for c in wv[1]: c.font=Font(bold=True)
for d in range(D):
    if wd[d]==4:
        names=[people[i].title() for i in range(P) if X[i,d] and C[i,d]=='N']
        wv.append([date[d],len(names),', '.join(names),sum(1 for i in range(P) if Xb[i,d] and Cb[i,d]=='N')]); wv.cell(wv.max_row,1).number_format='dd/mm/yyyy'
wv.column_dimensions['A'].width=12; wv.column_dimensions['B'].width=16; wv.column_dimensions['C'].width=60; wv.column_dimensions['D'].width=16
# --- Domingos noche
wdn=wb.create_sheet('Domingos noche')
wdn.append(['Fecha','Personas de noche','Quién (propuesta)','Personas (actual)'])
for c in wdn[1]: c.font=Font(bold=True)
for d in range(D):
    if wd[d]==6:
        names=[people[i].title() for i in range(P) if X[i,d] and C[i,d]=='N']
        wdn.append([date[d],len(names),', '.join(names),sum(1 for i in range(P) if Xb[i,d] and Cb[i,d]=='N')]); wdn.cell(wdn.max_row,1).number_format='dd/mm/yyyy'
        if len(names)!=2: wdn.cell(wdn.max_row,2).fill=PatternFill('solid',fgColor='FFC7CE')
wdn.column_dimensions['A'].width=12; wdn.column_dimensions['B'].width=16; wdn.column_dimensions['C'].width=60; wdn.column_dimensions['D'].width=16
wdn['F1']='Domingos noche por persona'; wdn['F1'].font=Font(bold=True)
wdn['F2']='Persona'; wdn['G2']='Propuesta'; wdn['H2']='Actual'
for c in ('F2','G2','H2'): wdn[c].font=Font(bold=True)
for k,i in enumerate([i for i in range(P) if any(role[i,d]=='N' for d in range(D))]):
    wdn.cell(3+k,6,people[i].title()); wdn.cell(3+k,7,sum(1 for d in range(D) if wd[d]==6 and X[i,d] and C[i,d]=='N')); wdn.cell(3+k,8,sum(1 for d in range(D) if wd[d]==6 and Xb[i,d] and Cb[i,d]=='N'))
wdn.column_dimensions['F'].width=16
# --- Criterios
wk=wb.create_sheet('Criterios')
for l in ['PARÁMETROS USADOS EN LA PROPUESTA','',
 'Rotación: Esther/David y Diego/Lorena alternan mañana-noche; Bea/Denia alternan mañana-tarde, en bloques de 5 semanas (algún bloque de 6 para equilibrar vacaciones). Cada semana de lunes a domingo es entera del mismo turno.',
 'Fijos: Chema y Juanma (C) mañana; Michel, Alfredo, Irene y Sofía tarde; Armando, Juanma V y Rubén noche.',
 'Vacaciones: se mantienen las del cuadrante original (celdas negras) y el fin de semana anterior a unas vacaciones es libre.',
 'Días azules: no obligatorios, se pueden trabajar. Cierres (25-26 mar, 24-25 y 31 dic): nadie trabaja. Noches especiales 00:00-07:00 (1 ene, 6 ene, 27 mar, 12 oct): solo personal de noche, 7 h. La noche anterior a un festivo o cierre no se trabaja (5 ene, 24 mar, 11 oct, 23 dic), y el 6 ene y el 12 oct no hay mañana ni tarde. Primer fin de semana (2-3 ene): solo Chema de mañana y Michel de tarde.',
 'Cobertura objetivo: mañana 3, tarde 3, noche 3 de lunes a jueves (hasta 4 si sobra gente). Domingo noche 2.',
 'Mínimo semanal: 4 días por semana completa (5 lo habitual), 3 en semanas con festivo no laborable o cierre; proporcional si hay vacaciones parte de la semana.',
 'Fines de semana de mañana y tarde: mismo formato que el original (1 persona hasta el 3 de septiembre, 2 desde el 4 de septiembre); quien trabaja el sábado trabaja también el domingo.',
 'Descanso: máximo 6 días seguidos (se penaliza el sexto), se penalizan los descansos de 1 solo día y los días de trabajo aislados. Tras una noche no se entra en mañana o tarde al día siguiente.',
 'Noches: se favorecen tandas continuas de 4-5 noches aunque crucen de una semana a otra (p. ej. de jueves a martes); se penaliza mucho trabajar viernes noche y domingo noche con el sábado libre.',
 'Sábado noche: solo se trabaja del 27 de febrero al 29 de mayo y del 17 de julio al 4 de diciembre (ambos incluidos). Viernes noche: objetivo 3 personas cuando hay sábado noche y 2 cuando no lo hay.',
 'Fines de semana y sábados noche: se intenta alternarlos (se permiten algunos seguidos) y como máximo 2 al mes; 3 solo si las vacaciones lo obligan. Se permite alguna semana de 6 días, penalizada para que sea poco frecuente.',
 'Horas: cada persona cumple su JA (entre JA y JA + 8 h). Domingo noche y noches especiales = 7 h.',
 'Equidad: fines de semana equilibrados dentro de cada grupo (solo día / rotan mañana-noche / solo noche) y noches equilibradas entre los 4 que rotan mañana-noche.',
 'La propuesta sale de un modelo de optimización; es una primera versión para revisar, no un cuadrante cerrado.']:
    wk.append([l])
wk.column_dimensions['A'].width=160; wk['A1'].font=Font(bold=True,size=13)
wb.save(os.path.join(OUT,f'Propuesta_Cuadrante_2027_{TAG}.xlsx'))
pickle.dump(dict(per=per,perb=perb,cov=dict(cov),covb=dict(covb),ws=WS,wsb=weekend_stats(Xb),nr=night_runs(X,C),nrb=night_runs(Xb,Cb),mix=mixed_weeks(X,C),mixb=mixed_weeks(Xb,Cb)),open(os.path.join(WORK,f'metricas_{TAG}.pkl'),'wb'))
print('ok',len(CODE))
