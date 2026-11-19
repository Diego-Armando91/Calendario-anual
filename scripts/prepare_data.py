"""Paso 1. Lee la hoja 'Rotación anual' del Excel original y guarda work/base.json.
Filas 4-18 = 15 personas; columnas D..ND = 365 días de 2027 (la cabecera guarda 2026 con formato d-mmm,
pero las letras de la fila 2 confirman que es 2027: el 1 de enero es viernes).
Relleno negro (tema 1) = vacaciones; FF00B0F0 = azul (no obligatorio); FFFF0000 = rojo (festivo)."""
import json, datetime as dt, openpyxl
from paths import *
wv=openpyxl.load_workbook(SRC,data_only=True); wf=openpyxl.load_workbook(SRC)
ws=wv['Rotación anual']; wsf=wf['Rotación anual']
def fill(c):
    f=c.fill
    if f is None or f.fill_type!='solid': return None
    col=f.fgColor
    if col.type=='rgb': return col.rgb
    if col.type=='theme': return f'theme{col.theme}'
people=[];seq={};fl={}
start=dt.date(2027,1,1)
for i in range(365):  # comprobación de días de la semana
    assert ws.cell(2,4+i).value=='LMXJVSD'[(start+dt.timedelta(i)).weekday()]
for r in range(4,19):
    nm=ws.cell(r,1).value; people.append(nm)
    seq[nm]=[(ws.cell(r,4+i).value.strip() if isinstance(ws.cell(r,4+i).value,str) and ws.cell(r,4+i).value.strip() else None) for i in range(365)]
    fl[nm]=[fill(wsf.cell(r,4+i)) for i in range(365)]
summ={}
for r in range(26,41):
    v=ws.cell(r,1).value
    if v: summ[v]=dict(fin=ws.cell(r,2).value,sab=ws.cell(r,3).value,horas=ws.cell(r,4).value,ja=ws.cell(r,7).value)
json.dump(dict(people=people,seq=seq,fl=fl,summ=summ),open(os.path.join(WORK,'base.json'),'w'),ensure_ascii=False)
print('OK', len(people),'personas')
