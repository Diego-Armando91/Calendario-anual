"""Lee una propuesta editada a mano (hoja 'Propuesta anual') y la compara con la propuesta generada.
  python leer_excel.py ../data/Propuesta_Cuadrante_2027_v8_Diego.xlsx propuesta_v8.csv [salida.csv]
Calcula horas por persona frente a su JA a partir de los códigos de cada celda."""
import sys, openpyxl, numpy as np
from common import *
import model as M
CLm={'M':'M','MS':'M','MD':'M','C':'M','T':'T','TS':'T','TD':'T','N':'N','NV':'N','NS':'N'}
wb=openpyxl.load_workbook(sys.argv[1],data_only=True); ws=wb['Propuesta anual']
rows={}
for r in range(1,ws.max_row+1):
    v=ws.cell(r,1).value
    if isinstance(v,str) and v.strip().upper() in people: rows[v.strip().upper()]=r
col0=None
for c in range(1,30):
    v=ws.cell(3,c).value
    if hasattr(v,'month') and v.month==1 and v.day==1: col0=c; break
assert col0 and len(rows)==P, (col0,rows)
code=np.full((P,D),'',dtype=object); x=np.zeros((P,D),int); raros=[]
for i,p in enumerate(people):
    for d in range(D):
        v=ws.cell(rows[p],col0+d).value
        v=v.strip().upper() if isinstance(v,str) else ''
        if v:
            if v not in CLm: raros.append((p,str(date[d]),v)); continue
            code[i,d]=CLm[v]; x[i,d]=1
base=load_solution(sys.argv[2]) if len(sys.argv)>2 else None
print('Celdas con códigos desconocidos:',raros)
cambios=[]
for i,p in enumerate(people):
    for d in range(D):
        gen=M.role[i,d] if (base is not None and base[i,d]) else ''
        if code[i,d]!=gen: cambios.append((p,str(date[d]),date[d].strftime('%a'),gen or '-',code[i,d] or '-'))
print('Cambios a mano frente a la propuesta generada:',len(cambios))
for c in cambios: print('  ',c)
print('\nHORAS frente a JA (domingo noche y noches especiales = 7 h):')
for i,p in enumerate(people):
    h=sum(7 if (code[i,d]=='N' and (wd[d]==6 or date[d] in SPECIAL)) else 8 for d in range(D) if x[i,d])
    diff=h-int(JA[i]); est='OK' if 0<=diff<=C.HOURS_TOL else ('FALTAN' if diff<0 else 'SOBRAN')
    print(f'  {p:15s} días {int(x[i].sum()):3d}  horas {h:5d}  JA {int(JA[i]):5d}  diferencia {diff:+4d}  {est}')
mismatch=[(people[i],str(date[d]),code[i,d],M.role[i,d]) for i in range(P) for d in range(D) if x[i,d] and code[i,d]!=M.role[i,d]]
print('\nDías trabajados en un turno distinto al de su semana:',mismatch)
if len(sys.argv)>3: save_solution(x,sys.argv[3])
