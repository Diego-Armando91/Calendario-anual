"""Convierte un calendario con hoja 'Rotación anual' (formato original, vacaciones en negro) al formato 'Propuesta anual'
(códigos, V = vacaciones) para poder pasarle informe_pdf.py y calendario_persona.py.
  python convertir_encargado.py ../data/CALENDARIO2027V1.1_encargado.xlsx ../data/CALENDARIO2027V1.1_encargado_PA.xlsx [PERSONA,desde,hasta]
El tercer argumento (opcional) marca como A (acumulación) los días laborables sin turno de ese bloque (p. ej. el bloque en rojo de reducción de jornada)."""
import sys, copy, datetime as dt, openpyxl
from openpyxl.styles import PatternFill
from common import people
src = openpyxl.load_workbook(sys.argv[1])['Rotación anual']
wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Propuesta anual'
ws.cell(4, 1, 'NOMBRE'); s = dt.date(2027, 1, 1)
for d in range(365):
    ws.cell(3, 4 + d, dt.datetime.combine(s + dt.timedelta(d), dt.time())); ws.cell(4, 4 + d, 'LMXJVSD'[(s + dt.timedelta(d)).weekday()])
r = 5; conflict = []; nA = 0
ACU = None
if len(sys.argv) > 3:
    _p, _a, _b = sys.argv[3].split(','); ACU = (_p.upper(), dt.date.fromisoformat(_a), dt.date.fromisoformat(_b))
for sr in range(4, 19):
    n = str(src.cell(sr, 1).value).strip().upper()
    assert n in people, n
    ws.cell(r, 1, n); ws.cell(r, 2, src.cell(sr, 2).value); ws.cell(r, 3, src.cell(sr, 30 - 27 + 0).value if False else None)
    for d in range(365):
        c = src.cell(sr, 4 + d); v = c.value.strip().upper() if isinstance(c.value, str) and c.value.strip() else None
        f = c.fill; fc = None
        if f is not None and f.fill_type == 'solid': fc = f.fgColor.rgb if f.fgColor.type == 'rgb' else f'theme{f.fgColor.theme}'
        o = ws.cell(r, 4 + d)
        if fc in ('theme1', 'FF000000'):
            if v: conflict.append((n, str(s + dt.timedelta(d)), v))
            o.value = 'V'
        else: o.value = v
        dd = s + dt.timedelta(d)
        if ACU and n == ACU[0] and ACU[1] <= dd <= ACU[2] and not v and fc not in ('theme1', 'FF000000') and dd.weekday() < 5:
            o.value = 'A'; nA += 1; continue
        if fc and fc.endswith('FF0000'): o.fill = PatternFill('solid', fgColor='FFFF0000')
        elif fc == 'FF00B0F0': o.fill = PatternFill('solid', fgColor='FF00B0F0')
    r += 1
wb.save(sys.argv[2]); print('ok', sys.argv[2], 'vacaciones con código:', conflict, 'días A:', nA)
