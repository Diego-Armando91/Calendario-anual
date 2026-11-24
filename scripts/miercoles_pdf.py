"""PDF con cuántas personas trabajan cada miércoles en mañana, tarde y noche (desde el Excel 'Propuesta anual').
  python miercoles_pdf.py ../data/Propuesta_Cuadrante_2027_v13_Diego_17.xlsx"""
import sys, os, collections, datetime as dt, openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from paths import OUT
ws = openpyxl.load_workbook(sys.argv[1], data_only=True)['Propuesta anual']
s = dt.date(2027, 1, 1); MC = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic']
rows = [r for r in range(5, 20) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip()]
NOTA = {dt.date(2027,1,6): ('festivo', ''), dt.date(2027,3,24): ('', 'víspera de cierre')}
data = [['Miércoles', 'Mañana\n(a las 7:00)', 'Central\n(10:00)', 'Tarde', 'Noche\n(mié 23:00 → jue)']]; hl = []
stats = collections.defaultdict(list)
for d in range(365):
    dd = s + dt.timedelta(d)
    if dd.weekday() != 2: continue
    c = [str(ws.cell(r, 4 + d).value or '').strip().upper() for r in rows]
    m, cc, t, n = (sum(x == k for x in c) for k in ('M', 'C', 'T', 'N'))
    nm, nn = NOTA.get(dd, ('', ''))
    data.append([f"{dd.day} {MC[dd.month-1]}", f"{m}" + (f" ({nm})" if nm else ''), 'sí' if cc else '—', t, f"{n}" + (f" ({nn})" if nn else '')])
    k = len(data) - 1
    if not nm:
        stats['m'].append(m); stats['t'].append(t)
        if m <= 2: hl.append((1, k))
        if t <= 2: hl.append((3, k))
    else: hl.append((1, k)); hl.append((3, k))
    if not nn and not nm:
        stats['n'].append(n)
        if n <= 2: hl.append((4, k))
    elif nn: hl.append((4, k))
ss = getSampleStyleSheet()
H1 = ParagraphStyle('h1', parent=ss['Heading1'], fontSize=16, textColor=colors.HexColor('#1F3864'))
H2 = ParagraphStyle('h2', parent=ss['Heading2'], fontSize=12, textColor=colors.HexColor('#2E5597'))
B = ParagraphStyle('b', parent=ss['BodyText'], fontSize=9, leading=12)
SM = ParagraphStyle('sm', parent=B, fontSize=7.6, leading=9.5)
def tbl(data, w, fs, hl=()):
    t = Table(data, colWidths=w, repeatRows=1)
    st = [('FONT', (0,0), (-1,-1), 'Helvetica', fs), ('FONT', (0,0), (-1,0), 'Helvetica-Bold', fs),
          ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#DCE6F1')), ('GRID', (0,0), (-1,-1), .25, colors.HexColor('#B0B0B0')),
          ('ALIGN', (1,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
          ('TOPPADDING', (0,0), (-1,-1), 0.6), ('BOTTOMPADDING', (0,0), (-1,-1), 0.6)]
    for r in range(2, len(data), 2): st.append(('BACKGROUND', (0,r), (-1,r), colors.HexColor('#F5F7FA')))
    for c, r in hl: st.append(('BACKGROUND', (c,r), (c,r), colors.HexColor('#F8CBAD')))
    t.setStyle(TableStyle(st)); return t
story = [Paragraph('Miércoles de 2027 — personas por turno', H1),
         Paragraph('Número de personas que trabajan cada miércoles en cada turno. <b>Mañana</b>: los que entran a las 7:00 (M), sin contar la central. '
                   '<b>Central</b>: si ese día está la central de mañana, que entra a las 10:00. <b>Noche</b>: la noche que empieza el miércoles a las 23:00 y termina el jueves. '
                   'En naranja, los turnos con 2 personas o menos.', B), Spacer(1, 4)]
story.append(tbl(data, [2.6*cm, 3.4*cm, 2.2*cm, 2.2*cm, 4.4*cm], 7.6, hl))
story.append(Spacer(1, 8)); story.append(Paragraph('Resumen', H2))
story.append(Paragraph('Sin contar el miércoles 6 de enero (festivo) ni la noche del 24 de marzo (víspera del cierre).', SM))
sm = [['', 'Mañana (7:00)', 'Tarde', 'Noche']]
sm.append(['Media'] + [f"{sum(stats[k])/len(stats[k]):.1f}".replace('.', ',') for k in 'mtn'])
for v in (5, 4, 3, 2):
    sm.append([f'Miércoles con {v}'] + [collections.Counter(stats[k])[v] for k in 'mtn'])
story.append(tbl(sm, [3.6*cm, 3*cm, 3*cm, 3*cm], 8))
story.append(Spacer(1, 6))
cm2 = sum(1 for k in range(1, len(data)) if data[k][1] == '2'); cm2c = sum(1 for k in range(1, len(data)) if data[k][1] == '2' and data[k][2] == 'sí')
story.append(Paragraph(f'• La <b>mañana</b> es el turno más justo: {cm2} miércoles hay solo 2 personas a las 7:00; en {cm2c} de ellos llega la central a las 10:00.', B))
story.append(Paragraph('• La <b>noche</b> es el turno con más gente entre semana: lo más habitual son 5 personas.', B))
story.append(Paragraph('• El 24 de marzo no hay noche porque es la víspera del cierre del 25 de marzo.', B))
out = os.path.join(OUT, 'Miercoles_2027_por_turno.pdf')
SimpleDocTemplate(out, pagesize=A4, leftMargin=1.6*cm, rightMargin=1.6*cm, topMargin=0.9*cm, bottomMargin=0.8*cm, title='Miércoles 2027 por turno').build(story)
print('ok', out)
