"""PDF con los días del año en que un turno tiene 5 personas (desde la hoja 'Propuesta anual').
Los días con A (acumulación) cuentan como no trabajados.
  python cinco_pdf.py ../data/Propuesta_Cuadrante_2027_v13_Diego_17.xlsx"""
import sys, os, collections, datetime as dt, openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from paths import OUT
ws = openpyxl.load_workbook(sys.argv[1], data_only=True)['Propuesta anual']
s = dt.date(2027, 1, 1); MC = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic']
MES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
DS = ['lun','mar','mié','jue','vie','sáb','dom']; DL = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo']
rows = [r for r in range(5, 20) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip()]
G = [('M7', 'Mañana a las 7:00 (sin la central)', ('M','MS','MD')), ('MC', 'Mañana contando la central', ('M','MS','MD','C')),
     ('T', 'Tarde', ('T','TS','TD')), ('N', 'Noche (empieza ese día a las 23:00)', ('N','NV','NS'))]
cnt = {k: collections.Counter() for k, _, _ in G}; days = {k: [] for k, _, _ in G}
for d in range(365):
    dd = s + dt.timedelta(d); c = [str(ws.cell(r, 4 + d).value or '').strip().upper() for r in rows]
    for k, _, codes in G:
        n = sum(x in codes for x in c); cnt[k][n] += 1
        if n >= 5: days[k].append(dd)
ss = getSampleStyleSheet()
H1 = ParagraphStyle('h1', parent=ss['Heading1'], fontSize=16, textColor=colors.HexColor('#1F3864'))
H2 = ParagraphStyle('h2', parent=ss['Heading2'], fontSize=12, textColor=colors.HexColor('#2E5597'), spaceBefore=8)
B = ParagraphStyle('b', parent=ss['BodyText'], fontSize=9, leading=12)
SM = ParagraphStyle('sm', parent=B, fontSize=7.6, leading=9.5)
def tbl(data, w, fs=8, hl=()):
    t = Table(data, colWidths=w, repeatRows=1)
    st = [('FONT', (0,0), (-1,-1), 'Helvetica', fs), ('FONT', (0,0), (-1,0), 'Helvetica-Bold', fs),
          ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#DCE6F1')), ('GRID', (0,0), (-1,-1), .25, colors.HexColor('#B0B0B0')),
          ('ALIGN', (1,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
          ('TOPPADDING', (0,0), (-1,-1), 1.5), ('BOTTOMPADDING', (0,0), (-1,-1), 1.5)]
    for r in range(2, len(data), 2): st.append(('BACKGROUND', (0,r), (-1,r), colors.HexColor('#F5F7FA')))
    for c, r in hl: st.append(('BACKGROUND', (c,r), (c,r), colors.HexColor('#C6E0B4')))
    t.setStyle(TableStyle(st)); return t
story = [Paragraph('Días de 2027 con 5 personas por turno', H1),
         Paragraph('Para cada día del año se cuenta cuántas personas trabajan en cada turno. La <b>mañana</b> se da de dos formas: los que entran a las 7:00 '
                   '(sin la central) y contando también la central, que entra a las 10:00. La <b>noche</b> de un día es la que empieza ese día a las 23:00. '
                   'Los días marcados con <b>A</b> (acumulación) no se cuentan como trabajados. En ningún turno se pasa de 5 personas.', B)]
story.append(Paragraph('Resumen del año', H2))
data = [['Turno', 'Días con 5', 'Con 4', 'Con 3', 'Con 2', 'Con 1', 'Sin nadie']]
for k, lab, _ in G: data.append([lab] + [cnt[k][v] for v in (5, 4, 3, 2, 1, 0)])
story.append(tbl(data, [6.4*cm] + [1.8*cm]*6, hl=[(1, r) for r in range(1, 5)]))
story.append(Paragraph('Los días «sin nadie» son cierres, festivos, la noche anterior a un festivo y, en la noche, los sábados fuera de temporada.', SM))
story.append(Paragraph('Días con 5 por mes', H2))
data = [['Mes'] + [lab.split(' (')[0] for k, lab, _ in G]]
for m in range(1, 13): data.append([MES[m-1]] + [sum(1 for d in days[k] if d.month == m) for k, _, _ in G])
data.append(['TOTAL'] + [len(days[k]) for k, _, _ in G])
story.append(tbl(data, [3*cm, 4.6*cm, 4.6*cm, 2.6*cm, 2.6*cm]))
story.append(Paragraph('Días con 5 por día de la semana', H2))
data = [['Día'] + [lab.split(' (')[0] for k, lab, _ in G]]
for w in range(7): data.append([DL[w]] + [sum(1 for d in days[k] if d.weekday() == w) for k, _, _ in G])
story.append(tbl(data, [3*cm, 4.6*cm, 4.6*cm, 2.6*cm, 2.6*cm]))
story.append(Spacer(1, 4))
story.append(Paragraph('• La mañana a las 7:00 nunca llega a 5: su máximo son 4 personas. Con la central sí llega a 5 algunos días.', B))
story.append(Paragraph('• Los días de 5 se concentran al principio de la semana (lunes y martes en mañana y tarde; martes y miércoles en la noche), '
                       'porque los libres se acumulan alrededor del fin de semana.', B))
story.append(Paragraph('• En julio y agosto casi no hay días de 5 por las vacaciones.', B))
story.append(PageBreak()); story.append(Paragraph('Lista de días con 5 personas', H1))
for k, lab, _ in G:
    if not days[k]: continue
    story.append(Paragraph(f'{lab} — {len(days[k])} días', H2))
    data = [['Mes', 'Días']]
    for m in range(1, 13):
        dm = [d for d in days[k] if d.month == m]
        if dm: data.append([MES[m-1], Paragraph(', '.join(f'{DS[d.weekday()]} {d.day}' for d in dm), SM)])
    t = tbl(data, [3*cm, 14*cm]); t.setStyle(TableStyle([('ALIGN', (1,1), (1,-1), 'LEFT')])); story.append(t)
out = os.path.join(OUT, 'Dias_5_personas_2027.pdf')
SimpleDocTemplate(out, pagesize=A4, leftMargin=1.6*cm, rightMargin=1.6*cm, topMargin=1.2*cm, bottomMargin=1.2*cm, title='Días con 5 personas por turno 2027').build(story)
print('ok', out, {k: len(v) for k, v in days.items()})
