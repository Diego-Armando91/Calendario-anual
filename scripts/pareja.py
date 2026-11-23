"""Cruza el calendario de Diego (Excel editable 'Propuesta anual') con el de su pareja (PDF de Inditex 'Consulta Calendario')
y genera un informe PDF de solapes de turno y problemas para el cuidado de los hijos.
  python pareja.py ../data/Propuesta_Cuadrante_2027_v13_Diego_13.xlsx ../data/pilar/Calendario_2027_pilar.pdf Pilar
Criterio (con 30 min de desplazamiento): Diego mañana 06:45-15:00 -> fuera 06:15-15:30; noche 23:00-07:00 -> fuera 22:30-07:30
(y después duerme hasta ~14:00). La pareja: su horario ±30 min. Hay SOLAPE cuando los dos están fuera a la vez."""
import sys, os, re, collections, datetime as dt, openpyxl, pdfplumber
import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from paths import OUT, WORK
XL, PDFP = sys.argv[1], sys.argv[2]; NOMBRE = sys.argv[3] if len(sys.argv) > 3 else 'Pareja'
Y = 2027; D0 = dt.date(Y, 1, 1); DAYS = [D0 + dt.timedelta(i) for i in range(365)]
MES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre']
MC = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic']
DS = ['lunes','martes','miércoles','jueves','viernes','sábado','domingo']
f = lambda d: f"{DS[d.weekday()][:3]} {d.day} {MC[d.month-1]}"
CUT = dt.date(Y, 9, 15)   # desde aquí la pareja tiene jornada reducida: mañana 10-14, tarde 16-20 (no se cuentan solapes)
RED = {'M': ('10:00', '14:00'), 'T': ('16:00', '20:00')}

# ---------------- calendario de la pareja (PDF Inditex)
pg = pdfplumber.open(PDFP).pages[0]
words = pg.extract_words(extra_attrs=['size'])
heads = [w for w in words if w['text'].upper() in [m.upper() for m in MES] and w['size'] > 8]
rects = pg.rects
def fill_at(x, y):
    best = None
    for r in rects:
        if r['x0'] - 0.5 <= x <= r['x1'] + 0.5 and r['top'] - 0.5 <= y <= r['bottom'] + 0.5:
            c = tuple(round(v, 2) for v in (r.get('non_stroking_color') or ())) if isinstance(r.get('non_stroking_color'), (tuple, list)) else None
            area = (r['x1'] - r['x0']) * (r['bottom'] - r['top'])
            if c and (best is None or area < best[0]) and area > 200: best = (area, c)
    return best[1] if best else None
P = {}   # fecha -> dict(ini, fin, tipo, color)
for h in heads:
    m = [x.upper() for x in MES].index(h['text'].upper()) + 1
    xc_h = h['x0'] + 20; bx0, bx1 = (36, 210) if xc_h < 210 else ((210, 385) if xc_h < 385 else (385, 570)); by0 = h['top'] + 30; by1 = by0 + 140
    days = [w for w in words if bx0 <= w['x0'] <= bx1 and by0 <= w['top'] <= by1 and w['text'].isdigit() and 6 < w['size'] < 7]
    rows = sorted({round(w['top']) for w in days})
    for w in days:
        n = int(w['text']); ri = rows.index(round(w['top']))
        if (ri == 0 and n > 20) or (ri >= 4 and n < 15): continue      # días de otro mes
        try: d = dt.date(Y, m, n)
        except ValueError: continue
        xc = w['x0'] + 3
        tw = [t for t in words if abs(t['top'] - (w['top'] + 7.6)) < 2 and re.match(r'^\d\d:\d\d$', t['text']) and xc - 14 <= t['x0'] <= xc + 14]
        tw.sort(key=lambda t: t['x0'])
        col = fill_at(xc, w['top'] + 3)
        ini = fin = None
        if len(tw) >= 2: ini, fin = tw[0]['text'], tw[1]['text']
        P[d] = dict(ini=ini, fin=fin, color=col)
missing=[d for d in DAYS if d not in P]; print("faltan",missing) if missing else None
def ptipo(d):
    x = P[d]
    if x['ini']:
        h = int(x['ini'][:2]); return 'M' if h < 12 else 'T'
    c = x['color']
    if c == (0.57, 0.82, 0.31): return 'VAC'
    if c == (0.89, 0.31, 0.13): return 'FEST'
    return 'LIB'
# ---------------- calendario de Diego (Excel)
ws = openpyxl.load_workbook(XL)['Propuesta anual']
row = [r for r in range(1, 40) if str(ws.cell(r, 1).value).strip().upper() == 'DIEGO ARMANDO'][0]
CL = {'M': 'M', 'MS': 'M', 'MD': 'M', 'N': 'N', 'NV': 'N', 'NS': 'N', 'V': 'VAC'}
DG = {d: CL.get(str(ws.cell(row, 4 + i).value or '').strip().upper(), 'LIB') for i, d in enumerate(DAYS)}
# ---------------- intervalos (minutos desde el 1 de enero)
def mins(d, hhmm):
    h, m = map(int, hhmm.split(':')); return ((d - D0).days * 24 + h) * 60 + m
def p_int(d):
    x = P[d]
    if not x['ini']: return None
    a = mins(d, x['ini']); b = mins(d, x['fin']) if x['fin'] != '00:00' else mins(d + dt.timedelta(1), '00:00')
    return (a - 30, b + 30)
def d_int(d):
    t = DG[d]
    if t == 'M': return (mins(d, '06:45') - 30, mins(d, '15:00') + 30)
    if t == 'N': return (mins(d, '23:00') - 30, mins(d + dt.timedelta(1), '07:00') + 30)
    return None
eventos = []   # (fecha_ref, tipo, detalle, minutos de solape)
TIP_RAW = collections.OrderedDict([
    ('MM', 'Mismo turno de mañana (los dos fuera toda la mañana)'),
    ('NM', 'Diego sale de noche y la pareja entra de mañana'),
    ('NT', 'Diego entra de noche y la pareja está de tarde'),
    ('MT', 'Relevo mañana → tarde imposible (Diego vuelve y la pareja ya se ha ido)'),
    ('JUSTO', 'Relevo muy justo (menos de 30 min entre la llegada de uno y la salida del otro)'),
    ('SUEÑO', 'La pareja trabaja mientras Diego duerme tras la noche'),
])
TIP = collections.OrderedDict((k, v.replace('la pareja', NOMBRE).replace('La pareja', NOMBRE)) for k, v in TIP_RAW.items())
def hh(m): m %= 1440; return f"{m//60:02d}:{m%60:02d}"
for i, d in enumerate(DAYS):
    if d >= CUT: continue
    pi = p_int(d); di = d_int(d)
    prev = DAYS[i - 1] if i else None
    dprev = d_int(prev) if prev else None
    # mismo día
    if pi and di:
        ov = min(pi[1], di[1]) - max(pi[0], di[0])
        if DG[d] == 'M' and ptipo(d) == 'M': eventos.append((d, 'MM', f"Diego 06:45-15:00 · {NOMBRE} {P[d]['ini']}-{P[d]['fin']}", ov))
        elif DG[d] == 'M' and ptipo(d) == 'T':
            if ov > 0: eventos.append((d, 'MT', f"Diego 06:45-15:00 · {NOMBRE} {P[d]['ini']}-{P[d]['fin']}", ov))
            else: eventos.append((d, 'JUSTO', f"Diego llega ~15:30 · {NOMBRE} sale ~{hh(pi[0])} ({P[d]['ini']}-{P[d]['fin']})", 0))
        elif DG[d] == 'N' and ptipo(d) == 'T' and ov > 0:
            eventos.append((d, 'NT', f"{NOMBRE} {P[d]['ini']}-{P[d]['fin']} · Diego sale ~22:30 hacia la noche", ov))
    # noche de ayer + mañana de hoy
    if prev and DG[prev] == 'N' and pi and ptipo(d) == 'M':
        ov = min(pi[1], dprev[1]) - max(pi[0], dprev[0])
        eventos.append((d, 'NM', f"Diego noche {f(prev)} (sale 07:00) · {NOMBRE} {P[d]['ini']}-{P[d]['fin']}", max(ov, 0)))
    # la pareja trabaja de tarde mientras Diego duerme (noche de ayer) -> solo si entra antes de las 14
    if prev and DG[prev] == 'N' and pi and ptipo(d) == 'T' and pi[0] < mins(d, '14:00'):
        eventos.append((d, 'SUEÑO', f"Diego viene de noche · {NOMBRE} {P[d]['ini']}-{P[d]['fin']}", 0))
# días en los que trabajan los dos / libres los dos
both = [d for d in DAYS if DG[d] in ('M', 'N') and P[d]['ini'] and d < CUT]
both_red = [d for d in DAYS if DG[d] in ('M', 'N') and P[d]['ini'] and d >= CUT]
def red_h(d): a, b = RED[ptipo(d)]; return f'{a}-{b}'
libres = [d for d in DAYS if DG[d] in ('LIB', 'VAC') and not P[d]['ini']]
finde_ambos = [d for d in DAYS if d.weekday() >= 5 and DG[d] in ('M', 'N') and P[d]['ini'] and d < CUT]
problema = sorted({e[0] for e in eventos if e[1] in ('MM', 'NM', 'NT', 'MT')})
cnt = collections.Counter(e[1] for e in eventos)
pm = collections.defaultdict(collections.Counter)
for e in eventos: pm[e[0].month][e[1]] += 1

# ---------------- PDF
ss = getSampleStyleSheet()
H1 = ParagraphStyle('h1', parent=ss['Heading1'], fontSize=16, textColor=colors.HexColor('#1F3864'))
H2 = ParagraphStyle('h2', parent=ss['Heading2'], fontSize=12.5, textColor=colors.HexColor('#2E5597'), spaceBefore=8)
B = ParagraphStyle('b', parent=ss['BodyText'], fontSize=9.2, leading=12.5)
SM = ParagraphStyle('sm', parent=B, fontSize=7.6, leading=9.6)
def tbl(data, widths, fs=7.8, hl=()):
    t = Table(data, colWidths=widths, repeatRows=1)
    st = [('FONT', (0, 0), (-1, -1), 'Helvetica', fs), ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', fs), ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#DCE6F1')),
          ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#B0B0B0')), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
          ('TOPPADDING', (0, 0), (-1, -1), 1.5), ('BOTTOMPADDING', (0, 0), (-1, -1), 1.5)]
    for r in range(2, len(data), 2): st.append(('BACKGROUND', (0, r), (-1, r), colors.HexColor('#F5F7FA')))
    for c, r, col in hl: st.append(('BACKGROUND', (c, r), (c, r), colors.HexColor(col)))
    t.setStyle(TableStyle(st)); return t
story = [Paragraph(f'Turnos de Diego y {NOMBRE} en 2027 — cuidado de los hijos', H1)]
story.append(Paragraph(f'Se cruzan día a día los dos calendarios: el de Diego (cuadrante de expedición) y el de {NOMBRE} (calendario de Inditex, '
    f'{sum(1 for d in DAYS if P[d]["ini"])} días con turno). Se considera que una persona está «fuera de casa» desde 30 minutos antes de entrar hasta 30 minutos después de salir. '
    'Diego: mañana 06:45-15:00 y noche 23:00-07:00 (después de la noche duerme hasta media tarde). '
    f'Hay <b>solape</b> cuando los dos están fuera a la vez y alguien tiene que quedarse con los niños.', B))
story.append(Paragraph(f'<b>Del 15 de septiembre a final de año {NOMBRE} tiene jornada reducida</b> (4 horas: mañana 10:00-14:00, tarde 16:00-20:00). '
    'En ese periodo no se cuentan solapes: solo se marcan los días en que trabajáis los dos (apartado propio al final). Todos los recuentos de solapes son del 1 de enero al 14 de septiembre.', B))
story.append(Paragraph('Resumen', H2))
data = [['Situación', 'Días']]; hl = []
for k, lab in TIP.items():
    data.append([Paragraph(lab, SM), cnt.get(k, 0)])
    if k in ('MM', 'NM', 'NT', 'MT') and cnt.get(k, 0): hl.append((1, len(data) - 1, '#F8CBAD'))
data.append([Paragraph('<b>Días con algún solape (necesitan ayuda con los niños)</b>', SM), len(problema)]); hl.append((1, len(data) - 1, '#F8CBAD'))
data.append([Paragraph('Días que trabajáis los dos (en cualquier turno), hasta el 14 sep', SM), len(both)])
data.append([Paragraph(f'Días que trabajáis los dos del 15 sep a fin de año ({NOMBRE} con jornada reducida)', SM), len(both_red)])
data.append([Paragraph('Días que <b>compartís turno</b> (los dos de mañana el mismo día)', SM), cnt.get('MM', 0)])
data.append([Paragraph('Fines de semana (sábado o domingo) que trabajáis los dos, hasta el 14 sep', SM), len(finde_ambos)])
data.append([Paragraph('Días que libráis los dos (incluye vacaciones)', SM), len(libres)])
story.append(tbl(data, [13 * cm, 2.5 * cm], fs=8, hl=hl))
story.append(Paragraph('En naranja, las situaciones en que los niños se quedan sin ninguno de los dos.', SM))
# por mes
story.append(Paragraph('Solapes por mes', H2))
data = [['Mes', 'Mañana-mañana', 'Noche→mañana', 'Noche-tarde', 'Mañana→tarde', 'Relevo justo', 'Total con solape']]
for m in range(1, 10):
    c = pm[m]; tot = len({e[0] for e in eventos if e[0].month == m and e[1] in ('MM', 'NM', 'NT', 'MT')})
    data.append([MES[m - 1].capitalize() + (' (1-14)' if m == 9 else ''), c['MM'], c['NM'], c['NT'], c['MT'], c['JUSTO'], tot])
data.append(['TOTAL', cnt['MM'], cnt['NM'], cnt['NT'], cnt['MT'], cnt['JUSTO'], len(problema)])
story.append(tbl(data, [2.6 * cm, 2.4 * cm, 2.4 * cm, 2.2 * cm, 2.4 * cm, 2.2 * cm, 2.6 * cm], fs=7.8))
# gráfico
fig, ax = plt.subplots(figsize=(9.5, 3.2)); x = range(9); bot = [0] * 9
for k, lab, colr in (('MM', 'Mañana-mañana', '#A9D18E'), ('NM', 'Noche→mañana', '#E97132'), ('NT', 'Noche-tarde', '#C9A7E0'), ('MT', 'Mañana→tarde', '#4472C4')):
    v = [pm[m + 1][k] for m in x]; ax.bar(x, v, bottom=bot, label=lab, color=colr); bot = [a + b for a, b in zip(bot, v)]
ax.set_xticks(list(x)); ax.set_xticklabels([m[:3] for m in MES[:9]]); ax.spines[['top', 'right']].set_visible(False)
ax.legend(ncol=4, fontsize=8, frameon=False, loc='upper center', bbox_to_anchor=(0.5, 1.15)); ax.set_ylabel('días')
plt.tight_layout(); im = os.path.join(WORK, '_pareja.png'); plt.savefig(im, dpi=170); plt.close()
story.append(Spacer(1, 4)); story.append(Image(im, width=17 * cm, height=5.7 * cm))
# por día de la semana
story.append(Paragraph('Hasta el 14 de septiembre. Desde el 15 de septiembre no se cuentan solapes (jornada reducida).', SM))
story.append(Paragraph('¿Qué días de la semana fallan más?', H2))
wdc = collections.Counter(d.weekday() for d in problema)
story.append(Paragraph(', '.join(f"{DS[k]}: <b>{wdc.get(k, 0)}</b>" for k in range(7)) + '.', B))
# detalle
story.append(PageBreak()); story.append(Paragraph('Detalle de los días con problemas', H1))
for k in ('NM', 'MM', 'NT', 'MT', 'JUSTO', 'SUEÑO'):
    ev = [e for e in eventos if e[1] == k]
    if not ev: continue
    story.append(Paragraph(f'{TIP[k]} — {len(ev)} días', H2))
    data = [['Fecha', 'Qué pasa', 'Tiempo sin nadie']]
    for d, _, det, ov in ev: data.append([f(d), Paragraph(det, SM), f"{ov // 60}h {ov % 60:02d}m" if ov > 0 else '—'])
    story.append(tbl(data, [3 * cm, 11.5 * cm, 2.6 * cm], fs=7.4))
# fines de semana
story.append(PageBreak()); story.append(Paragraph('Fines de semana', H1))
sats = [d for d in DAYS if d.weekday() == 5]
data = [['Fin de semana', 'Diego', NOMBRE, '¿Algún día los dos fuera?']]
for s in sats:
    sun = s + dt.timedelta(1)
    if sun.year != Y: continue
    dd = '/'.join(DG[x] if DG[x] != 'LIB' else '-' for x in (s, sun))
    pp = '/'.join((red_h(x) if x >= CUT else P[x]['ini'] + '-' + P[x]['fin']) if P[x]['ini'] else ('VAC' if ptipo(x) == 'VAC' else '-') for x in (s, sun))
    prob = [x for x in (s, sun) if x in problema]
    if DG[s] in ('M', 'N') or DG[sun] in ('M', 'N') or P[s]['ini'] or P[sun]['ini']:
        col4 = ', '.join(DS[x.weekday()] for x in prob) or ''
        if s >= CUT: col4 = ('trabajáis los dos: ' + ', '.join(DS[x.weekday()] for x in (s, sun) if x in both_red)) if any(x in both_red for x in (s, sun)) else ''
        data.append([f"{s.day}-{sun.day} {MC[sun.month-1]}", dd, pp, col4])
story.append(tbl(data, [3 * cm, 2.6 * cm, 6.4 * cm, 5 * cm], fs=7.2))
story.append(Paragraph(f'Desde el 15 de septiembre, horario reducido de {NOMBRE} y solo se indica si trabajáis los dos. Diego: M mañana, N noche, VAC vacaciones, - libre. Se listan los fines de semana en que trabaja al menos uno de los dos.', SM))
# jornada reducida
story.append(PageBreak()); story.append(Paragraph(f'Del 15 de septiembre a final de año — días que trabajáis los dos', H1))
story.append(Paragraph(f'{NOMBRE} trabaja 4 horas: de mañana 10:00-14:00 y de tarde 16:00-20:00. No se cuentan solapes; solo se marcan los días en que trabajáis los dos: '
    f'<b>{len(both_red)} días</b>.', B))
cm_ = collections.Counter((d.month, DG[d], ptipo(d)) for d in both_red)
data = [['Mes'] + [Paragraph(f'<b>Diego {a}<br/>{NOMBRE} {b}</b>', SM) for a, b in (('mañana','mañana'),('mañana','tarde'),('noche','mañana'),('noche','tarde'))] + ['Total']]
for m in range(9, 13):
    data.append([MES[m-1].capitalize() + (' (15-30)' if m == 9 else ''), cm_[(m,'M','M')], cm_[(m,'M','T')], cm_[(m,'N','M')], cm_[(m,'N','T')], sum(1 for d in both_red if d.month == m)])
data.append(['TOTAL', sum(cm_[(m,'M','M')] for m in range(9,13)), sum(cm_[(m,'M','T')] for m in range(9,13)), sum(cm_[(m,'N','M')] for m in range(9,13)), sum(cm_[(m,'N','T')] for m in range(9,13)), len(both_red)])
story.append(tbl(data, [2.6*cm, 3.3*cm, 3.3*cm, 3.3*cm, 3.3*cm, 1.6*cm], fs=7.2))
story.append(Paragraph('Diego noche = la noche que empieza ese día a las 23:00.', SM))
story.append(Spacer(1, 6))
data = [['Fecha', 'Diego', NOMBRE + ' (reducida)']]
for d in both_red: data.append([f(d), {'M': 'Mañana 06:45-15:00', 'N': 'Noche 23:00-07:00'}[DG[d]], ('Mañana ' if ptipo(d) == 'M' else 'Tarde ') + red_h(d)])
story.append(tbl(data, [3.4*cm, 5*cm, 5*cm], fs=7.4))
# conclusiones
story.append(PageBreak()); story.append(Paragraph('Qué nos complica más el cuidado de los niños', H1))
tops = sorted(((cnt[k], TIP[k]) for k in ('NM', 'MM', 'NT', 'MT')), reverse=True)
for n, lab in tops:
    if n: story.append(Paragraph(f'• <b>{lab}</b>: {n} días.', B))
story.append(Paragraph('Lecturas útiles', H2))
mes_peor = max(range(1, 10), key=lambda m: len({e[0] for e in eventos if e[0].month == m and e[1] in ('MM', 'NM', 'NT', 'MT')}))
story.append(Paragraph(f"• El mes más complicado es <b>{MES[mes_peor-1]}</b>.", B))
story.append(Paragraph(f"• Del 1 de enero al 14 de septiembre hay <b>{len(problema)} días con solape</b>. Del 15 de septiembre a final de año, con la jornada reducida de {NOMBRE}, coincidís trabajando {len(both_red)} días.", B))
story.append(Paragraph(f"• Hasta el 14 de septiembre, los dos trabajáis a la vez el fin de semana {len(finde_ambos)} días: esos días no hay colegio y hace falta ayuda seguro.", B))
story.append(Paragraph(f'• «Noche → mañana»: Diego llega a casa sobre las 07:30 y {NOMBRE} tiene que salir sobre las 06:30 (o 07:30 en domingo): '
    'durante ese hueco y la mañana (mientras Diego duerme) los niños necesitan a alguien.', B))
story.append(Paragraph(f'• «Noche-tarde»: {NOMBRE} vuelve a las 23:30 (00:30 si sale a las 00:00) y Diego se va a las 22:30: queda una hora (o dos) sin nadie y la acostada de los niños.', B))
story.append(Paragraph(f'• «Mañana → tarde»: Diego vuelve a las 15:30 y {NOMBRE} tiene que salir a las 14:30: falta una hora a mediodía (comida y salida del colegio).', B))
story.append(Paragraph('Qué se podría pedir para reducirlos', H2))
story.append(Paragraph('No hay una combinación de turnos sin ningún hueco, pero unas son mucho peores que otras:', B))
story.append(Paragraph(f'• <b>Lo peor</b>: Diego de noche y {NOMBRE} de mañana al día siguiente (falta toda la mañana, porque Diego tiene que dormir) y los dos de mañana el mismo día (falta toda la mañana).', B))
story.append(Paragraph(f'• <b>Menos malo</b>: Diego de mañana y {NOMBRE} de tarde (falta 1 hora a mediodía) o Diego de noche y {NOMBRE} de tarde con salida a las 23:00 (falta 1 hora por la noche, con los niños ya acostados).', B))
story.append(Paragraph(f'• <b>Lo ideal</b>: que las semanas de noche de Diego coincidan con los días libres o de tarde de {NOMBRE}, que sus semanas de mañana no coincidan, '
    'y que el fin de semana no trabajéis los dos el mismo día. Con eso desaparecerían casi todos los días de «noche → mañana» y «mañana-mañana», que son los que más tiempo dejan a los niños sin nadie.', B))
out = os.path.join(OUT, f'Informe_Turnos_Diego_{NOMBRE}_2027.pdf')
SimpleDocTemplate(out, pagesize=A4, leftMargin=1.4 * cm, rightMargin=1.4 * cm, topMargin=1.3 * cm, bottomMargin=1.3 * cm, title=f'Turnos Diego y {NOMBRE} 2027').build(story)
print('ok', out); print('pareja días con turno', sum(1 for d in DAYS if P[d]['ini']), 'tipos', collections.Counter(ptipo(d) for d in DAYS))
print('both_red', len(both_red)); print('eventos', dict(cnt), 'días con solape', len(problema), 'ambos trabajan', len(both), 'finde ambos', len(finde_ambos), 'libres ambos', len(libres))
