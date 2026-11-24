"""Datos compartidos por pareja.py y calendario_pareja.py: lee el turno de Diego (Excel 'Propuesta anual') y el de la pareja
(PDF de Inditex) y calcula los solapes. Lee sys.argv: <excel> <pdf pareja> [nombre]."""
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

