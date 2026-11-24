"""Calendario anual (PNG) con los turnos de Diego y de su pareja en cada día: arriba Diego, abajo la pareja.
Borde rojo: solape con los niños (hasta el 14 sep, incluye salir de noche y que la pareja entre de mañana); borde azul: trabajan los dos sin solape.
  python calendario_pareja.py ../data/CALENDARIO2027V1.1_encargado_PA.xlsx ../data/pilar/Calendario_2027_pilar.pdf Pilar V1.1"""
import sys, os, calendar, datetime as dt, openpyxl
TAG = sys.argv.pop(4) if len(sys.argv) > 4 else ''
from pareja_datos import *
from matplotlib.patches import Rectangle
from paths import OUT
# códigos de Diego tal cual están en el Excel (MS, NV...) y días rojos
ws = openpyxl.load_workbook(XL)['Propuesta anual']
row = [r for r in range(1, 40) if str(ws.cell(r, 1).value).strip().upper() == 'DIEGO ARMANDO'][0]
DCODE, DROJO = {}, set()
for i, d in enumerate(DAYS):
    c = ws.cell(row, 4 + i); v = str(c.value or '').strip().upper()
    DCODE[d] = v if v in ('M','MS','MD','N','NV','NS') else ''
    fl = c.fill
    if fl is not None and fl.fill_type == 'solid' and fl.fgColor.type == 'rgb' and str(fl.fgColor.rgb).endswith('FF0000'): DROJO.add(d)
COL = {'M': '#A9D18E', 'T': '#C9A7E0', 'N': '#E97132', 'VAC': '#404040', 'FEST': '#FF0000', 'LIB': '#F2F2F2'}
def dk(d):
    if DG[d] == 'VAC': return 'VAC'
    if DG[d] in ('M', 'N'): return DG[d]
    return 'FEST' if d in DROJO else 'LIB'
def pk(d): return ptipo(d)
def plab(d):
    if not P[d]['ini']: return ''
    if d >= CUT: a, b = RED[ptipo(d)]
    else: a, b = P[d]['ini'], P[d]['fin']
    return f"{a[:2]}-{b[:2]}"
PROB = set(problema); BOTH = {d for d in DAYS if DG[d] in ('M', 'N') and P[d]['ini']}
LIBRES = {d for d in DAYS if DG[d] in ('LIB', 'VAC') and not P[d]['ini']}
MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
fig = plt.figure(figsize=(22, 19.5)); gs = fig.add_gridspec(4, 3, hspace=0.32, wspace=0.12, top=0.905, bottom=0.14, left=0.03, right=0.97)
for m in range(1, 13):
    ax = fig.add_subplot(gs[(m-1)//3, (m-1)%3]); mat = calendar.monthcalendar(Y, m)
    ax.set_xlim(0, 7); ax.set_ylim(0, len(mat) + 1.1); ax.invert_yaxis(); ax.axis('off'); ax.set_title(MESES[m-1], fontsize=20, fontweight='bold', pad=6)
    for c, l in enumerate('LMXJVSD'): ax.text(c + .5, .5, l, ha='center', va='center', fontsize=12, fontweight='bold', color='#B00020' if c >= 5 else '#333')
    for r, week in enumerate(mat):
        for c, day in enumerate(week):
            if not day: continue
            d = dt.date(Y, m, day); y = r + 1; a, b = dk(d), pk(d)
            ax.add_patch(Rectangle((c, y), 1, .5, facecolor=COL[a], edgecolor='#BBBBBB', linewidth=.5))
            ax.add_patch(Rectangle((c, y + .5), 1, .5, facecolor=COL[b], edgecolor='#BBBBBB', linewidth=.5))
            ta = 'white' if a in ('VAC',) else '#111'; tb = 'white' if b in ('VAC',) else '#111'
            ax.text(c + .06, y + .17, str(day), ha='left', va='center', fontsize=11, fontweight='bold', color=ta)
            if DCODE[d]: ax.text(c + .62, y + .3, DCODE[d], ha='center', va='center', fontsize=10, color=ta)
            if plab(d): ax.text(c + .5, y + .76, plab(d), ha='center', va='center', fontsize=9.5, color=tb)
            if d in PROB: ax.add_patch(Rectangle((c + .04, y + .04), .92, .92, facecolor='none', edgecolor='#C00000', linewidth=3.8))
            elif d in BOTH: ax.add_patch(Rectangle((c + .04, y + .04), .92, .92, facecolor='none', edgecolor='#1F3864', linewidth=3.2, linestyle=(0, (2, 1))))
            if d in LIBRES and a != 'FEST': ax.text(c + .9, y + .17, '★', ha='center', va='center', fontsize=10, color='#2E7D32' if a != 'VAC' else 'white')
fig.suptitle(f'Calendario 2027 — Diego y {NOMBRE}' + (f' · {TAG}' if TAG else ''), fontsize=28, fontweight='bold', y=0.988)
fig.text(0.5, 0.95, f'En cada día: arriba Diego (código de turno), abajo {NOMBRE} (horario). Desde el 15 de septiembre, {NOMBRE} con jornada reducida (10-14 / 16-20).',
         ha='center', fontsize=13, color='#333')
h = [Rectangle((0,0),1,1,facecolor=COL['M'],edgecolor='#999',label='Mañana'), Rectangle((0,0),1,1,facecolor=COL['T'],edgecolor='#999',label=f'Tarde ({NOMBRE})'),
     Rectangle((0,0),1,1,facecolor=COL['N'],edgecolor='#999',label='Noche (Diego)'), Rectangle((0,0),1,1,facecolor=COL['VAC'],edgecolor='#999',label='Vacaciones'),
     Rectangle((0,0),1,1,facecolor=COL['FEST'],edgecolor='#999',label='Festivo / cierre'), Rectangle((0,0),1,1,facecolor=COL['LIB'],edgecolor='#999',label='Libre'),
     Rectangle((0,0),1,1,facecolor='white',edgecolor='#C00000',linewidth=3,label='Solape: hace falta ayuda con los niños (hasta 14 sep)'),
     Rectangle((0,0),1,1,facecolor='white',edgecolor='#1F3864',linewidth=3,linestyle=(0,(2,1)),label='Trabajáis los dos (sin solape)'),
     plt.Line2D([], [], marker='*', color='#2E7D32', linestyle='None', markersize=14, label='Libráis los dos')]
fig.legend(handles=h, loc='lower center', ncol=3, fontsize=13, frameon=False, bbox_to_anchor=(0.5, 0.03))
fig.text(0.5, 0.012, f'{len(BOTH)} días trabajáis los dos · {len(PROB)} con solape hasta el 14 de septiembre · {len(LIBRES)} días libráis los dos',
         ha='center', fontsize=13, color='#444')
out = os.path.join(OUT, f'Calendario_2027_Diego_{NOMBRE}' + (f'_{TAG}' if TAG else '') + '.png')
plt.savefig(out, dpi=130, facecolor='white'); print('ok', out, len(BOTH), len(PROB), len(LIBRES))
