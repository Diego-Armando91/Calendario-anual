"""Todos los parámetros de la propuesta. Cambia aquí las reglas, no en model.py."""
import datetime as dt
D=dt.date
YEAR_START=D(2027,1,1)
# Cierres: nadie trabaja
CLOSED={D(2027,3,25),D(2027,3,26),D(2027,12,24),D(2027,12,25),D(2027,12,31)}
# Noches especiales 00:00-07:00 (7 h): solo personal de noche
SPECIAL={D(2027,1,1),D(2027,1,6),D(2027,3,27),D(2027,10,12)}
# Sábado noche solo se trabaja en estos periodos (ambos incluidos)
SAT_NIGHT_PERIODS=[(D(2027,2,27),D(2027,5,29)),(D(2027,7,17),D(2027,12,4))]
# Fines de semana de mañana/tarde: 1 persona hasta esta fecha, 2 desde ella (formato del cuadrante original)
WEEKEND_TWO_FROM=D(2027,9,4)
# Días con cobertura copiada del cuadrante original (festivos y alrededores)
OVERRIDE_DATES=[D(2027,1,1),D(2027,1,5),D(2027,1,6),D(2027,3,24),D(2027,3,25),D(2027,3,26),D(2027,3,27),D(2027,3,28),
                D(2027,10,11),D(2027,10,12)]+[D(2027,12,x) for x in range(20,32)]
# Roles
FIX={'CHEMA':'M','JUANMA':'M','MICHEL':'T','ALFREDO':'T','IRENE':'T','SOFIA':'T','ARMANDO':'N','JUANMA V':'N','RUBEN S':'N'}
PAIRS=[('ESTHER','M','DAVID P','N'),('DIEGO ARMANDO','M','LORENA','N'),('BEA','M','DENIA','T')]  # alternan cada bloque
BLOCK_LENGTHS=(4,5,6)          # semanas por bloque permitidas (5 preferido)
# Grupos de equidad de fines de semana
G_DAY=('CHEMA','BEA','JUANMA','MICHEL','ALFREDO','IRENE','DENIA','SOFIA')
G_ROT=('ESTHER','DIEGO ARMANDO','DAVID P','LORENA')
G_NIGHT=('ARMANDO','JUANMA V','RUBEN S')
# Cobertura objetivo (personas por turno)
TGT_WEEKDAY=3; CAP_WEEKDAY=4            # lunes-jueves (y viernes de mañana/tarde)
FRI_NIGHT_WITH_SAT=3; FRI_NIGHT_NO_SAT=2 # viernes noche
SUN_NIGHT=2; SAT_NIGHT=1
# Mínimo de días trabajados por semana (lunes-domingo) cuando la semana está entera disponible
WEEK_MIN_NORMAL=4        # con 5 como lo habitual (ver CLAUDE.md: las horas anuales no dan para 5 todas las semanas)
WEEK_MIN_HOLIDAY=3       # semanas con festivo no laborable / cierre
# Pesos de la función objetivo
W=dict(
    under_weekday=40, under_fri_night=40, under_weekend=150,   # falta de cobertura (por persona y día)
    over_weekday=8, over_weekend=30,
    single_rest=30,          # descanso de 1 solo día entre dos días trabajados
    single_rest_fri_sun_night=90,  # viernes noche + sábado libre + domingo noche (sin sábado noche)
    six_days=12,             # cada ventana de 6 días seguidos trabajados
    isolated_work=6,         # día de trabajo suelto
    run_start_night=30, run_start_day=4,   # cada inicio de tanda (favorece tandas largas y seguidas)
    fair_weekends=60,        # (máx - mín) de fines de semana dentro de cada grupo
    fair_rot_nights=20,      # (máx - mín) de noches entre los 4 que rotan M/N
    consecutive_weekends=90, # dos fines de semana seguidos
    month_3_weekends=150,     # tercer fin de semana en un mes
    week_min=200,            # por cada día que falte para el mínimo semanal
)
HOURS_TOL=8   # horas anuales entre JA y JA+8
# Fines de semana antes de vacaciones libres (data/libres_antes_vacaciones.json)
LIBRES_ANTES_VACACIONES=True
# Límites por persona y año (CP-SAT)
MAX_RACHAS6=4        # rachas de 6 días seguidos
MAX_DESCANSOS_1=3    # descansos de un solo día
# Noches que no se trabajan: la noche anterior a un festivo/cierre cuenta como el festivo
# (p. ej. la noche del 6 de enero es la del martes 5).
NIGHT_OFF={D(2027,1,5),D(2027,3,24),D(2027,10,11),D(2027,12,23)}
# Días de festivo en que mañana y tarde no trabajan (además de CLOSED; ya lo hace SPECIAL)
DAY_OFF={D(2027,1,6),D(2027,10,12)}
# Asignaciones fijas: fecha -> turno -> personas (el resto de ese turno libra ese día)
FIXED={D(2027,1,2):{'M':['CHEMA'],'T':['MICHEL']}, D(2027,1,3):{'M':['CHEMA'],'T':['MICHEL']}}
