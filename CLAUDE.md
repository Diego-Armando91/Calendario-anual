# Cuadrante 2027 — Expedición (almacén de Zaragoza)

Contexto para Claude Code. El usuario es **Diego** (Diego Armando, código **M4**). Escríbele siempre en **español**, en tono cercano y directo. Necesita entregables que pueda mirar de un vistazo:
- una **imagen clara** (PNG)
- un **Excel** con el mismo formato que el suyo

Antes de dar algo por bueno, comprueba siempre los números con `scripts/evaluate.py` y mira la imagen.

## 1. Objetivo

Rehacer el cuadrante anual 2027 de las 15 personas del equipo de expedición. Hay que cumplir las reglas de la sección 4, mantener las vacaciones de cada uno y respetar las horas anuales (JA).

Diego ha visto cuatro versiones (v1–v4) y ha ido añadiendo reglas. La v4 le gustó, pero tiene **semanas casi vacías**. Por ejemplo:
- a Diego le quedan 2 semanas sin trabajar en diciembre
- a Armando, una en enero
- a Rubén, dos entre enero y febrero

Esto **no puede ser** (regla 4.12). La v6 (CP-SAT) lo resolvió; la versión actual es la **v11** (`work/propuesta_v11.csv`, hecha desde la v10 editada por Diego).

## 2. Datos (`data/`)

### Excel original

`CALENDARIO2027REVISAR.xlsx` es el cuadrante original. Hoja **`Rotación anual`**:
- **Filas 4–18:** 15 personas. Columna A, nombre; B, código; C, sección.
- **Columnas D..ND:** 365 días de 2027.
  - La cabecera de fechas está guardada como **2026** con formato d-mmm, pero es 2027.
  - La fila 2 tiene las letras L M X J V S D, y el 1 de enero es viernes.
- **Códigos:**
  - M mañana; MS / MD mañana sábado / domingo
  - T tarde; TS / TD tarde sábado / domingo
  - N noche; NV viernes noche; NS sábado noche
  - C central de mañana (solo Juanma C1)
- **Rellenos:**
  - negro (color de tema 1) = **vacaciones**
  - `FF00B0F0` azul = día "azul" (**no obligatorio**; mejor cubrir los descansos)
  - `FFFF0000` rojo = festivo. **Ojo:** en el original hay rojos que no son festivos (11 feb, 15 ene, bloques de Chema, 20-22 dic…). Diego los limpió: los rojos válidos están en `data/rojos.json` (se pintan en rojo, no se trabajan y son la única excepción a la regla de noches sueltas).
- **Resumen por persona:** filas 26–40.
  - Columnas: fines de semana, sábados, horas totales y **JA (jornada anual)** en la columna G.
  - JA: 1664 h en M1, M4, N1 y N5; 1686 h en los de día; 1638 h en N2–N4.
- **Desde / hasta de vacaciones:** filas 71–117.
- **Otras hojas:**
  - `Resumen` explica la rotación tipo de 15 semanas. Las parejas M1/N1, M4/N5 y M3/T4 se intercambian cada 5 semanas.
  - `calendario2027` es una vista diaria.
  - La tabla "Fines de semana asignados" de `Resumen` **no coincide** con `Rotación anual`. Manda `Rotación anual`, porque reproduce las horas.

### Personas

| Código | Persona | Turno |
|---|---|---|
| M1 | Esther | Rota mañana ↔ noche, pareja de David (N1) |
| M2 | Chema | Siempre mañana |
| M3 | Bea | Rota mañana ↔ tarde, pareja de Denia (T4) |
| M4 | Diego Armando | Rota mañana ↔ noche, pareja de Lorena (N5) |
| C1 | Juanma | Mañana / central |
| N1 | David P | Rota mañana ↔ noche, pareja de Esther (M1) |
| N2, N3, N4 | Armando, Juanma V, Rubén S | Siempre noche |
| N5 | Lorena | Rota mañana ↔ noche, pareja de Diego (M4) |
| T1, T2, T3, T5 | Michel, Alfredo, Irene, Sofía | Siempre tarde |
| T4 | Denia | Rota mañana ↔ tarde, pareja de Bea (M3) |

### Fechas especiales

Están en `scripts/config.py`.
- **Cierres** (nadie trabaja): 25 y 26 de marzo; 24, 25 y 31 de diciembre.
- **Noches especiales de 00:00 a 07:00** (7 h, solo personal de noche): 1 y 6 de enero, 27 de marzo y 12 de octubre.
- **Festivos laborables** (recuadro rojo en los calendarios): 29 de enero, 5 de marzo, 23 de abril, 1 de mayo y 15 de agosto.

### Horas

Turno normal, 8 h. **Domingo noche y noches especiales, 7 h.** Horario de Diego: mañana de 6:45 a 15:00 y noche de 23:00 a 07:00.

### Otros ficheros

- `Leyenda_referencia_Diegof.xlsx`: leyenda y estilo que Diego validó para los calendarios individuales.
- `Propuesta_Cuadrante_2027_v3.xlsx` y `Propuesta_Cuadrante_2027_v4.xlsx`: propuestas anteriores.
- `work/propuesta_v3.csv` y `work/propuesta_v4.csv`: las mismas propuestas como matriz 0/1. El turno de cada día sale de `work/roles.json`.

## 3. Cómo está montado (`scripts/`)

| Paso | Script | Qué hace |
|---|---|---|
| 1 | `prepare_data.py` | Lee el Excel y escribe `work/base.json` |
| 2 | `roles.py` | Turno de cada persona por semana (lunes–domingo) en bloques de unas 5 semanas. Equilibra las noches de cada pareja según sus vacaciones. Escribe `work/roles.json` |
| — | `model.py` | Modelo MILP con `scipy.optimize.milp` (HiGHS). `x[i,d]=1` si la persona i trabaja el día d en el turno de su semana. Las reglas duras son restricciones; las blandas, holguras con peso (`config.W`) |
| 3 | `solve.py` | Solución inicial por horizonte rodante, o a partir de un CSV. Después LNS: libera un grupo (noche / mañana / tarde) durante 9 semanas, deja el resto fijo y acepta si mejora |
| 4 | `evaluate.py` | Métricas frente al cuadrante actual: descansos de 1 día, rachas, semanas de 6 días, fines de semana seguidos, tandas de noche, viernes y domingo noche, semanas bajo el mínimo, horas y violaciones duras |
| 5 | `export_excel.py` | Excel **editable** de la propuesta: las hojas de resumen son fórmulas sobre "Propuesta anual" (V = vacaciones) |
| 6 | `export_image.py` | Imagen comparativa con el cuadrante actual |
| 7 | `calendario_persona.py` | Calendario de una persona (PNG + Excel) con la leyenda de Diego |

**Todos los parámetros están en `scripts/config.py`.**

```bash
pip install -r requirements.txt
cd scripts
python prepare_data.py && python roles.py
python solve.py --inicio propuesta_v5.csv --minutos 45 --salida propuesta_v6.csv   # o --inicio rolling
python evaluate.py propuesta_v5.csv
python export_excel.py propuesta_v5.csv v5 && python export_image.py v5
python calendario_persona.py propuesta_v5.csv "DIEGO ARMANDO" v5
```

### Lecciones aprendidas

- **El MILP del año completo de una vez da soluciones flojas.** Lo que funciona:
  - empezar por horizonte rodante o desde una versión anterior
  - aplicar mucho LNS por grupos
  - para cambios grandes en la noche, liberar el grupo de noche el año entero (`--reparar-noche`, unos 7 min)
- **Los pesos son delicados.** Si se suben los inicios de tanda de noche, aparecen muchas tandas de 6. Si se baja el peso del viernes noche, se reduce el patrón viernes noche + sábado libre + domingo noche.
- **Más tiempo de LNS da mejor resultado.** Si se puede, instala **OR-Tools** (`pip install ortools`) y pasa el modelo a CP-SAT, que es mucho mejor para cuadrantes. Las restricciones se trasladan casi tal cual.
- **Idea que Diego pidió y que aún no está implementada:** una plantilla de noche fija para el ciclo de 5 semanas (una "rueda" de las 5 personas de noche) que se repita y que solo se rompa por vacaciones.

## 4. Reglas de Diego (todas, por orden de importancia)

**Duras**
1. Se respetan las vacaciones (celdas negras) y los cierres.
   - **1c (v8):** la noche anterior a un festivo o cierre no se trabaja (5 ene, 24 mar, 11 oct, 23 dic: `config.NIGHT_OFF`); el 6 ene y el 12 oct no hay mañana ni tarde (`config.DAY_OFF`). El 2-3 ene solo trabajan Chema (mañana) y Michel (tarde) (`config.FIXED`).
   - **1d (v9):** Juanma (C1) no trabaja los fines de semana que tiene a su hijo (alternos, `data/no_disponible.json`). Se tratan como libres: no rebajan su mínimo semanal y recupera las horas entre semana.
   - **1e (v9):** de noche, nada de noches sueltas: tandas de al menos 2 noches seguidas (duro) y mejor 3 o más (se penalizan las de 2). Solo se puede romper junto a un día rojo (festivo no laborable o cierre) (`config.MIN_NIGHT_RUN`).
   - **1b (v7):** el fin de semana anterior a unas vacaciones es libre. Está en `data/libres_antes_vacaciones.json` (los que pintó Diego en su v6 más los de vacaciones que empiezan de lunes a viernes). **A Diego no se le aplica la regla general**: solo los que pintó él (22-23 de mayo y 19-20 de junio).
2. **Cada semana de lunes a domingo es entera de un mismo turno.**
3. Las parejas rotan en bloques de 5 semanas (alguno de 4 o 6 para cuadrar vacaciones).
4. Máximo 6 días seguidos.
5. No se puede pasar de noche a mañana o tarde, ni de tarde a mañana, al día siguiente.
6. Las horas anuales quedan entre JA y JA + 8 h.
7. **Sábado noche solo del 27 de febrero al 29 de mayo y del 17 de julio al 4 de diciembre** (ambos incluidos).

**Blandas**
8. **Domingo noche: siempre o casi siempre 2 personas.**
9. Viernes noche: 3 personas la mayoría de las veces. Si hace falta, 2 (sobre todo cuando no hay sábado noche).
   - **v9:** si alguien del turno de noche tiene vacaciones esa semana, basta con 2 el viernes noche y 1 el domingo noche (`config.FRI_NIGHT_VAC`, `config.SUN_NIGHT_VAC`).
10. Fines de semana de mañana y tarde con el mismo formato que el original: 1 persona hasta el 3 de septiembre y 2 desde el 4 de septiembre. Quien trabaja el sábado trabaja también el domingo.
11. **Evitar los descansos de un solo día.** El peor caso es viernes noche, sábado libre y domingo noche.
12. **NUEVA: todas las semanas se trabaja.** Mínimo 3 días si hay festivo no laborable, 4 en el mejor de los casos y 5 lo normal. Nada de semanas vacías.
    - JA = 1664 h son unos 208 días. Con unas 47 semanas sin vacaciones, la media es de unos 4,4 días por semana.
    - 5 días todas las semanas no cabe. Lo que sale es unas 25 semanas de 5, unas 17 de 4 y las de festivo de 3.
    - En `config` el mínimo es 4 (3 con festivo, proporcional si hay vacaciones). Explícaselo a Diego si pregunta por qué no son todas de 5.
13. Noches lo más seguidas posible: tandas continuas de 4 o 5 noches, que pueden cruzar de una semana a otra (por ejemplo, de jueves a martes, o de domingo a jueves). Que no se partan todas las semanas.
14. **v11: como mucho 2 fines de semana seguidos** (duro para día y rotativos; penalizado para noche fija, `config.MAX_FINDES_SEGUIDOS`). Al reoptimizar se mantiene ±1 el reparto de sábados, viernes, sábados y domingos noche de la versión de Diego (`--mantener-equidad`); Diego y Armando pueden hacer más sábados noche y Rubén S menos domingos noche (`config.EQUIDAD_EXC`).
    Fines de semana y sábados noche alternos, como máximo 2 al mes (3 solo si las vacaciones obligan). Se permiten **dos seguidos de vez en cuando, pero no a menudo**.
15. Alguna semana de 6 días está permitida, pero no puede ser lo normal.
    - **v7: como mucho 4 rachas de 6 días seguidos por persona y año, crucen o no de semana** (`config.MAX_RACHAS6`).
    - **v7: como mucho 3 descansos de un solo día por persona y año** (`config.MAX_DESCANSOS_1`).
16. Los fines de semana se reparten por igual dentro de cada grupo. Los de noche fija hacen menos.
17. Las personas que rotan quedan lo más iguales posible entre sí (noches y fines de semana).
18. Los días azules no son obligatorios.

## 5. Estado de las versiones

| | Actual | v3 | v4 | v5 (primer intento) |
|---|---|---|---|---|
| Descansos de 1 día | 232 | 49 | 57 | **107** (empeora: hay que bajarlo) |
| Semanas de 6 días | 142 | 48 | 57 | 34 |
| Fines de semana seguidos | 34 | 26 | 33 | 26 |
| Semanas mezcladas | 23 | 0 | 0 | 0 |
| Semanas bajo el mínimo (regla 12) | — | — | **121** | 1 (Juanma V, 31 ago – 5 sep) |
| Viernes noche con 3 o más | 11 | 43 | 27 | 26 |
| Domingos noche con 2 | 39 | 50 | 51 | 50 |

La v5 (`work/propuesta_v5.csv`) cumple casi del todo el mínimo semanal, pero duplica los descansos de 1 día.

Siguiente paso:
- partir de la v5 y dar mucho más LNS, o usar CP-SAT
- probar la plantilla fija de noche
- quitar la semana de 1 día de Juanma V
- revisar los viernes noche con 0 o 1 persona (2, 23 y 30 de julio y 3 de diciembre)

Pendiente en la v4:
- 16 casos de viernes noche + sábado libre + domingo noche, en las semanas sin sábado noche
- en agosto, a Diego le tocan 3 fines de semana seguidos

## 6. Qué entregar a Diego

- **Comparativa:** `salidas/Propuesta_vX_vs_Actual_2027.png`.
- **Excel de la propuesta:** `salidas/Propuesta_Cuadrante_2027_vX.xlsx`, con las hojas:
  - Propuesta anual
  - Rotaciones
  - Resumen personas
  - Cobertura
  - Viernes noche
  - Criterios
- **Calendarios individuales**, cuando los pida: `calendario_persona.py`. Leyenda:
  - verde mañana, con borde amarillo si es sábado o domingo
  - lila tarde
  - naranja noche (NV es viernes noche); naranja claro sábado noche
  - gris oscuro vacaciones
  - rojo festivo no laborable o cierre
  - recuadro rojo festivo laborable
- **Resumen de 6 a 10 líneas:** qué mejora, qué empeora y qué ha quedado sin cumplir (con fechas concretas), y una pregunta sobre qué ajustar.
- **Antes de entregar**, `evaluate.py` debe dar **0 violaciones duras**, y hay que revisar la imagen.
