# Calendario anual 2027 · Cuadrante de expedición

Herramientas para generar y evaluar el cuadrante anual 2027 del equipo de expedición (15 personas).
Todo el contexto y las reglas están en [`CLAUDE.md`](CLAUDE.md).

## Estructura

- `data/`: Excel original y propuestas anteriores.
- `scripts/`: modelo, solver y exportadores. Los parámetros están en `scripts/config.py`.
- `work/`: datos preparados (`base.json`, `roles.json`) y propuestas en CSV (matriz 0/1).
- `salidas/`: Excel, imágenes comparativas y calendarios individuales.

## Uso

```bash
pip install -r requirements.txt
cd scripts
python prepare_data.py && python roles.py
python solve_cpsat.py --inicio propuesta_v5.csv --minutos 30 --minimo-duro --salida propuesta_v6.csv
python evaluate.py propuesta_v6.csv
python export_excel.py propuesta_v6.csv v6 && python export_image.py v6
python calendario_persona.py propuesta_v6.csv "DIEGO ARMANDO" v6
```

`solve_cpsat.py` (OR-Tools CP-SAT) sustituye al MILP por LNS de `solve.py`: resuelve el año entero de una vez y da
mejores resultados en menos tiempo. `solve.py` se mantiene como alternativa.
