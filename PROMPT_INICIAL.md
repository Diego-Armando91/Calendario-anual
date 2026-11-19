# Prompt para pegar en Claude Code

Abre esta carpeta (`cuadrante2027`) en Claude Code y pega esto:

---

Lee `CLAUDE.md` entero antes de empezar. Es el contexto de mi cuadrante 2027 (soy Diego, M4).

Quiero mejorar la propuesta: la v4 me gusta, pero deja semanas casi vacías. Por ejemplo, a mí me quedan 2 semanas sin trabajar en diciembre, a Armando una en enero y a Rubén dos entre enero y febrero. Todas las semanas hay que trabajar: mínimo 3 días si hay festivo no laborable, 4 en el mejor de los casos y 5 lo normal. Mantén todas las demás reglas de la sección 4 de `CLAUDE.md`.

Pasos:
1. Instala lo necesario (`pip install -r requirements.txt`). Si puedes, instala también `ortools` y valora pasar el modelo a CP-SAT.
2. Ejecuta `prepare_data.py` y `roles.py`, y comprueba con `evaluate.py propuesta_v4.csv` que reproduces las métricas de la v4.
3. Parte de `work/propuesta_v5.csv` (ya cumple el mínimo semanal casi del todo, pero tiene demasiados descansos de 1 día) y genera la v6 con `solve.py`. Dale tiempo, al menos 30–60 minutos de LNS. Si hace falta, ajusta los pesos de `config.py`, y prueba la plantilla fija de noche para el ciclo de 5 semanas.
4. La v6 tiene que dar 0 violaciones duras y 0 semanas por debajo del mínimo (o las mínimas posibles, con su explicación). Tiene que ser igual o mejor que la v4 en descansos de 1 día, domingos noche con 2 personas y noches seguidas.
5. Entrégame la imagen comparativa, el Excel de la propuesta y mi calendario (`calendario_persona.py ... "DIEGO ARMANDO" v6`). Escribe un resumen corto en español: qué mejora, qué empeora y qué no se ha podido cumplir, con fechas.
