# SAT (Problema de Satisfacción Booleana): algoritmo exacto vs. heurística WalkSAT

**Problema:** SAT, el Problema de Satisfacción Booleana: dada una fórmula booleana, decidir si existe una asignación de verdadero/falso que la haga verdadera. Es NP-completo (teorema de Cook-Levin, 1971). Las instancias de los experimentos se restringen a 3-SAT (tres literales por cláusula).

**Algoritmos que se comparan sobre las mismas fórmulas:**

- Exacto (línea de base): fuerza bruta, prueba las 2^n asignaciones. Es completo, pero crece exponencialmente.
- Heurística: WalkSAT, una búsqueda local estocástica. Es rápida con n grande, pero incompleta: si no encuentra solución devuelve UNKNOWN, que no prueba insatisfacibilidad.

El tema del trabajo es SAT; WalkSAT es la heurística elegida para resolverlo.

Solo se necesita Python 3.9+ y matplotlib (para los gráficos). El resto es biblioteca estándar.

## Paso a paso desde cero

1. Instalar dependencias (opcional: antes crear un venv con `python -m venv .venv` y activarlo):

        pip install -r requirements.txt

2. Experimento exacto (fuerza bruta vs. WalkSAT sobre las MISMAS fórmulas). `LIMITE` es el tiempo máximo (s) por ejecución de fuerza bruta:

   Linux / macOS / Git Bash:

        LIMITE=3600 python experimentos.py exacto 10 12 14 16 18 20 22 24 26 28 30 32

   Windows PowerShell:

        $env:LIMITE=3600; python experimentos.py exacto 10 12 14 16 18 20 22 24 26 28 30 32

   Se puede correr por tandas: los resultados se acumulan en `resultados_exacto.csv`. Tarda mucho para n grandes.

3. WalkSAT con n grande (100 a 10.000):

        python experimentos.py escala

4. Gráficos y tabla:

        python graficos.py

   Genera `graficos/*.png` y `tabla_comparacion.md` (tabla completa para pegar en el informe).

5. Control aparte de los modelos: `python verificar_modelos.py` re-corre WalkSAT (10 semillas) sobre las instancias planted de `instancias/` y confirma con `contar_insatisfechas` que cada modelo devuelto deja 0 cláusulas insatisfechas (120 de 120). WalkSAT por sí solo declara SAT cuando no quedan cláusulas insatisfechas según sus contadores internos; no llama a `contar_insatisfechas`.

6. Documentos (opcional): `SAT_Informe.docx`, `SAT_Presentacion.pptx`, `SAT_Guion_Exposicion.docx` y `GUIA_DE_ESTUDIO.docx` se generan con Node.js leyendo los CSV y `hardware.txt`, así que las cifras no se copian a mano. Una vez, `cd generadores && npm install`; después, desde la raíz: `node generadores/informe.js`, `node generadores/presentacion.js`, `node generadores/guion.js` y `node generadores/estudio.js`.

Opcional: `python ejemplo_horarios.py` (ejemplo de juguete HIPOTÉTICO de horarios del informe, sección 10.3), `python sat.py` (ejemplo del enunciado) y `python experimentos.py densidad` (variación de cláusulas por variable, n=16).

## Salidas

- `resultados_exacto.csv`, `resultados_escala.csv`, `resultados_densidad.csv`, `historial_n500.csv`.
- `resultados_exacto.csv` incluye `error_relativo` = (cláusulas insatisfechas de la mejor asignación de WalkSAT) / m, con 0.0 = óptimo alcanzado; y `optimo_alcanzado` (sí / no / n.a.: n.a. si el exacto dio TIMEOUT o la fórmula es insatisfacible por construcción, donde el óptimo Max-SAT no se conoce).
- `hardware.txt`: se escribe al iniciar `exacto` con CPU, núcleos, RAM, SO, versión de Python, LIMITE y fecha/hora. Citarlo en el informe junto con los tiempos.
- `instancias/`: fórmulas en DIMACS CNF. Las semillas son fijas, se pueden repetir los resultados.

## Condiciones controladas

Para que los tiempos sean comparables: mismo equipo para todas las corridas, enchufado (plan de energía de alto rendimiento), sin otros programas abiertos, mismo `LIMITE` y misma versión de Python. Si se corre por tandas, hacerlo en la misma máquina y verificar que `hardware.txt` corresponda.

## Autoría y uso de IA

El código (`sat.py`, `experimentos.py`, `graficos.py`, `verificar_modelos.py`, generadores de documentos) y los documentos del trabajo fueron desarrollados con asistencia de una herramienta de inteligencia artificial (Claude, de Anthropic), y revisados, ejecutados y comprendidos por los autores, que son responsables de su contenido.
