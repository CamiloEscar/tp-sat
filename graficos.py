# Genera los gráficos a partir de los CSV de resultados reales
import csv
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

AZUL, NARANJA, VERDE, GRIS = "#1F4E79", "#C55A11", "#2E7D32", "#777777"
plt.rcParams.update({"font.size": 12, "axes.grid": True, "grid.alpha": 0.3, "figure.figsize": (7.5, 4.6)})

def leer_csv(ruta):   # utf-8; si falla (CSV viejos en cp1252) cae a cp1252
    try:
        t = open(ruta, newline="", encoding="utf-8").read()
    except UnicodeDecodeError:
        t = open(ruta, newline="", encoding="cp1252").read()
    return list(csv.DictReader(t.splitlines()))

ex = leer_csv("resultados_exacto.csv")
es = leer_csv("resultados_escala.csv")
os.makedirs("graficos", exist_ok=True)

# tabla completa para el informe (acepta CSV viejos sin error_relativo)
def _err(r):
    return float(r["error_relativo"]) if r.get("error_relativo") not in (None, "") else (100 - float(r["mejor_pct_clausulas"])) / 100

with open("tabla_comparacion.md", "w", encoding="utf-8") as a:
    a.write("| n | tipo | estado exacto | t exacto (s) | éxitos WalkSAT/10 | t WalkSAT (s) | error_relativo |\n|---|---|---|---|---|---|---|\n")
    for r in sorted(ex, key=lambda r: (int(r["n"]), r["tipo"])):
        a.write(f"| {r['n']} | {r['tipo']} | {r['estado_exacto']} | {r['t_exacto_s']} | {r['walksat_exitos_de_10']} | {r['t_walksat_medio_s']} | {_err(r):.4f} |\n")
fin = [r for r in ex if r["estado_exacto"] != "TIMEOUT"]     # TIMEOUT no es un tiempo medido
unsat = [(int(r["n"]), float(r["t_exacto_s"]), int(r["asignaciones"])) for r in fin if r["tipo"] == "insatisfacible"]
plant = [(int(r["n"]), float(r["t_exacto_s"]), int(r["asignaciones"]), float(r["t_walksat_medio_s"]), int(r["walksat_exitos_de_10"])) for r in fin if r["tipo"].startswith("planted")]
walk = [(int(r["n"]), float(r["t_walksat_medio_s"])) for r in ex if r["tipo"].startswith("planted")]   # WalkSAT sí terminó siempre
tout = [(int(r["n"]), float(r["t_exacto_s"]), int(r["asignaciones"])) for r in ex if r["estado_exacto"] == "TIMEOUT"]

def marcar_timeout(j):      # j=1: y = tiempo (== LIMITE), j=2: y = asignaciones evaluadas hasta cortar
    if tout:
        plt.plot([t[0] for t in tout], [t[j] for t in tout], "x", color="red", ms=10, mew=2, label="TIMEOUT (no terminó)")

def guardar(nombre):
    plt.tight_layout(); plt.savefig(f"graficos/{nombre}.png", dpi=160); plt.close()

# 1) n vs tiempo fuerza bruta
plt.plot([a for a, b, c in unsat], [b for a, b, c in unsat], "o-", color=AZUL, label="Insatisfacible (recorre las 2ⁿ)")
plt.plot([a[0] for a in plant], [a[1] for a in plant], "s--", color=NARANJA, label="Satisfacible (se detiene al encontrar)")
plt.yscale("log"); plt.xlabel("Cantidad de variables n"); plt.ylabel("Tiempo (s, escala log)")
marcar_timeout(1)
plt.title("Fuerza bruta: tiempo según n"); plt.legend(); guardar("1_tiempo_fuerza_bruta")

# 2) n vs tiempo WalkSAT (escala)
for tipo, color in (("planted", NARANJA), ("aleatoria", AZUL)):
    ns = sorted({int(r["n"]) for r in es if r["tipo"] == tipo})
    med = [sum(float(r["t_s"]) for r in es if r["tipo"] == tipo and int(r["n"]) == n) / 3 for n in ns]
    plt.plot(ns, med, "o-", color=color, label="planted (m/n=4,26)" if tipo == "planted" else "aleatoria (m/n=4,0)")
plt.xscale("log"); plt.yscale("log"); plt.xlabel("Cantidad de variables n (escala log)"); plt.ylabel("Tiempo medio (s, escala log)")
plt.title("WalkSAT: tiempo según n (3 instancias por punto)"); plt.legend(); guardar("2_tiempo_walksat")

# 3) comparación en las mismas fórmulas satisfacibles + insatisfacibles
plt.plot([a[0] for a in plant], [a[1] for a in plant], "s--", color=NARANJA, label="Fuerza bruta (satisfacible)")
plt.plot([a for a, b, c in unsat], [b for a, b, c in unsat], "o-", color=AZUL, label="Fuerza bruta (insatisfacible)")
plt.plot([a[0] for a in walk], [a[1] for a in walk], "^-", color=VERDE, label="WalkSAT (satisfacible)")
plt.yscale("log"); plt.xlabel("Cantidad de variables n"); plt.ylabel("Tiempo (s, escala log)")
marcar_timeout(1)
plt.title("Exacto vs. WalkSAT sobre las mismas fórmulas"); plt.legend(); guardar("3_comparacion_tiempos")

# 4) n vs % de éxito de WalkSAT
for tipo, color, marca in (("planted", NARANJA, "s"), ("aleatoria", AZUL, "o")):
    ns = sorted({int(r["n"]) for r in es if r["tipo"] == tipo})
    pct = [100 * sum(r["estado"] == "SAT" for r in es if r["tipo"] == tipo and int(r["n"]) == n) / 3 for n in ns]
    plt.plot(ns, pct, marca + "-", color=color, label="planted (m/n=4,26)" if tipo == "planted" else "aleatoria (m/n=4,0)")
plt.xscale("log"); plt.ylim(-5, 105); plt.xlabel("Cantidad de variables n (escala log)"); plt.ylabel("Corridas que encontraron solución (%)")
plt.title("WalkSAT: éxito con 5.000.000 de flips como máximo"); plt.legend(); guardar("4_exito_walksat")

# 5) n vs asignaciones evaluadas
plt.plot([a for a, b, c in unsat], [c for a, b, c in unsat], "o-", color=AZUL, label="Insatisfacible (siempre 2ⁿ)")
plt.plot([a[0] for a in plant], [a[2] for a in plant], "s--", color=NARANJA, label="Satisfacible")
plt.yscale("log"); plt.xlabel("Cantidad de variables n"); plt.ylabel("Asignaciones evaluadas (escala log)")
marcar_timeout(2)
plt.title("Fuerza bruta: asignaciones evaluadas"); plt.legend(); guardar("5_asignaciones")

# 6) cláusulas insatisfechas por iteración
if not os.path.exists("historial_n500.csv"):
    print("AVISO: falta historial_n500.csv (correr 'python experimentos.py escala'); se saltea el gráfico 6")
else:
    h = leer_csv("historial_n500.csv")
    plt.plot([int(r["iteracion"]) for r in h], [int(r["insatisfechas"]) for r in h], color=AZUL)
    plt.xlabel("Iteración (flip)"); plt.ylabel("Cláusulas insatisfechas")
    plt.title("WalkSAT: n=500, m=2000 (aleatoria): cláusulas insatisfechas"); guardar("6_insatisfechas_por_iteracion")
print("ok")
