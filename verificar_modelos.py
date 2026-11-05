# Control aparte: re-corre WalkSAT (mismos parametros que experimentos.py exacto) sobre las
# instancias planted y re-verifica cada modelo devuelto con contar_insatisfechas.
import glob
from sat import leer_dimacs, walksat, contar_insatisfechas

ok = total = 0
for ruta in sorted(glob.glob("instancias/planted_*.cnf")):
    formula, n = leer_dimacs(ruta)
    for s in range(10):
        r = walksat(formula, n, max_flips=100000, p=0.5, reinicios=1, semilla=s)
        if r["estado"] == "SAT":
            total += 1
            ok += contar_insatisfechas(formula, r["modelo"]) == 0
print(f"modelos SAT devueltos: {total}; verificados con 0 insatisfechas: {ok}")
assert total == ok
