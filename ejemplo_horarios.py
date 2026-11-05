# Ejemplo de juguete HIPOTETICO del informe (seccion 10.3): 3 comisiones x 3 franjas.
# Variable x(c,f) = 3*(c-1) + f. C1 y C2 comparten docente, C2 y C3 comparten alumnos,
# y el docente de C3 no puede en la franja 1. Uso: python ejemplo_horarios.py
from sat import fuerza_bruta, walksat, contar_insatisfechas

x = lambda c, f: 3 * (c - 1) + f
F = []
F += [[x(c, 1), x(c, 2), x(c, 3)] for c in (1, 2, 3)]                          # al menos una franja
F += [[-x(c, f), -x(c, g)] for c in (1, 2, 3) for f, g in ((1, 2), (1, 3), (2, 3))]  # a lo sumo una
F += [[-x(a, f), -x(b, f)] for a, b in ((1, 2), (2, 3)) for f in (1, 2, 3)]    # conflictos
F += [[-x(3, 1)]]                                                               # disponibilidad
horario = lambda mod: {c: next(f for f in (1, 2, 3) if mod[x(c, f)]) for c in (1, 2, 3)}

e = fuerza_bruta(F, 9)
w = walksat(F, 9, semilla=1)
print(f"{len(F)} clausulas, 9 variables")
print("Exacto:  ", e["estado"], "asignaciones:", e["asignaciones"], "franjas:", horario(e["modelo"]))
print("WalkSAT: ", w["estado"], "pasos:", w["flips"], "franjas:", horario(w["modelo"]))
assert contar_insatisfechas(F, e["modelo"]) == 0 and contar_insatisfechas(F, w["modelo"]) == 0
