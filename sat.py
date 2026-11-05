# SAT / 3-SAT: generador de instancias, fuerza bruta (exacto) y WalkSAT (heurística)
# Teoría de la Computabilidad 2026 - Trabajo Final
#
# Una fórmula se guarda como lista de cláusulas; cada cláusula es una lista de enteros con signo:
#   3 significa x3, -3 significa NOT x3.
#   Ejemplo: [[1, -2, 3], [-1, 2, 4], [2, -3, -4]]

import random
import time


# ------------------------------------------------------------------
# GENERADOR DE INSTANCIAS
# ------------------------------------------------------------------
def generar_aleatoria(n, m, k=3, semilla=0):
    """k-SAT aleatorio: cada cláusula elige k variables distintas y le pone signo al azar.
    Puede salir satisfacible o no (cerca de m/n = 4,26 se mezclan)."""
    azar = random.Random(semilla)
    formula = []
    for _ in range(m):
        variables = azar.sample(range(1, n + 1), k)
        formula.append([v if azar.random() < 0.5 else -v for v in variables])
    return formula


def generar_planted(n, m, k=3, semilla=0):
    """Satisfacible garantizada: se elige una solución oculta y cada cláusula se
    arregla para que la solución la cumpla."""
    azar = random.Random(semilla)
    solucion = [None] + [azar.random() < 0.5 for _ in range(n)]   # solucion[i] es el valor de x_i
    formula = []
    for _ in range(m):
        variables = azar.sample(range(1, n + 1), k)
        clausula = [v if azar.random() < 0.5 else -v for v in variables]
        cumple = any((lit > 0) == solucion[abs(lit)] for lit in clausula)
        if not cumple:                                   # dar vuelta un literal al azar
            i = azar.randrange(k)
            clausula[i] = -clausula[i]
        formula.append(clausula)
    return formula, solucion


def generar_insatisfacible(n, m, semilla=0):
    """Insatisfacible garantizada: las 8 cláusulas con todas las combinaciones de signos sobre
    3 variables (nadie puede cumplirlas a la vez) más cláusulas aleatorias hasta llegar a m."""
    azar = random.Random(semilla)
    tres = azar.sample(range(1, n + 1), 3)
    formula = []
    for mascara in range(8):
        formula.append([tres[i] if (mascara >> i) & 1 else -tres[i] for i in range(3)])
    formula += generar_aleatoria(n, m - 8, 3, semilla + 7919)
    azar.shuffle(formula)                                # mezclar para que no queden al principio
    return formula


# ------------------------------------------------------------------
# FORMATO DIMACS CNF (para usar benchmarks externos)
# ------------------------------------------------------------------
def escribir_dimacs(ruta, formula, n, comentarios=()):
    with open(ruta, "w") as f:
        for c in comentarios:
            f.write("c " + c + "\n")
        f.write(f"p cnf {n} {len(formula)}\n")
        for cl in formula:
            f.write(" ".join(map(str, cl)) + " 0\n")


def leer_dimacs(ruta):
    formula, n, actual = [], 0, []
    with open(ruta) as f:
        for linea in f:
            linea = linea.strip()
            if not linea or linea[0] == "c":
                continue
            if linea[0] == "p":
                n = int(linea.split()[2])
                continue
            for numero in linea.split():
                numero = int(numero)
                if numero == 0:
                    formula.append(actual)
                    actual = []
                else:
                    actual.append(numero)
    return formula, n


# ------------------------------------------------------------------
# VERIFICACIÓN
# ------------------------------------------------------------------
def contar_insatisfechas(formula, asignacion):
    """asignacion[i] = True/False para la variable i (la posición 0 no se usa)."""
    malas = 0
    for cl in formula:
        if not any((lit > 0) == asignacion[abs(lit)] for lit in cl):
            malas += 1
    return malas


# ------------------------------------------------------------------
# ALGORITMO EXACTO: fuerza bruta -> O(2^n * m * k)
# ------------------------------------------------------------------
def fuerza_bruta(formula, n, limite_s=None):
    """Prueba las 2^n asignaciones. El número 'bits' representa una asignación:
    el bit (i-1) vale 1 si x_i es verdadera. Devuelve SAT, UNSAT o TIMEOUT."""
    inicio = time.perf_counter()
    asignaciones = 0
    clausulas = 0

    for bits in range(2 ** n):
        asignaciones += 1
        cumple_todas = True
        for cl in formula:
            clausulas += 1
            cumple = False
            for lit in cl:
                valor = (bits >> (abs(lit) - 1)) & 1
                if (lit > 0 and valor == 1) or (lit < 0 and valor == 0):
                    cumple = True
                    break
            if not cumple:
                cumple_todas = False
                break                      # esta asignación ya falló
        if cumple_todas:
            modelo = [None] + [bool((bits >> i) & 1) for i in range(n)]
            return dict(estado="SAT", modelo=modelo, asignaciones=asignaciones,
                        clausulas=clausulas, tiempo=time.perf_counter() - inicio)
        # cada tanto miramos el reloj para respetar el límite de tiempo
        if limite_s and (asignaciones & 0x3FFF) == 0 and time.perf_counter() - inicio > limite_s:
            return dict(estado="TIMEOUT", modelo=None, asignaciones=asignaciones,
                        clausulas=clausulas, tiempo=time.perf_counter() - inicio)

    return dict(estado="UNSAT", modelo=None, asignaciones=asignaciones,
                clausulas=clausulas, tiempo=time.perf_counter() - inicio)


# ------------------------------------------------------------------
# HEURÍSTICA: WalkSAT (incompleta)
# ------------------------------------------------------------------
def walksat(formula, n, max_flips=100000, p=0.5, reinicios=1, semilla=0, limite_s=None,
            guardar_historial=False):
    """Parte de una asignación al azar. Mientras haya cláusulas insatisfechas:
       1) elige una insatisfecha al azar
       2) con probabilidad p da vuelta una variable cualquiera de esa cláusula;
          si no, da vuelta la que rompe menos cláusulas hoy satisfechas (si alguna rompe 0, esa).
    Si no encuentra solución devuelve UNKNOWN (NO significa que sea insatisfacible)."""
    azar = random.Random(semilla)
    inicio = time.perf_counter()
    m = len(formula)

    # en qué cláusulas aparece cada literal
    aparece = {}
    for i, cl in enumerate(formula):
        for lit in cl:
            aparece.setdefault(lit, []).append(i)

    mejor_insat = m
    mejor_asig = None
    flips_total = 0
    historial = []

    for intento in range(reinicios):
        asig = [None] + [azar.random() < 0.5 for _ in range(n)]
        # cuántos literales verdaderos tiene cada cláusula
        verdaderos = [sum((lit > 0) == asig[abs(lit)] for lit in cl) for cl in formula]
        insat = [i for i in range(m) if verdaderos[i] == 0]
        pos = {c: j for j, c in enumerate(insat)}        # posición de cada cláusula en la lista

        for _ in range(max_flips):
            if not insat:
                break
            if len(insat) < mejor_insat:
                mejor_insat, mejor_asig = len(insat), asig[:]
            if limite_s and (flips_total & 0x3FF) == 0 and time.perf_counter() - inicio > limite_s:
                break
            c = insat[azar.randrange(len(insat))]
            cl = formula[c]

            # cuántas cláusulas se rompen al dar vuelta cada variable
            rompe = []
            for lit in cl:
                lit_verdadero = lit if asig[abs(lit)] == (lit > 0) else -lit   # el literal que hoy es verdadero
                rompe.append(sum(1 for cc in aparece.get(lit_verdadero, []) if verdaderos[cc] == 1))

            minimo = min(rompe)
            if minimo == 0:
                elegido = cl[rompe.index(0)]
            elif azar.random() < p:
                elegido = cl[azar.randrange(len(cl))]
            else:
                elegido = cl[rompe.index(minimo)]

            # dar vuelta la variable y actualizar contadores
            v = abs(elegido)
            era_verdadero = elegido if asig[v] == (elegido > 0) else -elegido
            ahora_verdadero = -era_verdadero
            asig[v] = not asig[v]
            for cc in aparece.get(era_verdadero, []):
                verdaderos[cc] -= 1
                if verdaderos[cc] == 0:
                    pos[cc] = len(insat)
                    insat.append(cc)
            for cc in aparece.get(ahora_verdadero, []):
                verdaderos[cc] += 1
                if verdaderos[cc] == 1:                  # dejó de estar insatisfecha
                    j = pos.pop(cc)
                    ultimo = insat.pop()
                    if ultimo != cc:
                        insat[j] = ultimo
                        pos[ultimo] = j
            flips_total += 1
            if guardar_historial:
                historial.append(len(insat))

        if not insat:
            return dict(estado="SAT", modelo=asig, insatisfechas=0, satisfechas=m, pct=100.0,
                        flips=flips_total, reinicios_usados=intento + 1,
                        tiempo=time.perf_counter() - inicio, historial=historial)
        if limite_s and time.perf_counter() - inicio > limite_s:
            break

    return dict(estado="UNKNOWN", modelo=mejor_asig, insatisfechas=mejor_insat,
                satisfechas=m - mejor_insat, pct=100.0 * (m - mejor_insat) / m,
                flips=flips_total, reinicios_usados=reinicios,
                tiempo=time.perf_counter() - inicio, historial=historial)


if __name__ == "__main__":
    # Ejemplo del enunciado
    F = [[1, -2, 3], [-1, 2, 4], [2, -3, -4]]
    r = fuerza_bruta(F, 4)
    print("Fuerza bruta:", r["estado"], r["modelo"], "asignaciones:", r["asignaciones"])
    w = walksat(F, 4, semilla=1)
    print("WalkSAT:", w["estado"], w["modelo"], "flips:", w["flips"])
    G = generar_insatisfacible(10, 40, semilla=3)
    print("Insatisfacible controlada:", fuerza_bruta(G, 10)["estado"], "| WalkSAT:", walksat(G, 10, 2000, semilla=1)["estado"])
