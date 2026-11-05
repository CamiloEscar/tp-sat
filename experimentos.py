# Experimentos del trabajo final: SAT (3-SAT)
#
#   python experimentos.py exacto 10 12 14 16 18 20 22 24   -> fuerza bruta vs WalkSAT sobre las MISMAS fórmulas
#   python experimentos.py escala                            -> WalkSAT con n grande
#   python experimentos.py densidad                          -> qué pasa al variar cláusulas/variables (n=16)
#
# Las fórmulas se guardan en la carpeta instancias/ (formato DIMACS) para poder reutilizarlas.
# Opcional: LIMITE=3600 python experimentos.py exacto ...  (límite de tiempo de la fuerza bruta en segundos)

import csv
import ctypes
import datetime
import os
import platform
import subprocess
import sys

from sat import (generar_planted, generar_insatisfacible, generar_aleatoria, escribir_dimacs,
                 fuerza_bruta, walksat)

LIMITE = float(os.environ.get("LIMITE", 300))
RAZON = 4.26          # cláusulas por variable: la zona más difícil del 3-SAT aleatorio


def guardar(nombre, filas):
    campos = []
    for f in filas:
        for k in f:
            if k not in campos:
                campos.append(k)
    with open(nombre, "w", newline="", encoding="utf-8") as a:
        w = csv.DictWriter(a, fieldnames=campos)
        w.writeheader()
        w.writerows(filas)


def correr_walksat_varias_veces(formula, n, veces=10, max_flips=100000):
    exitos, tiempos, flips, mejor_pct = 0, [], [], 0
    for s in range(veces):
        r = walksat(formula, n, max_flips=max_flips, p=0.5, reinicios=1, semilla=s)
        exitos += r["estado"] == "SAT"
        tiempos.append(r["tiempo"])
        flips.append(r["flips"])
        mejor_pct = max(mejor_pct, r["pct"])
    return exitos, sum(tiempos) / veces, sum(flips) / veces, mejor_pct


def _cpu():
    nombre = ""
    if os.name == "nt":
        for cmd in (["wmic", "cpu", "get", "name"],
                    ["powershell", "-NoProfile", "-Command", "(Get-CimInstance Win32_Processor).Name"]):
            try:
                out = subprocess.run(cmd, capture_output=True, text=True, timeout=15).stdout.split("\n")
                nombre = " ".join(l.strip() for l in out if l.strip() and l.strip().lower() != "name")
            except Exception:
                pass
            if nombre:
                break
    return nombre or platform.processor() or "desconocido"


def _ram_gb():
    try:
        if os.name == "nt":
            class MEM(ctypes.Structure):
                _fields_ = [("l", ctypes.c_ulong), ("u", ctypes.c_ulong), ("total", ctypes.c_ulonglong)] +                            [(f"x{i}", ctypes.c_ulonglong) for i in range(6)]
            mem = MEM()
            mem.l = ctypes.sizeof(MEM)
            ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(mem))
            return mem.total / 2**30
        for linea in open("/proc/meminfo", encoding="utf-8"):
            if linea.startswith("MemTotal"):
                return int(linea.split()[1]) / 2**20
    except Exception:
        pass
    return None


def escribir_hardware():
    ram = _ram_gb()
    lineas = [f"CPU: {_cpu()}", f"Procesador (platform): {platform.processor() or 'n/d'}",
              f"Nucleos logicos (os.cpu_count): {os.cpu_count()}",
              f"RAM total: {f'{ram:.1f} GB' if ram else 'n/d'}", f"SO: {platform.platform()}",
              f"Python: {sys.version.split()[0]}", f"LIMITE (s): {LIMITE:g}",
              f"Fecha/hora: {datetime.datetime.now().isoformat(timespec='seconds')}"]
    with open("hardware.txt", "w", encoding="utf-8") as a:
        a.write("\n".join(lineas) + "\n")


def exacto(tamanios):
    os.makedirs("instancias", exist_ok=True)
    escribir_hardware()
    archivo = "resultados_exacto.csv"
    filas = []
    if os.path.exists(archivo):                        # se pueden correr varios tamaños en tandas
        try:
            t = open(archivo, newline="", encoding="utf-8").read()
        except UnicodeDecodeError:                      # CSV viejos en cp1252
            t = open(archivo, newline="", encoding="cp1252").read()
        filas = list(csv.DictReader(t.splitlines()))
    for n in tamanios:
        m = round(RAZON * n)
        planted, _ = generar_planted(n, m, 3, semilla=n)
        insat = generar_insatisfacible(n, m, semilla=n)
        for tipo, formula in (("planted (SAT)", planted), ("insatisfacible", insat)):
            ruta = f"instancias/{'planted' if 'planted' in tipo else 'insat'}_n{n}_m{m}.cnf"
            escribir_dimacs(ruta, formula, n, [f"tipo {tipo}", f"n={n} m={m} semilla={n}"])
            e = fuerza_bruta(formula, n, LIMITE)
            exitos, t_w, fl, pct = correr_walksat_varias_veces(formula, n)
            fila = dict(tipo=tipo, n=n, m=m, estado_exacto=e["estado"], t_exacto_s=round(e["tiempo"], 4),
                        asignaciones=e["asignaciones"], clausulas_evaluadas=e["clausulas"],
                        walksat_exitos_de_10=exitos, t_walksat_medio_s=round(t_w, 5),
                        flips_medio=round(fl, 1), mejor_pct_clausulas=round(pct, 2),
                        error_relativo=round((100 - pct) / 100, 5),
                        # n.a.: óptimo Max-SAT desconocido (TIMEOUT, o UNSAT por construcción)
                        optimo_alcanzado=("n.a." if e["estado"] == "TIMEOUT" or tipo == "insatisfacible"
                                          else "sí" if pct >= 100 else "no"))
            filas = [f for f in filas if not (f["tipo"] == tipo and int(f["n"]) == n)] + [fila]
            print(fila, flush=True)
            guardar(archivo, filas)


def escala():
    filas = []
    for n in (100, 500, 1000, 5000, 10000):
        for tipo, razon in (("planted", 4.26), ("aleatoria", 4.0)):
            m = round(razon * n)
            for semilla in (1, 2, 3):
                if tipo == "planted":
                    formula, _ = generar_planted(n, m, 3, semilla)
                else:
                    formula = generar_aleatoria(n, m, 3, semilla)
                # la semilla de WalkSAT es distinta a la del generador (si no, arrancaría en la solución oculta)
                r = walksat(formula, n, max_flips=5000000, p=0.5, reinicios=1, semilla=semilla + 1000, limite_s=120)
                fila = dict(tipo=tipo, razon_m_n=razon, n=n, m=m, semilla=semilla, estado=r["estado"],
                            flips=r["flips"], t_s=round(r["tiempo"], 3), pct_clausulas=round(r["pct"], 3))
                filas.append(fila)
                print(fila, flush=True)
                guardar("resultados_escala.csv", filas)
    # evolución de las cláusulas insatisfechas (una corrida, aleatoria n=500, razón 4.0)
    formula = generar_aleatoria(500, 2000, 3, 1)
    r = walksat(formula, 500, max_flips=200000, p=0.5, semilla=1001, guardar_historial=True)
    with open("historial_n500.csv", "w", newline="", encoding="utf-8") as a:
        a.write("iteracion,insatisfechas\n")
        for i, v in enumerate(r["historial"]):
            a.write(f"{i + 1},{v}\n")
    print("historial n=500:", r["estado"], r["flips"], "flips")


def densidad():
    n, filas = 16, []
    for razon in (2.0, 3.0, 4.0, 4.26, 5.0, 6.0):
        m = round(razon * n)
        sat, walk_encontro, walk_mal = 0, 0, 0
        for semilla in range(30):
            formula = generar_aleatoria(n, m, 3, semilla)
            e = fuerza_bruta(formula, n)
            w = walksat(formula, n, max_flips=20000, p=0.5, reinicios=1, semilla=semilla + 1000)
            if e["estado"] == "SAT":
                sat += 1
                walk_encontro += w["estado"] == "SAT"
            elif w["estado"] == "SAT":
                walk_mal += 1                         # no debería pasar nunca
        fila = dict(razon_m_n=razon, m=m, instancias=30, satisfacibles_segun_exacto=sat,
                    walksat_encontro_solucion=walk_encontro, walksat_dijo_SAT_en_UNSAT=walk_mal)
        filas.append(fila)
        print(fila, flush=True)
        guardar("resultados_densidad.csv", filas)


if __name__ == "__main__":
    if sys.argv[1] == "exacto":
        exacto([int(x) for x in sys.argv[2:]])
    elif sys.argv[1] == "escala":
        escala()
    else:
        densidad()
