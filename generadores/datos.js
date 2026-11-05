// Lee los CSV/TXT reales del TP y calcula las cifras que usan presentacion.js y guion.js.
const fs = require("fs");
const path = require("path");
const RAIZ = path.join(__dirname, "..");
const leer = (f) => { // utf-8; si falla (CSV viejos en cp1252) cae a latin1
  const b = fs.readFileSync(path.join(RAIZ, f));
  try { return new TextDecoder("utf-8", { fatal: true }).decode(b).replace(/^﻿/, ""); }
  catch (e) { return b.toString("latin1"); }
};

function csv(f) {
  const [cab, ...filas] = leer(f).trim().split(/\r?\n/);
  const cols = cab.split(",");
  return filas.map((l) => {
    const v = l.split(",");
    const o = {};
    cols.forEach((c, i) => {
      const x = v[i];
      o[c] = x !== "" && x !== undefined && !isNaN(Number(x)) ? Number(x) : x;
    });
    return o;
  });
}

// --- formato es-AR: coma decimal, punto de miles
function fn(x, d = 0) {
  const [e, dec] = Math.abs(x).toFixed(d).split(".");
  return (
    (x < 0 ? "-" : "") +
    e.replace(/\B(?=(\d{3})+(?!\d))/g, ".") +
    (dec ? "," + dec : "")
  );
}
// duración legible: ms si < 1 s, si no segundos
function ft(s) {
  if (s < 0.001) return fn(s * 1000, 2) + " ms";
  if (s < 1) return fn(s * 1000, s < 0.01 ? 1 : 0) + " ms";
  if (s < 10) return fn(s, 2) + " s";
  if (s < 100) return fn(s, 1) + " s";
  return fn(s, 0) + " s";
}

// --- hardware
const hwTxt = Object.fromEntries(
  leer("hardware.txt")
    .trim()
    .split(/\r?\n/)
    .map((l) => {
      const i = l.indexOf(":");
      return [l.slice(0, i).trim(), l.slice(i + 1).trim()];
    }),
);
const hw = {
  cpu: hwTxt["CPU"].replace(/\s*\d+-Core Processor/, ""),
  nucleosFisicos: Number((hwTxt["CPU"].match(/(\d+)-Core/) || [])[1]),
  nucleosLogicos: Number(hwTxt["Nucleos logicos (os.cpu_count)"]),
  ram: hwTxt["RAM total"].replace(".", ",").replace(" GB", " GB"),
  so: hwTxt["SO"].startsWith("Windows-11") ? "Windows 11" : hwTxt["SO"],
  python: hwTxt["Python"],
  limite: Number(hwTxt["LIMITE (s)"]),
  fecha: hwTxt["Fecha/hora"].slice(0, 10).split("-").reverse().join("/"),
};

// --- exacto vs WalkSAT
const ex = csv("resultados_exacto.csv");
const unsat = ex
  .filter((r) => r.tipo === "insatisfacible")
  .sort((a, b) => a.n - b.n);
const planted = ex
  .filter((r) => r.tipo.startsWith("planted"))
  .sort((a, b) => a.n - b.n);
const porN = (arr, n) => arr.find((r) => r.n === n);
const terminado = (r) => r.estado_exacto !== "TIMEOUT";

// razon t(n+2)/t(n) en insatisfacibles; solo pares donde ambos terminaron
const razones = [];
for (let i = 1; i < unsat.length; i++) {
  const a = unsat[i - 1],
    b = unsat[i];
  if (terminado(a) && terminado(b))
    razones.push({ n: a.n, r: b.t_exacto_s / a.t_exacto_s });
}
const rGrandes = razones.filter((x) => x.n >= 18); // n chicos: ruido de medicion
const razon = {
  gm: Math.exp(
    rGrandes.reduce((s, x) => s + Math.log(x.r), 0) / rGrandes.length,
  ),
  min: Math.min(...rGrandes.map((x) => x.r)),
  max: Math.max(...rGrandes.map((x) => x.r)),
  desde: rGrandes[0].n,
  hasta: rGrandes[rGrandes.length - 1].n + 2,
  todas: razones,
};
const timeouts = ex.filter((r) => r.estado_exacto === "TIMEOUT");
const ultimoUnsat = unsat.filter(terminado).slice(-1)[0];
const primerTimeout = timeouts.length
  ? Math.min(...timeouts.map((r) => r.n))
  : null;
const tmaxWalkSat = Math.max(...planted.map((r) => r.t_walksat_medio_s));
const walkUnsatT = unsat.map((r) => r.t_walksat_medio_s);
const errUnsat = {
  min: Math.min(...unsat.map((r) => r.error_relativo)),
  max: Math.max(...unsat.map((r) => r.error_relativo)),
};
const todosPlantedDiez = planted.every((r) => r.walksat_exitos_de_10 === 10);
const todosUnsatCero = unsat.every((r) => r.walksat_exitos_de_10 === 0);

// --- escalabilidad
const es = csv("resultados_escala.csv");
const escala = [...new Set(es.map((r) => r.n))]
  .sort((a, b) => a - b)
  .map((n) => {
    const f = es.filter((r) => r.n === n);
    const pl = f.filter((r) => r.tipo === "planted"),
      al = f.filter((r) => r.tipo === "aleatoria");
    const ok = (a) => a.filter((r) => r.estado === "SAT").length;
    return {
      n,
      corridas: f.length,
      ok: ok(f),
      okPlanted: ok(pl),
      okAleat: ok(al),
      nPlanted: pl.length,
      nAleat: al.length,
      tMedio: f.reduce((s, r) => s + r.t_s, 0) / f.length,
      tMax: Math.max(...f.map((r) => r.t_s)),
      pctMin: Math.min(...f.map((r) => r.pct_clausulas)),
      flipsMax: Math.max(...f.map((r) => r.flips)),
    };
  });
const nMax = escala[escala.length - 1];

// --- densidad
const den = csv("resultados_densidad.csv");
const densidad = {
  filas: den,
  satTotal: den.reduce((s, r) => s + r.satisfacibles_segun_exacto, 0),
  walkTotal: den.reduce((s, r) => s + r.walksat_encontro_solucion, 0),
  falsos: den.reduce((s, r) => s + r.walksat_dijo_SAT_en_UNSAT, 0),
  instancias: den.reduce((s, r) => s + r.instancias, 0),
  porRazon: den[0].instancias,
  rMin: den[0].razon_m_n,
  rMax: den[den.length - 1].razon_m_n,
};

// --- historial n=500
const hist = csv("historial_n500.csv");
const historial = {
  pasos: hist.length,
  inicial: hist[0].insatisfechas,
  final: hist[hist.length - 1].insatisfechas,
};

const grafico = (f) => path.join(RAIZ, "graficos", f);
const faltan = [
  "1_tiempo_fuerza_bruta",
  "2_tiempo_walksat",
  "3_comparacion_tiempos",
  "4_exito_walksat",
  "5_asignaciones",
  "6_insatisfechas_por_iteracion",
].filter((g) => !fs.existsSync(grafico(g + ".png")));
if (faltan.length) console.warn("AVISO: faltan gráficos:", faltan.join(", "));

module.exports = {
  RAIZ,
  fn,
  ft,
  hw,
  ex,
  unsat,
  planted,
  porN,
  terminado,
  razon,
  timeouts,
  ultimoUnsat,
  primerTimeout,
  tmaxWalkSat,
  walkUnsatT,
  errUnsat,
  todosPlantedDiez,
  todosUnsatCero,
  escala,
  nMax,
  densidad,
  historial,
  grafico,
};
