// Genera SAT_Presentacion.pptx con los datos reales de los CSV. Uso: node generadores/presentacion.js
const path = require("path");
const PptxGenJS = require("pptxgenjs");
const D = require("./datos");
const { fn, ft } = D;

const OSCURO = "16222E",
  CREMA = "F5F2EA",
  NARANJA = "D9731A",
  TEAL = "2B7A8C",
  BORDE = "DDD8CC",
  CODIGO = "E8E4D8";
const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_16x9"; // 10 x 5,625 in
pptx.title = "SAT (3-SAT): algoritmo exacto frente a WalkSAT";

// --- estimador de desborde (sin render): ancho medio de caracter como fraccion del cuerpo
const ANCHO = { Arial: 0.52, Georgia: 0.56, "Courier New": 0.6 };
const avisos = [];
let slideActual = 0;
function estimar(txt, w, fs, face, bullet) {
  const util = w - (bullet ? 0.25 : 0);
  const porLinea = Math.max(1, Math.floor(util / ((ANCHO[face] * fs) / 72)));
  return String(txt)
    .split("\n")
    .reduce((s, p) => s + Math.max(1, Math.ceil(p.length / porLinea)), 0);
}

function nueva(bg = CREMA) {
  slideActual++;
  const s = pptx.addSlide();
  s.background = { color: bg };
  return s;
}
// texto simple o lista de parrafos [{t, bold, color, bullet}]
function tx(s, contenido, x, y, w, h, o = {}) {
  const fs = o.fs || 18,
    face = o.face || "Arial";
  const parr = Array.isArray(contenido) ? contenido : [{ t: contenido }];
  let lineas = 0;
  const runs = parr.map((p, i) => {
    const b = p.bullet ?? o.bullet;
    lineas += estimar(p.t, w, fs, face, b) + (p.gap ? p.gap : 0);
    return {
      text: p.t,
      options: {
        bold: p.bold ?? o.bold,
        color: p.color || o.color || OSCURO,
        breakLine: i < parr.length - 1,
        bullet: b ? { indent: 18 } : undefined,
        paraSpaceAfter: p.gap ? p.gap * fs * 1.2 : 0,
      },
    };
  });
  const alto = (lineas * fs * 1.2) / 72;
  if (alto > h + 0.01)
    avisos.push(
      `diap ${slideActual}: texto "${parr[0].t.slice(0, 35)}" ~${alto.toFixed(2)}in > caja ${h}in`,
    );
  s.addText(runs, {
    x,
    y,
    w,
    h,
    fontFace: face,
    fontSize: fs,
    margin: 0,
    valign: o.valign || "top",
    align: o.align || "left",
    italic: o.italic,
    fit: "none",
  });
}
function tarjeta(s, x, y, w, h, fill = "FFFFFF") {
  s.addShape(pptx.ShapeType.rect, {
    x,
    y,
    w,
    h,
    fill: { color: fill },
    line: { color: BORDE, width: 1 },
  });
}
function titulo(s, t) {
  if (estimar(t, 9, 28, "Georgia") > 1)
    avisos.push(`diap ${slideActual}: titulo largo "${t}"`);
  s.addText(t, {
    x: 0.5,
    y: 0.3,
    w: 9,
    h: 0.7,
    fontFace: "Georgia",
    fontSize: 28,
    color: OSCURO,
    margin: 0,
    valign: "middle",
    fit: "none",
  });
  s.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 1.02,
    w: 0.9,
    h: 0.045,
    fill: { color: NARANJA },
    line: { color: NARANJA, width: 0 },
  });
}
function imagen(s, archivo, x, y, w) {
  s.addImage({ path: D.grafico(archivo), x, y, w, h: (w * 736) / 1200 });
  return y + (w * 736) / 1200;
}
// tarjeta con encabezado naranja + cuerpo
function bloque(s, x, y, w, h, cab, cuerpo, o = {}) {
  tarjeta(s, x, y, w, h);
  tx(s, cab, x + 0.15, y + 0.12, w - 0.3, o.hCab || 0.4, {
    bold: true,
    color: o.colorCab || NARANJA,
  });
  tx(
    s,
    cuerpo,
    x + 0.15,
    y + 0.12 + (o.hCab || 0.4) + 0.05,
    w - 0.3,
    h - (o.hCab || 0.4) - 0.3,
    { bullet: o.bullet },
  );
}

// ===== datos derivados =====
const { hw, razon, unsat, planted, porN } = D;
const n30 = porN(unsat, 30),
  n28 = porN(unsat, 28),
  n26 = porN(unsat, 26),
  n32u = porN(unsat, 32),
  n32p = porN(planted, 32);
const nMax = D.nMax;
const minutos = (s) => fn(s / 60, 1);
const unsatTerminaron = unsat.filter(D.terminado);
const ultimo = D.ultimoUnsat;
const walkSegundos = (n) => D.escala.find((e) => e.n === n);
const e5000 = walkSegundos(5000);
const err = (x) => fn(x * 100, 1) + " %";

// ===== 1. Portada =====
{
  const s = nueva(OSCURO);
  tx(
    s,
    "TEORÍA DE LA COMPUTABILIDAD · TRABAJO FINAL 2026 · UTN FRCU",
    0.6,
    1.0,
    8.8,
    0.4,
    { color: "F0B77A" },
  );
  tx(s, "SAT (3-SAT): algoritmo exacto frente a WalkSAT", 0.6, 1.6, 8.8, 1.6, {
    fs: 38,
    face: "Georgia",
    color: "FFFFFF",
    valign: "middle",
  });
  s.addShape(pptx.ShapeType.rect, {
    x: 0.6,
    y: 3.4,
    w: 1.2,
    h: 0.05,
    fill: { color: NARANJA },
    line: { color: NARANJA, width: 0 },
  });
  tx(s, "Integrantes: Facundo Iglesias y Camilo Escar", 0.6, 3.75, 8.8, 0.35, { color: "D7DEE4" });
  tx(s, "Materia: Teoría de la Computabilidad · Docente: Prof. Pascal", 0.6, 4.1, 8.8, 0.35, { color: "D7DEE4" });
  tx(s, "Fecha: [completar fecha de exposición]", 0.6, 4.45, 8.8, 0.35, { color: "D7DEE4" });
}

// ===== 2. Problema =====
{
  const s = nueva();
  titulo(s, "El problema SAT y por qué importa");
  tx(
    s,
    "Dada una fórmula booleana, ¿existe una asignación de verdadero/falso que la haga verdadera?",
    0.5,
    1.2,
    9,
    0.65,
  );
  s.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 1.95,
    w: 9,
    h: 0.6,
    fill: { color: OSCURO },
    line: { color: OSCURO, width: 0 },
  });
  tx(
    s,
    "F = (x1 ∨ ¬x2 ∨ x3) ∧ (¬x1 ∨ x2 ∨ x4) ∧ (x2 ∨ ¬x3 ∨ ¬x4)",
    0.6,
    1.95,
    8.8,
    0.6,
    { face: "Courier New", color: CODIGO, valign: "middle" },
  );
  const w = 2.85,
    xs = [0.5, 3.575, 6.65];
  bloque(
    s,
    xs[0],
    2.8,
    w,
    2.45,
    "NP-completo",
    "Primer problema NP-completo (Cook 1971, Levin). 3-SAT también lo es.",
  );
  bloque(
    s,
    xs[1],
    2.8,
    w,
    2.45,
    "Verificar es fácil",
    "Dada una asignación se comprueba en tiempo polinomial: está en NP.",
  );
  bloque(
    s,
    xs[2],
    2.8,
    w,
    2.45,
    "Buscar es difícil",
    "Con n variables hay 2ⁿ asignaciones posibles: ese es el espacio de búsqueda.",
  );
}

// ===== 3. Fuerza bruta =====
{
  const s = nueva();
  titulo(s, "Algoritmo exacto: fuerza bruta");
  s.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 1.3,
    w: 5.6,
    h: 3.9,
    fill: { color: CODIGO },
    line: { color: BORDE, width: 1 },
  });
  tx(
    s,
    [
      "para cada asignación de 0 a 2ⁿ − 1:",
      "  para cada cláusula:",
      "    si ningún literal es verdadero:",
      "      descartar la asignación",
      "  si cumplió todas:",
      "    devolver SAT y la asignación",
      "devolver UNSAT",
    ].map((t) => ({ t })),
    0.65,
    1.45,
    5.4,
    3.6,
    { face: "Courier New", fs: 18, color: OSCURO },
  );
  tx(
    s,
    [
      { t: "Complejidad", bold: true, color: NARANJA },
      { t: "O(2ⁿ · m · k)", bold: true, gap: 0.4 },
      { t: "n variables, m cláusulas, k literales por cláusula.", gap: 0.4 },
      { t: "Exacta: si no hay solución, lo demuestra (UNSAT).", gap: 0.4 },
      { t: "Cada variable más duplica el trabajo." },
    ],
    6.4,
    1.3,
    3.2,
    3.9,
  );
}

// ===== 4. Hardware y metodologia =====
{
  const s = nueva();
  titulo(s, "Equipo y metodología");
  bloque(
    s,
    0.5,
    1.25,
    3.9,
    4.0,
    "Equipo de medición",
    [
      { t: `${hw.cpu}` },
      { t: `${hw.nucleosFisicos} núcleos, ${hw.nucleosLogicos} lógicos` },
      { t: `RAM ${hw.ram}` },
      { t: `${hw.so}, Python ${hw.python}` },
      { t: `Límite por corrida: ${fn(hw.limite)} s (1 h)` },
      { t: `Corrido el ${hw.fecha}` },
    ],
    { bullet: true },
  );
  bloque(
    s,
    4.6,
    1.25,
    4.9,
    4.0,
    "Metodología",
    [
      { t: "Mismas fórmulas (DIMACS, semillas fijas) para los dos algoritmos" },
      {
        t: `n = 10 a 32 de a 2: una planted (con solución) y una insatisfacible`,
      },
      { t: "WalkSAT: 10 corridas por fórmula, p = 0,5, 100.000 pasos" },
      {
        t: `Escala: n = 100 a ${fn(nMax.n)}, 3 planted + 3 aleatorias por n, hasta 5.000.000 de pasos`,
      },
    ],
    { bullet: true },
  );
}

// ===== 5. Resultados del exacto =====
{
  const s = nueva();
  titulo(s, "Resultado: tiempo de la fuerza bruta");
  const fondo = imagen(s, "1_tiempo_fuerza_bruta.png", 0.4, 1.2, 5.3);
  tx(
    s,
    "Insatisfacible: recorre las 2ⁿ. La x roja es TIMEOUT: no es un tiempo medido.",
    0.4,
    fondo + 0.1,
    5.3,
    0.9,
  );
  tx(
    s,
    [
      {
        t: `n=${ultimo.n}: ${fn(ultimo.t_exacto_s)} s (${minutos(ultimo.t_exacto_s)} min)`,
        bullet: true,
        gap: 0.3,
      },
      {
        t: `n=${n32u.n}: no terminó en 1 h (cortado a ${fn(n32u.t_exacto_s, 0)} s)`,
        bullet: true,
        bold: true,
        color: NARANJA,
        gap: 0.3,
      },
      {
        t: `t(n+2)/t(n) ≈ ${fn(razon.gm, 1)} (de ${fn(razon.min, 1)} a ${fn(razon.max, 1)}, n=${razon.desde} a ${razon.hasta}): ≈ 4 = 2²`,
        bullet: true,
        gap: 0.3,
      },
      {
        t: `n=${n28.n}: ${fn(n28.t_exacto_s)} s; n=${n26.n}: ${fn(n26.t_exacto_s)} s`,
        bullet: true,
        gap: 0.3,
      },
      {
        t: `Con solución, n=${n32p.n}: ${fn(n32p.t_exacto_s)} s (frena al hallarla)`,
        bullet: true,
      },
    ],
    5.9,
    1.25,
    3.7,
    4.1,
  );
}

// ===== 6. WalkSAT =====
{
  const s = nueva();
  titulo(s, "Heurística: WalkSAT");
  tx(s, "Cómo funciona", 0.5, 1.2, 4.4, 0.4, { bold: true, color: NARANJA });
  tx(
    s,
    [
      "1) Asignación inicial al azar.",
      "2) Elegir una cláusula insatisfecha.",
      "3) Dar vuelta una variable de esa cláusula: la que rompe menos, o una al azar con probabilidad p.",
      "4) Repetir hasta cumplir todo o agotar los pasos.",
    ].map((t) => ({ t, gap: 0.2 })),
    0.5,
    1.65,
    4.4,
    2.8,
  );
  tx(s, "Es incompleta", 5.2, 1.2, 4.3, 0.4, { bold: true, color: NARANJA });
  tx(
    s,
    [
      {
        t: "Si no encuentra solución devuelve UNKNOWN. Eso NO significa que la fórmula sea insatisfacible.",
        gap: 0.3,
      },
      { t: "Solo el algoritmo exacto puede demostrarlo." },
    ],
    5.2,
    1.65,
    4.3,
    2.8,
  );
  tarjeta(s, 0.5, 4.5, 9, 0.8, CODIGO);
  tx(
    s,
    "Parámetros: p = 0,5 · 100.000 pasos (n ≤ 32) o 5.000.000 (n ≥ 100) · semillas fijas.",
    0.65,
    4.5,
    8.7,
    0.8,
    { valign: "middle" },
  );
}

// ===== 7. Comparacion: grafico =====
{
  const s = nueva();
  titulo(s, "Exacto vs. WalkSAT: mismas fórmulas");
  imagen(s, "3_comparacion_tiempos.png", 0.4, 1.2, 5.6);
  tx(
    s,
    [
      {
        t: `Con solución: WalkSAT ${D.todosPlantedDiez ? "10/10" : "?"} en todos los n, como máximo ${ft(D.tmaxWalkSat)}`,
        bullet: true,
        gap: 0.3,
      },
      {
        t: `n=${n32p.n}: exacto ${fn(n32p.t_exacto_s)} s; WalkSAT ${ft(n32p.t_walksat_medio_s)}`,
        bullet: true,
        gap: 0.3,
      },
      {
        t: `Insatisfacible: exacto demuestra UNSAT hasta n=${ultimo.n}; con n=${n32u.n} no terminó en 1 h`,
        bullet: true,
      },
    ],
    6.2,
    1.25,
    3.4,
    4.0,
  );
}

// ===== 8. Comparacion: tabla =====
{
  const s = nueva();
  titulo(s, "Éxitos, error relativo y tiempos");
  const h = (t) => ({
    text: t,
    options: { bold: true, color: "FFFFFF", fill: { color: OSCURO } },
  });
  const filas = [
    [
      h("n"),
      h("Exacto SAT"),
      h("Exacto UNSAT"),
      h("WalkSAT SAT"),
      h("WalkSAT UNSAT"),
    ],
  ];
  for (const n of [20, 26, 30, 32]) {
    const p = porN(planted, n),
      u = porN(unsat, n);
    const uCell = D.terminado(u)
      ? { text: ft(u.t_exacto_s) }
      : { text: "no terminó en 1 h", options: { bold: true, color: NARANJA } };
    filas.push([
      { text: String(n), options: { bold: true } },
      { text: ft(p.t_exacto_s) },
      uCell,
      { text: `${p.walksat_exitos_de_10}/10 · ${ft(p.t_walksat_medio_s)}` },
      { text: `${u.walksat_exitos_de_10}/10 · ${err(u.error_relativo)}` },
    ]);
  }
  s.addTable(filas, {
    x: 0.5,
    y: 1.25,
    w: 9,
    colW: [0.7, 1.5, 2.5, 2.4, 1.9],
    rowH: 0.42,
    fontFace: "Arial",
    fontSize: 18,
    color: OSCURO,
    border: { type: "solid", color: BORDE, pt: 1 },
    fill: { color: "FFFFFF" },
    margin: [0, 6, 0, 6],
    valign: "middle",
  });
  const d = D.densidad;
  tx(
    s,
    [
      {
        t: "WalkSAT: éxitos/10 · tiempo (SAT) o error relativo (UNSAT).",
        gap: 0.3,
      },
      {
        t: `En UNSAT agotó los 100.000 pasos (≈ ${ft(D.walkUnsatT.reduce((a, b) => a + b) / D.walkUnsatT.length)}) y dejó ${err(D.errUnsat.min)} a ${err(D.errUnsat.max)} de cláusulas sin cumplir (n=10 a 32).`,
        gap: 0.3,
      },
      {
        t: `n=16, razón m/n de ${fn(d.rMin, 0)} a ${fn(d.rMax, 0)}, ${d.porRazon} fórmulas c/u: halló la solución en ${d.walkTotal} de ${d.satTotal} satisfacibles y dijo SAT falsamente ${d.falsos} veces.`,
      },
    ],
    0.5,
    3.7,
    9,
    1.75,
  );
}

// ===== 9. Escalabilidad =====
{
  const s = nueva();
  titulo(s, "Escalabilidad de WalkSAT");
  const fondo = imagen(s, "2_tiempo_walksat.png", 0.4, 1.2, 4.6);
  const h = (t) => ({
    text: t,
    options: { bold: true, color: "FFFFFF", fill: { color: OSCURO } },
  });
  const filas = [[h("n"), h("Éxitos"), h("t medio")]];
  for (const e of D.escala)
    filas.push([
      { text: fn(e.n), options: { bold: true } },
      {
        text: `${e.ok}/${e.corridas}`,
        options: e.ok === 0 ? { bold: true, color: NARANJA } : {},
      },
      { text: ft(e.tMedio) },
    ]);
  s.addTable(filas, {
    x: 5.5,
    y: 1.25,
    w: 4.0,
    colW: [1.3, 1.2, 1.5],
    rowH: 0.4,
    fontFace: "Arial",
    fontSize: 18,
    color: OSCURO,
    border: { type: "solid", color: BORDE, pt: 1 },
    fill: { color: "FFFFFF" },
    margin: [0, 6, 0, 6],
    valign: "middle",
  });
  const p = e5000,
    minPct = fn(100 - nMax.pctMin, 2);
  tx(
    s,
    [
      {
        t: `n=${fn(nMax.n)}: ${nMax.ok} de ${nMax.corridas} llegó a la solución en ${fn(nMax.flipsMax)} pasos (≈ ${ft(nMax.tMax)}); quedó hasta ${minPct} % sin cumplir.`,
        gap: 0.2,
      },
      {
        t: `Es el tope del experimento, no prueba de UNSAT. n=5.000: planted ${p.okPlanted}/${p.nPlanted}, aleatoria ${p.okAleat}/${p.nAleat}.`,
      },
    ],
    0.4,
    fondo + 0.1,
    9.2,
    1.3,
  );
}

// ===== 10. Casos reales =====
{
  const s = nueva();
  titulo(s, "Aplicación real: horarios en Argentina");
  bloque(
    s,
    0.5,
    1.25,
    4.4,
    3.4,
    "UTDT: horarios de clases",
    "Ferreiro y Marenco (53 JAIIO/SIIIO): programación lineal entera (ILP) para armar horarios, 35 módulos. Impacto solo cualitativo. No es SAT puro, pero se codifica en CNF.",
  );
  bloque(
    s,
    5.1,
    1.25,
    4.4,
    3.4,
    "UNR: coloreo de grafos y SAT",
    "Escalante, Chiesa y Fekete (SUMA 2019). Colorear un grafo con k colores se reduce a SAT: x(v,c) = el vértice v lleva el color c; adyacentes, colores distintos.",
  );
  tx(s, "Puente: horarios → coloreo de grafos → SAT", 0.5, 4.8, 9, 0.5, {
    bold: true,
    color: TEAL,
    valign: "middle",
  });
}

// ===== 11. Usos y herramientas =====
{
  const s = nueva();
  titulo(s, "Usos industriales y herramientas");
  bloque(
    s,
    0.5,
    1.25,
    4.9,
    4.0,
    "Usos documentados",
    [
      { t: "Bounded model checking (verificación de hardware y software)" },
      { t: "Dependencias de paquetes: Eclipse (Sat4j), libsolv" },
      { t: "Timetabling con SAT/MaxSAT (Asín Achá y Nieuwenhuis, 2014)" },
    ],
    { bullet: true },
  );
  bloque(
    s,
    5.6,
    1.25,
    3.9,
    4.0,
    "Herramientas",
    [
      { t: "MiniSat, Glucose" },
      { t: "Kissat (ganó la SAT Competition 2024)" },
      { t: "Z3" },
      { t: "OR-Tools CP-SAT" },
    ],
    { bullet: true },
  );
}

// ===== 12. FRCU hipotetico + impacto =====
{
  const s = nueva();
  titulo(s, "Modelo FRCU e impacto económico");
  bloque(s, 0.5, 1.25, 4.6, 4.05, "Modelo HIPOTÉTICO (FRCU)", [
    {
      t: "x(c,f): la comisión c va en la franja f. Al menos una franja, a lo sumo una, sin choques de docente ni de alumnos.",
      gap: 0.2,
    },
    {
      t: "Ejemplo: 9 variables y 19 cláusulas, resuelto por ambos algoritmos.",
      gap: 0.2,
    },
    {
      t: "No hay evidencia pública de que la FRCU use SAT.",
      bold: true,
      color: NARANJA,
    },
  ]);
  bloque(s, 5.3, 1.25, 4.2, 4.05, "Impacto económico", [
    { t: "Sin cifra pública atribuible a SAT.", bold: true, gap: 0.2 },
    {
      t: "Analogía (no es SAT): horario ferroviario neerlandés, 50 M€ por año (Kroon, Huisman et al., Interfaces 2009).",
      gap: 0,
    },
  ]);
}

// ===== 13. Viabilidad =====
{
  const s = nueva();
  titulo(s, "Viabilidad de la aplicación");
  const w = 4.4;
  bloque(
    s,
    0.5,
    1.25,
    w,
    2.0,
    "Técnica",
    `Fuerza bruta no escala (n=${n32u.n}: más de 1 h). Para producción: solvers como MiniSat o Kissat.`,
  );
  bloque(
    s,
    5.1,
    1.25,
    w,
    2.0,
    "Datos",
    "Hoy solo hay PDFs de horarios. Faltan docentes, aulas y capacidades en formato usable.",
  );
  bloque(
    s,
    0.5,
    3.4,
    w,
    1.95,
    "Económica",
    "Costo bajo: software libre. Ahorro: sin cifra, solo cualitativo.",
  );
  bloque(
    s,
    5.1,
    3.4,
    w,
    1.95,
    "Límite del modelo",
    "Preferencias (huecos, carga) piden MaxSAT u ILP, no SAT puro.",
  );
}

// ===== 14. Reflexion critica =====
{
  const s = nueva();
  titulo(s, "Reflexión crítica");
  bloque(
    s,
    0.5,
    1.25,
    5.0,
    4.05,
    "Límites de nuestro trabajo",
    [
      { t: "Una fórmula y una corrida del exacto por tamaño" },
      {
        t: `Planted irregular: n=30 (${ft(porN(planted, 30).t_exacto_s)}) más rápido que n=28 (${ft(porN(planted, 28).t_exacto_s)})`,
      },
      { t: "Una sola máquina, en Python" },
      { t: "p = 0,5 sin ajustar; 3 fórmulas por punto en escala" },
      { t: "Aplicación hipotética" },
    ],
    { bullet: true },
  );
  bloque(
    s,
    5.7,
    1.25,
    3.8,
    4.05,
    "Qué haríamos",
    [
      { t: "Repetir y promediar" },
      { t: "Comparar con MiniSat" },
      { t: "Usar benchmarks estándar" },
      { t: "Probar otros valores de p" },
    ],
    { bullet: true },
  );
}

// ===== 15. Conclusiones =====
{
  const s = nueva(OSCURO);
  s.addText("Conclusiones", {
    x: 0.5,
    y: 0.3,
    w: 9,
    h: 0.7,
    fontFace: "Georgia",
    fontSize: 28,
    color: "FFFFFF",
    margin: 0,
    valign: "middle",
  });
  const cols = [
    [
      "> 1 h",
      `Exacto con n=${n32u.n} insatisfacible: no terminó. n=${ultimo.n} tardó ${minutos(ultimo.t_exacto_s)} min.`,
    ],
    [
      ft(n32p.t_walksat_medio_s),
      `WalkSAT en la fórmula con solución de n=${n32p.n} (exacto: ${fn(n32p.t_exacto_s)} s).`,
    ],
    [
      "UNKNOWN ≠ UNSAT",
      "La heurística gana velocidad y pierde la garantía de demostrar que no hay solución.",
    ],
  ];
  cols.forEach(([g, t], i) => {
    const x = 0.5 + i * 3.1;
    tx(s, g, x, 1.3, 2.9, 0.9, {
      fs: i === 2 ? 22 : 36,
      face: "Georgia",
      bold: true,
      color: "F0B77A",
      valign: "middle",
    });
    tx(s, t, x, 2.35, 2.8, 1.5, { color: "D7DEE4" });
  });
  tx(
    s,
    "Exacto cuando n es chico o hay que demostrar UNSAT; WalkSAT cuando n crece y basta una solución. En la práctica se usan solvers industriales.",
    0.5,
    4.0,
    9,
    1.3,
    { color: "FFFFFF" },
  );
}

// ===== 16. Referencias =====
{
  const s = nueva();
  titulo(s, "Referencias");
  tx(
    s,
    [
      "Cook, S. (1971). The complexity of theorem-proving procedures. STOC.",
      "Selman, Kautz y Cohen (1994). Noise strategies for improving local search. AAAI-94.",
      "Ferreiro y Marenco. Horarios de clases con ILP, UTDT. 53 JAIIO/SIIIO.",
      "Escalante, Chiesa y Fekete. Coloreo de grafos y SAT. SUMA 2019, UNR.",
      "Asín Achá y Nieuwenhuis (2014). Curriculum-based course timetabling with SAT and MaxSAT. Ann. Oper. Res.",
      "Kroon, Huisman et al. (2009). The new Dutch timetable: the OR revolution. Interfaces 39(1).",
      "SAT Competition 2024 (resultados).",
    ].map((t) => ({ t, gap: 0.15 })),
    0.5,
    1.25,
    9,
    4.1,
  );
}

const salida = path.join(D.RAIZ, "SAT_Presentacion.pptx");
pptx.writeFile({ fileName: salida }).then(() => {
  console.log("OK", salida, `(${slideActual} diapositivas)`);
  if (avisos.length)
    console.warn("AVISOS de desborde estimado:\n" + avisos.join("\n"));
});
