// Genera SAT_Informe.docx. Uso (desde la raiz del proyecto):  node generadores/informe.js
// Todos los numeros de resultados se leen de hardware.txt, resultados_*.csv e historial_n500.csv.
const fs = require("fs");
const path = require("path");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  ImageRun,
  HeadingLevel,
  AlignmentType,
  PageBreak,
  Footer,
  PageNumber,
  ShadingType,
  BorderStyle,
  LevelFormat,
} = require("docx");

const RAIZ = path.resolve(__dirname, "..");
const leer = (f) => { // utf-8; si falla (CSV viejos en cp1252) cae a latin1
  const b = fs.readFileSync(path.join(RAIZ, f));
  try { return new TextDecoder("utf-8", { fatal: true }).decode(b).replace(/^﻿/, ""); }
  catch (e) { return b.toString("latin1"); }
};
const csv = (f) => {
  const [cab, ...filas] = leer(f).trim().split(/\r?\n/);
  const k = cab.split(",");
  return filas.map((l) =>
    Object.fromEntries(l.split(",").map((v, i) => [k[i], v.trim()])),
  );
};

// ---------------------------------------------------------------- datos
const hw = Object.fromEntries(
  leer("hardware.txt")
    .trim()
    .split(/\r?\n/)
    .map((l) => {
      const i = l.indexOf(":");
      return [l.slice(0, i).trim(), l.slice(i + 1).trim()];
    }),
);
const exacto = csv("resultados_exacto.csv")
  .map((r) => ({
    ...r,
    n: +r.n,
    m: +r.m,
    t: +r.t_exacto_s,
    asig: +r.asignaciones,
    claus: +r.clausulas_evaluadas,
    exitos: +r.walksat_exitos_de_10,
    tw: +r.walksat_t_medio_s || +r.t_walksat_medio_s,
    flips: +r.flips_medio,
    err: +r.error_relativo,
    mejor: +r.mejor_pct_clausulas,
  }))
  .sort((a, b) => a.n - b.n || (a.tipo.startsWith("planted") ? -1 : 1));
const escala = csv("resultados_escala.csv").map((r) => ({
  ...r,
  n: +r.n,
  m: +r.m,
  flips: +r.flips,
  t: +r.t_s,
  pct: +r.pct_clausulas,
}));
const densidad = csv("resultados_densidad.csv");
const hist = csv("historial_n500.csv").map((r) => +r.insatisfechas);
const LIM = parseFloat(hw["LIMITE (s)"]);

// ---------------------------------------------------------------- formato es-AR
const miles = (s) => s.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
const num = (x, d = 0) => {
  // 1234.5 -> "1.234,5"
  const [e, f] = Math.abs(x).toFixed(d).split(".");
  return (x < 0 ? "-" : "") + miles(e) + (f ? "," + f : "");
};
const seg = (x) =>
  (x >= 100
    ? num(x, 1)
    : x >= 1
      ? num(x, 2)
      : x >= 0.01
        ? num(x, 3)
        : num(x, 5)) + " s";
const dur = (s) =>
  s >= 3600
    ? num(s / 3600, 1) + " h"
    : s >= 120
      ? num(s / 60, 0) + " min"
      : num(s, 0) + " s";
const esTimeout = (r) => r.estado_exacto === "TIMEOUT";

const ins = exacto.filter((r) => r.tipo === "insatisfacible");
const pla = exacto.filter((r) => r.tipo.startsWith("planted"));
const medidas = ins.filter((r) => !esTimeout(r));
const nMax = Math.max(...exacto.map((r) => r.n));
const nMaxMedidoUnsat = Math.max(...medidas.map((r) => r.n));
const get = (arr, n) => arr.find((r) => r.n === n);
const to32 = exacto.find((r) => esTimeout(r));
if (!to32) console.warn("AVISO: no hay TIMEOUT en resultados_exacto.csv");

// razones t(n+2)/t(n) en las insatisfacibles medidas (el exacto recorre las 2^n)
const razones = [];
for (let i = 0; i + 1 < medidas.length; i++)
  razones.push({ n: medidas[i].n, r: medidas[i + 1].t / medidas[i].t });
const gm = (a) => Math.exp(a.reduce((s, x) => s + Math.log(x), 0) / a.length);
const razGrandes = razones.filter((x) => x.n >= 18).map((x) => x.r);
const razMedia = gm(razGrandes);
const razMin = Math.min(...razGrandes),
  razMax = Math.max(...razGrandes);
const ultimo = medidas[medidas.length - 1];
const kAjuste = ultimo.t / Math.pow(2, ultimo.n); // t ~ k 2^n con el mayor n medido
const proy = (n) => kAjuste * Math.pow(2, n);
const tasaTO = to32 ? to32.asig / to32.t : null; // asignaciones/s cuando se cortó
const fracTO = to32 ? to32.asig / Math.pow(2, to32.n) : null;
const proyTO = to32 ? Math.pow(2, to32.n) / tasaTO : null;

// escalabilidad: agregado por (n, tipo)
const grupos = [];
for (const n of [...new Set(escala.map((r) => r.n))]) {
  for (const tipo of ["planted", "aleatoria"]) {
    const g = escala.filter((r) => r.n === n && r.tipo === tipo);
    if (!g.length) continue;
    const ok = g.filter((r) => r.estado === "SAT");
    const sin = g
      .filter((r) => r.estado !== "SAT")
      .map((r) => Math.round(((100 - r.pct) / 100) * r.m));
    grupos.push({
      n,
      tipo,
      m: g[0].m,
      k: ok.length,
      tot: g.length,
      flips: g.reduce((s, r) => s + r.flips, 0) / g.length,
      t: g.reduce((s, r) => s + r.t, 0) / g.length,
      sinMin: sin.length ? Math.min(...sin) : null,
      sinMax: sin.length ? Math.max(...sin) : null,
      tMax: Math.max(...g.map((r) => r.t)),
    });
  }
}
const nEsc = Math.max(...escala.map((r) => r.n));
const g10 = grupos.filter((g) => g.n === nEsc);
const fallo10 = g10.every((g) => g.k === 0);
const maxFlipsEsc = Math.max(...escala.map((r) => r.flips));
const tMaxEsc = Math.max(...escala.map((r) => r.t));

// ---------------------------------------------------------------- helpers de documento
const FUENTE = "Calibri";
const runs = (s, o = {}) =>
  s.split("**").map(
    (t, i) =>
      new TextRun({
        text: t,
        bold: i % 2 === 1 || o.bold,
        italics: o.italics,
        size: o.size,
        font: o.font,
      }),
  );
const P = (s, o = {}) =>
  new Paragraph({
    children: runs(s, o),
    spacing: { after: 120, line: 276 },
    alignment: o.align || AlignmentType.JUSTIFIED,
    ...(o.p || {}),
  });
const H1 = (s) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun(s)],
    keepNext: true,
  });
const H2 = (s) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun(s)],
    keepNext: true,
  });
const B = (s) =>
  new Paragraph({
    numbering: { reference: "viñetas", level: 0 },
    children: runs(s),
    spacing: { after: 80, line: 276 },
  });
const N = (s) =>
  new Paragraph({
    numbering: { reference: "numeros", level: 0 },
    children: runs(s),
    spacing: { after: 80, line: 276 },
  });
const COD = (s) =>
  new Paragraph({
    children: [new TextRun({ text: s, font: "Consolas", size: 19 })],
    spacing: { after: 40 },
    indent: { left: 360 },
    shading: { type: ShadingType.CLEAR, fill: "F2F2F2" },
  });
const REF = (s) =>
  new Paragraph({
    children: runs(s),
    spacing: { after: 100 },
    indent: { left: 360, hanging: 360 },
  });
const NOTA = (s) =>
  new Paragraph({
    children: runs(s, { italics: true }),
    spacing: { after: 120 },
    shading: { type: ShadingType.CLEAR, fill: "FFF4D6" },
    border: {
      left: { style: BorderStyle.SINGLE, size: 12, color: "C58B00", space: 6 },
    },
  });

const borde = { style: BorderStyle.SINGLE, size: 4, color: "999999" };
const bordes = { top: borde, bottom: borde, left: borde, right: borde };
const TOT = 9026;
function tabla(cab, filas, anchos, o = {}) {
  const suma = anchos.reduce((a, b) => a + b, 0);
  const w = anchos.map((x) => Math.round((x * TOT) / suma));
  const sz = o.size || 18;
  const celda = (t, i, head, fill) =>
    new TableCell({
      width: { size: w[i], type: WidthType.DXA },
      borders: bordes,
      margins: { top: 50, bottom: 50, left: 70, right: 70 },
      shading: head
        ? { type: ShadingType.CLEAR, fill: "1F4E79" }
        : fill
          ? { type: ShadingType.CLEAR, fill }
          : undefined,
      children: [
        new Paragraph({
          alignment:
            o.izq && o.izq.includes(i)
              ? AlignmentType.LEFT
              : AlignmentType.CENTER,
          children: [
            new TextRun({
              text: String(t),
              size: sz,
              bold: head,
              color: head ? "FFFFFF" : undefined,
            }),
          ],
        }),
      ],
    });
  return new Table({
    width: { size: TOT, type: WidthType.DXA },
    columnWidths: w,
    rows: [
      new TableRow({
        tableHeader: true,
        children: cab.map((c, i) => celda(c, i, true)),
      }),
      ...filas.map(
        (f) =>
          new TableRow({
            cantSplit: true,
            children: f.celdas.map((c, i) => celda(c, i, false, f.fill)),
          }),
      ),
    ],
  });
}
const fila = (celdas, fill) => ({ celdas, fill });
const espacio = () => new Paragraph({ children: [], spacing: { after: 120 } });
const TITULO_FIG = (s) =>
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: s, italics: true, size: 19 })],
    spacing: { after: 200 },
  });
const faltantes = [];
function fig(archivo, epigrafe) {
  const ruta = path.join(RAIZ, "graficos", archivo);
  if (!fs.existsSync(ruta)) {
    faltantes.push(archivo);
    return [NOTA(`[falta el gráfico ${archivo}: correr graficos.py]`)];
  }
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      keepNext: true,
      spacing: { before: 120, after: 60 },
      children: [
        new ImageRun({
          type: "png",
          data: fs.readFileSync(ruta),
          transformation: { width: 520, height: 319 },
          altText: { title: epigrafe, description: epigrafe, name: archivo },
        }),
      ],
    }),
    TITULO_FIG(epigrafe),
  ];
}

// ---------------------------------------------------------------- contenido
const c = [];

// Portada
c.push(
  new Paragraph({
    spacing: { before: 2400, after: 200 },
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({
        text: "Teoría de la Computabilidad · Trabajo Final 2026",
        size: 26,
        color: "595959",
      }),
    ],
  }),
  new Paragraph({
    spacing: { after: 200 },
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({
        text: "SAT (Problema de Satisfacción Booleana): algoritmo exacto frente a la heurística WalkSAT",
        bold: true,
        size: 44,
        color: "1F4E79",
      }),
    ],
  }),
  new Paragraph({
    spacing: { after: 1200 },
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({
        text: "Problema NP-completo con instancias restringidas a 3-SAT: fuerza bruta como solución exacta y WalkSAT (búsqueda local) como heurística",
        italics: true,
        size: 26,
      }),
    ],
  }),
  new Paragraph({
    spacing: { after: 100 },
    children: [new TextRun({ text: "Integrantes: Facundo Iglesias y Camilo Escar", size: 24 })],
  }),
  new Paragraph({
    spacing: { after: 100 },
    children: [new TextRun({ text: "Materia: Teoría de la Computabilidad", size: 24 })],
  }),
  new Paragraph({
    spacing: { after: 100 },
    children: [new TextRun({ text: "Docente: Prof. Pascal", size: 24 })],
  }),
  new Paragraph({
    spacing: { after: 100 },
    children: [new TextRun({ text: "Fecha: [completar fecha de exposición]", size: 24 })],
  }),
  new Paragraph({ children: [new PageBreak()] }),
);

// 1
c.push(H1("1. Introducción y relevancia"));
c.push(
  P(
    "El problema de satisfacibilidad booleana (SAT) pregunta si existe una asignación de verdadero o falso para las variables de una fórmula que la haga verdadera. Fue el primer problema que se demostró NP-completo (Cook, 1971; Levin, 1973, de forma independiente) y por eso es el punto de partida para estudiar la dificultad de toda una clase de problemas.",
  ),
  P(
    "La consigna del trabajo propone SAT como problema a estudiar. Como SAT admite cláusulas de cualquier longitud, las instancias de los experimentos se restringen a 3-SAT (exactamente tres literales por cláusula), que sigue siendo NP-completo y permite generar y comparar fórmulas de forma uniforme. Sobre ese problema se contrastan dos formas de resolverlo: un algoritmo exacto de fuerza bruta, que sirve como línea de base, y WalkSAT, la heurística elegida, que es una búsqueda local.",
  ),
  P(
    "SAT importa también en la práctica: muchos problemas de decisión se pueden traducir a una fórmula booleana y entregar a un solver. Entre los usos documentados están la verificación de hardware y software mediante bounded model checking (Biere, Cimatti, Clarke y Zhu, 1999), la planificación como satisfacibilidad (Kautz y Selman, 1992), la resolución de dependencias de paquetes (por ejemplo, Eclipse con Sat4j, Le Berre y Parrain, 2010, y libsolv en openSUSE) y la programación de horarios con SAT y MaxSAT (Asín Achá y Nieuwenhuis, 2014). La sección 10 desarrolla estos casos.",
  ),
  P(
    "El objetivo de este trabajo es comprender la dificultad de SAT de forma experimental. Para resolverlo se programó, con asistencia de una herramienta de IA (ver sección 7.4), un algoritmo exacto de fuerza bruta y una heurística de búsqueda local (WalkSAT), se los comparó sobre exactamente las mismas fórmulas, se midió hasta dónde llega cada uno y se analizó una aplicación posible: la asignación de horarios de cursada.",
  ),
  P(
    `Todos los números de resultados de este informe salen de ejecuciones reales del código entregado, en el equipo descripto en la sección 7 (${hw["CPU"]}). Los tiempos de mediciones anteriores, hechas en otra máquina, se descartaron.`,
  ),
);

// 2
c.push(H1("2. Marco teórico"));
c.push(
  H2("2.1 Clases P, NP y NP-completo"),
  B(
    "**P:** problemas que se resuelven en tiempo polinomial con una máquina de Turing determinística.",
  ),
  B(
    "**NP:** problemas cuya solución propuesta se puede verificar en tiempo polinomial.",
  ),
  B("**NP-difícil:** todo problema de NP se reduce a él en tiempo polinomial."),
  B(
    "**NP-completo:** está en NP y es NP-difícil. No se conoce un algoritmo exacto polinomial; saber si P = NP es un problema abierto.",
  ),
  H2("2.2 Cook-Levin y Karp"),
  P(
    "Cook (1971) demostró que SAT es NP-completo: cualquier problema de NP se reduce en tiempo polinomial a SAT, porque el cómputo de una máquina de Turing no determinística con límite polinomial se puede describir con una fórmula booleana de tamaño polinomial. Levin (1973) obtuvo un resultado equivalente de forma independiente, por eso se habla del teorema de Cook-Levin. Karp (1972) mostró que 21 problemas combinatoriales clásicos son NP-completos mediante reducciones a partir de SAT, y con eso estableció que la NP-completitud aparece en muchísimos problemas distintos.",
  ),
  H2("2.3 SAT, 3-SAT y Max-SAT"),
  P(
    "Una fórmula en forma normal conjuntiva (CNF) es una conjunción (AND) de cláusulas; cada cláusula es una disyunción (OR) de literales; un literal es una variable o su negación.",
  ),
  B(
    "**SAT:** cláusulas de cualquier longitud. ¿Existe una asignación que cumpla todas?",
  ),
  B(
    "**3-SAT:** cada cláusula tiene exactamente tres literales. Sigue siendo NP-completo.",
  ),
  B(
    "**Max-SAT:** se busca la asignación que cumpla la mayor cantidad posible de cláusulas. Es un problema de optimización NP-difícil.",
  ),
  H2("2.4 Por qué SAT está en NP"),
  P(
    "Dada una asignación, basta recorrer las m cláusulas y comprobar que cada una tiene un literal verdadero. Eso lleva O(m · k) pasos, polinomial en el tamaño de la fórmula. Por lo tanto, una solución se verifica rápido.",
  ),
  H2("2.5 Por qué 3-SAT es NP-completo"),
  P(
    "Para llegar a 3-SAT se reduce SAT a 3-SAT en tiempo polinomial: una cláusula larga (a ∨ b ∨ c ∨ d) se reemplaza por (a ∨ b ∨ y) ∧ (¬y ∨ c ∨ d) con una variable nueva y, y se repite hasta que todas tengan tres literales; las cláusulas cortas se completan repitiendo literales o con variables auxiliares. La fórmula original es satisfacible si y solo si la nueva lo es. Como la transformación es polinomial, 3-SAT es al menos tan difícil como SAT. Por otro lado, 3-SAT está en NP, por lo que es NP-completo. (Una observación útil: 2-SAT, con dos literales por cláusula, sí se resuelve en tiempo polinomial.)",
  ),
);

// 3
c.push(H1("3. Formulación matemática"));
c.push(
  P(
    "Sea X = {x₁, …, xₙ} un conjunto de n variables booleanas. Un literal es xᵢ o ¬xᵢ. Una cláusula es una disyunción de literales y una fórmula F es una conjunción de m cláusulas. Una asignación es una función que le da un valor verdadero o falso a cada variable. El problema es: dada F, decidir si existe una asignación que la haga verdadera. En k-SAT cada cláusula tiene exactamente k literales (en este trabajo, k = 3).",
  ),
  P(
    "En el programa una fórmula se guarda como una lista de listas de enteros con signo: 3 representa x₃ y −3 representa ¬x₃. Por ejemplo, F = (x₁ ∨ ¬x₂ ∨ x₃) ∧ (¬x₁ ∨ x₂ ∨ x₄) ∧ (x₂ ∨ ¬x₃ ∨ ¬x₄) se representa como [[1, −2, 3], [−1, 2, 4], [2, −3, −4]].",
  ),
  P(
    "**Formato DIMACS CNF.** Es el formato de texto estándar que leen los solvers. Las líneas que empiezan con c son comentarios; la línea p cnf n m declara n variables y m cláusulas; cada cláusula es una línea de enteros terminada en 0. La fórmula F anterior es:",
  ),
  COD("c ejemplo del enunciado"),
  COD("p cnf 4 3"),
  COD("1 -2 3 0"),
  COD("-1 2 4 0"),
  COD("2 -3 -4 0"),
  P(
    "Todas las fórmulas de los experimentos se guardaron en este formato en la carpeta instancias/, de modo que se pueden reutilizar con cualquier solver externo.",
    { p: { spacing: { before: 120, after: 120 } } },
  ),
);

// 4
c.push(H1("4. Algoritmo exacto: fuerza bruta y por qué es costoso"));
c.push(
  P(
    "El algoritmo prueba, una por una, las 2ⁿ asignaciones posibles. Cada número entre 0 y 2ⁿ − 1, escrito en binario, representa una asignación: el bit i-ésimo indica el valor de la variable i. Para cada asignación se recorren las cláusulas; si alguna no tiene ningún literal verdadero, la asignación se descarta. Si una asignación cumple todas, se devuelve SAT y el modelo. Si se terminan todas sin éxito, se devuelve UNSAT. Si se supera el límite de tiempo configurado, se devuelve TIMEOUT.",
  ),
  P(
    "Es un algoritmo exacto y completo: siempre da la respuesta correcta, y es el único de los dos capaz de demostrar que una fórmula no tiene solución. Se registran las asignaciones y cláusulas evaluadas y el tiempo.",
  ),
  P(
    "**Complejidad temporal: O(2ⁿ · m · k)** en el peor caso, con n variables, m cláusulas y k literales por cláusula. Hay 2ⁿ asignaciones y, para cada una, hay que evaluar hasta m cláusulas de k literales. **Complejidad espacial:** O(m · k) para guardar la fórmula y O(n) para la asignación en curso.",
  ),
  P(
    "**Por qué es costoso.** El número de asignaciones se duplica con cada variable nueva: con n = 30 son 2³⁰ ≈ 1,07 × 10⁹; con n = 100 serían 2¹⁰⁰ ≈ 1,3 × 10³⁰, imposible de recorrer aunque se evaluara una asignación por nanosegundo. Además, el algoritmo no usa ninguna estructura de la fórmula para descartar grupos de asignaciones; solo corta cada asignación al encontrar la primera cláusula falsa. En una fórmula insatisfacible no hay atajo: hay que recorrer todo el espacio.",
  ),
);

// 5
c.push(H1("5. Algoritmo heurístico: WalkSAT"));
c.push(
  P(
    "WalkSAT (Selman, Kautz y Cohen, 1994) es una búsqueda local estocástica. Parte de una asignación al azar y repite lo siguiente hasta cumplir todas las cláusulas o agotar el límite de pasos:",
  ),
  N("Elegir al azar una cláusula insatisfecha."),
  N(
    "Calcular, para cada variable de esa cláusula, cuántas cláusulas hoy satisfechas se romperían al darla vuelta (break count).",
  ),
  N(
    "Si alguna variable no rompe ninguna cláusula, darla vuelta. Si no, con probabilidad p dar vuelta una variable cualquiera de la cláusula, y con probabilidad 1 − p la que menos rompe.",
  ),
  P(
    "**Parámetros.** Cantidad máxima de pasos (flips), probabilidad de movimiento aleatorio p, cantidad de reinicios, semilla y tiempo máximo. En los experimentos se usó p = 0,5, un reinicio y semillas fijas. La implementación mantiene, para cada cláusula, cuántos literales verdaderos tiene, de modo que cada paso cuesta aproximadamente O(k · c), con c el número medio de cláusulas en que aparece un literal. El espacio es O(m · k).",
    { p: { spacing: { before: 120, after: 120 } } },
  ),
  P(
    "**Ventajas.** Memoria lineal en el tamaño de la fórmula, pasos muy baratos, y en fórmulas satisfacibles grandes suele encontrar una solución en una fracción minúscula del espacio de 2ⁿ asignaciones. Además solo declara SAT cuando, según sus contadores internos, no quedan cláusulas insatisfechas.",
  ),
  P(
    "**Limitaciones.** Es incompleta: si no encuentra una asignación devuelve UNKNOWN, y eso no prueba que la fórmula sea insatisfacible; solo significa que no la encontró dentro de sus límites. Puede quedar atrapada en mínimos locales, su rendimiento depende de p y del tope de pasos, y es aleatoria: dos corridas con distinta semilla pueden diferir.",
  ),
  P(
    "**Alternativas.** GSAT (la versión previa, que elige la variable con mejor ganancia global), Novelty y sus variantes (que agregan memoria sobre qué variable se dio vuelta recientemente), y los solvers completos basados en CDCL (aprendizaje de cláusulas dirigido por conflictos, Marques-Silva y Sakallah, 1999, a partir de la línea Davis-Putnam y DLL), que sí pueden demostrar insatisfacibilidad. Se comparan en la sección 8.4.",
  ),
);

// 6
c.push(H1("6. Complejidad computacional"));
c.push(
  tabla(
    ["", "Fuerza bruta", "WalkSAT"],
    [
      fila([
        "Tiempo",
        "O(2ⁿ · m · k)",
        "O(pasos · k · c), sin garantía de terminar con éxito",
      ]),
      fila(["Espacio", "O(m · k + n)", "O(m · k + n)"]),
      fila([
        "Garantía",
        "Siempre correcto; demuestra SAT y UNSAT",
        "Solo puede confirmar SAT; si falla, UNKNOWN",
      ]),
      fila([
        "Cantidad de recorridos",
        "2ⁿ asignaciones",
        "Depende de la fórmula y del azar",
      ]),
    ],
    [20, 40, 40],
    { izq: [0] },
  ),
);
c.push(espacio());

// 7
c.push(H1("7. Metodología experimental"));
c.push(
  H2("7.1 Hardware y condiciones"),
  P(
    "Todas las corridas del algoritmo exacto y del experimento de escalabilidad se hicieron en un mismo equipo. Los datos están registrados por el propio programa en hardware.txt:",
  ),
  B(
    `**Procesador:** ${hw["CPU"]}, ${hw["Nucleos logicos (os.cpu_count)"]} núcleos lógicos (los algoritmos usan un solo hilo).`,
  ),
  B(`**Memoria RAM:** ${hw["RAM total"].replace(".", ",")}.`),
  B(`**Sistema operativo:** ${hw["SO"]}.`),
  B(
    `**Python:** ${hw["Python"]} (implementación en Python puro, sin extensiones compiladas).`,
  ),
  B(`**Límite de tiempo del exacto (LIMITE):** ${num(LIM)} s por ejecución.`),
  B(`**Inicio de la tanda:** ${hw["Fecha/hora"].replace("T", " ")}.`),
  P(
    "Condiciones controladas: mismo equipo, mismas versiones y mismo LIMITE en todas las corridas, el procedimiento del README pide equipo enchufado con plan de alto rendimiento y sin otros programas abiertos. Los tiempos se midieron con time.perf_counter(), sin incluir la generación de la fórmula. No se repitió cada medición del exacto: cada tiempo de las tablas es una sola corrida, así que pequeñas variaciones del sistema operativo no se promedian.",
  ),
  H2("7.2 Instancias"),
  B(
    "**Aleatorias:** k-SAT aleatorio, cada cláusula elige 3 variables distintas con signo al azar.",
  ),
  B(
    "**Planted (satisfacibles garantizadas):** se elige una solución oculta y cada cláusula se corrige para cumplirla.",
  ),
  B(
    "**Insatisfacibles controladas:** las 8 cláusulas con todas las combinaciones de signos sobre 3 variables (nadie puede cumplirlas a la vez) más cláusulas aleatorias hasta llegar a m.",
  ),
  B(
    "**Razón m/n ≈ 4,26.** En 3-SAT aleatorio la transición entre fórmulas casi siempre satisfacibles y casi siempre insatisfacibles ocurre empíricamente cerca de m/n ≈ 4,26, y ahí las instancias son más difíciles. Se usó m = round(4,26 · n) en la comparación con el exacto.",
  ),
  P(
    "Cada fórmula se guardó en DIMACS CNF en instancias/ con semilla fija (la semilla de la fórmula es n), para que ambos algoritmos resuelvan exactamente la misma fórmula y se pueda repetir todo.",
  ),
  H2("7.3 Diseño de cada experimento"),
  B(
    `**Exacto vs. WalkSAT:** n = ${exacto
      .map((r) => r.n)
      .filter((v, i, a) => a.indexOf(v) === i)
      .join(
        ", ",
      )}, una fórmula planted y una insatisfacible por tamaño. WalkSAT: 10 corridas con semillas 0 a 9 sobre cada fórmula, hasta 100.000 pasos cada una.`,
  ),
  B(
    "**Densidad:** n = 16, razones m/n = 2,0; 3,0; 4,0; 4,26; 5,0 y 6,0, con 30 fórmulas aleatorias por razón (semillas 0 a 29), resueltas con el exacto y con WalkSAT (hasta 20.000 pasos).",
  ),
  B(
    "**Escalabilidad:** WalkSAT con n = 100, 500, 1.000, 5.000 y 10.000, tres semillas por tamaño y por tipo (planted con m/n = 4,26; aleatoria con m/n = 4,0), hasta 5.000.000 de pasos y 120 s por corrida. La semilla de WalkSAT es distinta de la del generador, para que no arranque casualmente en la solución oculta.",
  ),
);

c.push(
  H2("7.4 Declaración de autoría y uso de IA"),
  P(
    "El código (sat.py, experimentos.py, graficos.py, verificar_modelos.py) y los documentos de este trabajo fueron desarrollados con asistencia de una herramienta de inteligencia artificial (Claude, de Anthropic), y revisados, ejecutados y comprendidos por los autores, que son responsables de su contenido. Los resultados provienen de las corridas reales de los programas en el equipo descripto en 7.1.",
  ),
);

// 8.1
const cuadroIns = medidas.map((r) => `n = ${r.n}: ${seg(r.t)}`);
c.push(H1("8. Resultados"));
c.push(H2("8.1 Tiempos del exacto"));
c.push(
  P(
    `La siguiente tabla resume el algoritmo exacto. Con LIMITE = ${num(LIM)} s se midieron todos los tamaños de n = ${exacto[0].n} a n = ${nMax}; el único caso que no terminó es la fórmula insatisfacible de n = ${to32 ? to32.n : "?"}.`,
  ),
);
c.push(
  tabla(
    [
      "n",
      "m",
      "Insat.: estado",
      "Insat.: tiempo",
      "Asignaciones",
      "Plant.: estado",
      "Plant.: tiempo",
      "Asignaciones",
    ],
    [...new Set(exacto.map((r) => r.n))].map((n) => {
      const i = get(ins, n),
        p = get(pla, n);
      return fila(
        [
          n,
          i.m,
          i.estado_exacto,
          esTimeout(i) ? "no terminó en 1 h (TIMEOUT)" : seg(i.t),
          num(i.asig) +
            (esTimeout(i) ? " (de " + num(Math.pow(2, n)) + ")" : ""),
          p.estado_exacto,
          esTimeout(p) ? "no terminó en 1 h (TIMEOUT)" : seg(p.t),
          num(p.asig),
        ],
        esTimeout(i) ? "FDE9E7" : undefined,
      );
    }),
    [6, 6, 11, 17, 18, 11, 15, 16],
    { size: 16 },
  ),
);
c.push(espacio());
c.push(
  ...fig(
    "1_tiempo_fuerza_bruta.png",
    "Gráfico 1. Tiempo de la fuerza bruta según n (escala logarítmica). La cruz roja marca el TIMEOUT.",
  ),
);
c.push(
  P(
    `**Interpretación frente a 2ⁿ.** En las fórmulas insatisfacibles el algoritmo debe recorrer todas las asignaciones: en los casos que terminaron se evaluaron exactamente 2ⁿ (Gráfico 5), y el tiempo es proporcional a 2ⁿ. Una recta en escala logarítmica indica crecimiento exponencial. Si t(n) ≈ k · 2ⁿ, al sumar dos variables el tiempo debería multiplicarse por 2² = 4. Con los datos reales, la razón t(n+2)/t(n) entre insatisfacibles consecutivas es:`,
  ),
);
c.push(
  tabla(
    ["n → n+2", ...razones.map((x) => `${x.n}→${x.n + 2}`)],
    [fila(["t(n+2)/t(n)", ...razones.map((x) => num(x.r, 2))])],
    [14, ...razones.map(() => 8)],
    { size: 16 },
  ),
);
c.push(espacio());
c.push(
  P(
    `Para n ≥ 18 (donde el tiempo ya es mucho mayor que el ruido del reloj) la razón vale entre ${num(razMin, 2)} y ${num(razMax, 2)}, con media geométrica ${num(razMedia, 2)}: muy cerca de 4, como predice 2². Las razones para n pequeño (10 a 16) se apartan más porque los tiempos son de milésimas de segundo y dominan los costos fijos. El ajuste t ≈ k · 2ⁿ con el mayor n medido (n = ${ultimo.n}, ${seg(ultimo.t)}) da k ≈ ${(kAjuste * 1e6).toFixed(2).replace(".", ",")} × 10⁻⁶ s.`,
  ),
);
if (to32) {
  c.push(
    P(
      `**El caso que no terminó.** La insatisfacible de n = ${to32.n} se cortó por el límite de ${num(LIM)} s (tiempo medido hasta el corte: ${seg(to32.t)}), habiendo evaluado ${num(to32.asig)} de las ${num(Math.pow(2, to32.n))} asignaciones (${num(fracTO * 100, 1)} %). Por lo tanto **no hay un tiempo medido para ese caso**: el exacto no terminó en 1 h. Como referencia, extrapolando el ajuste anterior (no medido) se esperaría del orden de ${dur(proy(to32.n))} para completarlo, y con la velocidad observada hasta el corte (${num(tasaTO, 0)} asignaciones por segundo) del orden de ${dur(proyTO)}; ambos valores superan la hora, coherente con el TIMEOUT. Son estimaciones, no resultados.`,
    ),
  );
}
const nonmono = pla
  .filter((r, i) => i > 0 && r.t < pla[i - 1].t && !esTimeout(r))
  .map((r) => r.n);
c.push(
  P(
    `En las fórmulas satisfacibles el tiempo depende de dónde cae la primera solución en el orden de recorrido, por eso es menor y mucho más irregular (por ejemplo, n = ${nonmono.join(", ")} tardaron menos que el tamaño anterior). La tendencia general sigue siendo exponencial: con n = ${nMax} la fórmula planted tardó ${seg(get(pla, nMax).t)} y evaluó ${num(get(pla, nMax).asig)} asignaciones.`,
  ),
);
c.push(
  ...fig(
    "5_asignaciones.png",
    "Gráfico 5. Asignaciones evaluadas según n (escala logarítmica).",
  ),
);

// 8.2
c.push(H2("8.2 Comparación exacto vs. WalkSAT sobre las mismas fórmulas"));
c.push(
  P(
    "La tabla incluye todos los tamaños y las dos clases de fórmula. WalkSAT se corrió 10 veces por fórmula; el tiempo es el promedio por corrida y error_relativo corresponde a la mejor corrida.",
  ),
);
c.push(
  tabla(
    [
      "n",
      "m",
      "Tipo",
      "Estado exacto",
      "t exacto",
      "Éxitos WalkSAT /10",
      "t WalkSAT (medio)",
      "error_ relativo",
      "óptimo alcanzado",
    ],
    exacto.map((r) =>
      fila(
        [
          r.n,
          r.m,
          r.tipo.startsWith("planted") ? "planted (SAT)" : "insatisf.",
          r.estado_exacto,
          esTimeout(r) ? "no terminó en 1 h (TIMEOUT)" : seg(r.t),
          r.exitos,
          seg(r.tw),
          num(r.err, 4),
          r.optimo_alcanzado,
        ],
        esTimeout(r)
          ? "FDE9E7"
          : r.tipo.startsWith("planted")
            ? undefined
            : "F4F7FB",
      ),
    ),
    [5, 5, 12, 10, 16, 10, 13, 10, 10],
    { size: 15 },
  ),
);
c.push(espacio());
c.push(
  ...fig(
    "3_comparacion_tiempos.png",
    "Gráfico 3. Fuerza bruta y WalkSAT sobre las mismas fórmulas (escala logarítmica).",
  ),
);
c.push(
  ...fig(
    "4_exito_walksat.png",
    "Gráfico 4. Porcentaje de corridas de WalkSAT que encontraron solución según n (ver sección 8.4).",
  ),
);
const maxTw = Math.max(...pla.map((r) => r.tw));
const insTw = ins.map((r) => r.tw);
c.push(
  P(
    `**Fórmulas satisfacibles.** WalkSAT encontró una solución en las 10 corridas de los ${pla.length} tamaños (${pla.reduce((s, r) => s + r.exitos, 0)} de ${pla.length * 10} corridas), con un tiempo medio máximo de ${seg(maxTw)}. El exacto pasó de ${seg(pla[0].t)} (n = ${pla[0].n}) a ${seg(get(pla, nMax).t)} (n = ${nMax}).`,
  ),
  P(
    `**Fórmulas insatisfacibles.** Solo el exacto lo demostró (UNSAT) en los tamaños que terminó. WalkSAT informó UNKNOWN en todas las corridas, consumió los 100.000 pasos (entre ${seg(Math.min(...insTw))} y ${seg(Math.max(...insTw))} por corrida) y su tiempo no mide nada útil: no puede concluir. Para n = ${to32 ? to32.n : "?"}, donde el exacto no terminó, WalkSAT tampoco puede decir nada: la fórmula es insatisfacible por construcción, pero eso lo sabemos por cómo se generó, no porque algún algoritmo lo haya establecido.`,
  ),
  P(
    "**Por qué el error relativo se define como fracción de cláusulas insatisfechas.** En un problema de optimización (por ejemplo, el viajante) el error relativo compara el valor obtenido con el óptimo: (valor − óptimo) / óptimo. SAT es un problema de decisión: la respuesta es sí o no, y una asignación que cumple el 99 % de las cláusulas sigue siendo una respuesta incorrecta. Para tener una medida comparable se mira el problema asociado, Max-SAT, cuyo óptimo es la mayor cantidad de cláusulas cumplibles. Se define entonces error_relativo = (cláusulas insatisfechas de la mejor asignación de WalkSAT) / m. Vale 0 si cumplió todas (óptimo alcanzado en una fórmula satisfacible) y crece a medida que se aleja. Es una medida de cercanía a la solución, no de correctitud.",
  ),
  P(
    "**Qué significa n.a.** La columna óptimo_alcanzado dice sí o no cuando el óptimo Max-SAT es conocido porque el exacto encontró una solución (óptimo = 0 cláusulas insatisfechas). Dice n.a. (no aplica) cuando ese óptimo no se conoce con los datos del experimento: en las insatisfacibles, porque el exacto solo establece que no hay asignación perfecta pero no calcula cuántas cláusulas mínimas quedan sin cumplir; y en el TIMEOUT, porque el exacto no respondió.",
  ),
);
const todasUnaClausula = ins.every((r) => Math.abs(r.err * r.m - 1) < 0.02);
if (todasUnaClausula) {
  c.push(
    P(
      `**Una observación derivada (no medida por el exacto).** En todas las insatisfacibles el error_relativo es exactamente 1/m (por ejemplo, ${num(get(ins, 10).err, 5)} = 1/${get(ins, 10).m} para n = 10): la mejor asignación de WalkSAT deja una sola cláusula sin cumplir. Como las 8 cláusulas con todas las combinaciones de signos sobre 3 variables no pueden cumplirse a la vez, cualquier asignación incumple al menos una cláusula; por lo tanto 1 es una cota inferior y WalkSAT habría alcanzado el óptimo Max-SAT en esas fórmulas. Es un razonamiento a partir de la construcción de las instancias; el programa deja n.a. porque no lo verifica.`,
    ),
  );
}
c.push(
  P(
    `Advertencia sobre las fórmulas planted: resultaron fáciles para WalkSAT (entre ${num(Math.min(...pla.map((r) => r.flips)), 1)} y ${num(Math.max(...pla.map((r) => r.flips)), 1)} pasos medios). Para no sacar conclusiones solo con casos triviales se agregaron los experimentos de las secciones 8.3 y 8.4.`,
  ),
);

// 8.3
c.push(H2("8.3 Dificultad variable: efecto de la razón m/n (n = 16)"));
c.push(
  P(
    "Se generaron 30 fórmulas aleatorias por cada razón m/n y se las resolvió con el exacto y con WalkSAT (hasta 20.000 pasos).",
  ),
);
c.push(
  tabla(
    [
      "m/n",
      "m",
      "Satisfacibles según el exacto (de 30)",
      "WalkSAT encontró solución",
      "WalkSAT dijo SAT en una UNSAT",
    ],
    densidad.map((r) =>
      fila([
        r.razon_m_n.replace(".", ","),
        r.m,
        r.satisfacibles_segun_exacto,
        r.walksat_encontro_solucion,
        r.walksat_dijo_SAT_en_UNSAT,
      ]),
    ),
    [10, 10, 30, 25, 25],
  ),
);
c.push(espacio());
const d0 = densidad[0],
  dN = densidad[densidad.length - 1];
const igual = densidad.every(
  (r) => r.satisfacibles_segun_exacto === r.walksat_encontro_solucion,
);
c.push(
  P(
    `Al aumentar la cantidad de cláusulas por variable, la fracción de fórmulas satisfacibles cae (de ${d0.satisfacibles_segun_exacto} de 30 a ${dN.satisfacibles_segun_exacto} de 30). ${igual ? "En todos los casos WalkSAT encontró solución exactamente cuando existía" : "WalkSAT no coincidió con el exacto en algún caso"}, y nunca afirmó SAT sobre una fórmula insatisfacible, como debe ser porque solo declara SAT cuando no quedan cláusulas insatisfechas según sus contadores internos. Cerca de m/n = 4,26 la mezcla de satisfacibles e insatisfacibles es la más ambigua (${densidad.find((r) => r.razon_m_n === "4.26").satisfacibles_segun_exacto} de 30). Este resultado vale para n = 16 y debe leerse con cuidado: con más variables las proporciones cambian y la transición se hace más abrupta. Estas 30 × 6 corridas son conteos, no tiempos, y se obtienen con semillas fijas, por lo que no dependen del hardware.`,
  ),
);

// 8.4
c.push(
  H2(
    `8.4 Escalabilidad de WalkSAT (n = 100 a ${num(nEsc)}) y viabilidad en casos reales`,
  ),
);
const filasEsc = [...new Set(grupos.map((g) => g.n))].map((n) => {
  const p = grupos.find((g) => g.n === n && g.tipo === "planted"),
    a = grupos.find((g) => g.n === n && g.tipo === "aleatoria");
  return fila(
    [
      num(n),
      num(p.m),
      `${p.k}/${p.tot}`,
      num(p.flips),
      seg(p.t),
      num(a.m),
      `${a.k}/${a.tot}`,
      num(a.flips),
      seg(a.t),
    ],
    p.k < p.tot || a.k < a.tot ? "FFF4D6" : undefined,
  );
});
c.push(
  tabla(
    [
      "n",
      "m (planted)",
      "Planted: éxito",
      "Pasos medios",
      "Tiempo medio",
      "m (aleatoria)",
      "Aleatoria: éxito",
      "Pasos medios",
      "Tiempo medio",
    ],
    filasEsc,
    [8, 10, 10, 13, 12, 11, 11, 13, 12],
    { size: 16 },
  ),
);
c.push(espacio());
c.push(
  ...fig(
    "2_tiempo_walksat.png",
    "Gráfico 2. Tiempo medio de WalkSAT según n (escala logarítmica).",
  ),
);
const g5r = grupos.find((g) => g.n === 5000 && g.tipo === "aleatoria"),
  g5p = grupos.find((g) => g.n === 5000 && g.tipo === "planted");
const sinTot = g10
  .map((g) => `${g.tipo}: entre ${g.sinMin} y ${g.sinMax}`)
  .join("; ");
c.push(
  P(
    `Hasta n = 1.000 WalkSAT resolvió todas las fórmulas en menos de un segundo. Con n = 5.000 las planted tardaron ${seg(g5p.t)} de media (${g5p.k}/${g5p.tot} resueltas) y las aleatorias resolvieron ${g5r.k} de ${g5r.tot}.`,
  ),
  fallo10
    ? P(
        `**n = ${num(nEsc)}: no se resolvió ninguna fórmula (0 de ${g10.reduce((s, g) => s + g.tot, 0)}).** Todas las corridas terminaron en UNKNOWN al llegar al tope de ${num(maxFlipsEsc)} pasos (unos ${seg(tMaxEsc)} como máximo; el límite de 120 s por corrida no se alcanzó, el que frenó fue el de pasos). Dejaron muy pocas cláusulas sin cumplir en la mejor asignación (${sinTot}; aproximado a partir del porcentaje registrado con 3 decimales). Eso no significa que las fórmulas sean insatisfacibles ni que WalkSAT no pudiera resolverlas con más pasos o con otros parámetros: es el límite de este experimento.`,
      )
    : P(`Con n = ${num(nEsc)} hubo corridas exitosas; ver tabla.`),
  P(
    "Las fórmulas planted, aun con más cláusulas por variable (4,26), fueron más fáciles que las aleatorias de razón 4,0 hasta n = 5.000: la dificultad depende de la estructura y no solo de n o de m/n. No se estudió de forma sistemática el efecto del parámetro p ni de los reinicios; queda como trabajo futuro.",
  ),
);
c.push(
  ...fig(
    "6_insatisfechas_por_iteracion.png",
    `Gráfico 6. Cláusulas insatisfechas por iteración (n = 500, m = 2.000, aleatoria; corrida de ${num(hist.length)} pasos).`,
  ),
);
const primeroMenor = (u) => hist.findIndex((v) => v < u) + 1;
c.push(
  P(
    `El Gráfico 6 muestra el comportamiento típico: comienza con ${hist[0]} cláusulas insatisfechas, baja a ${hist[99]} en el paso 100 y a ${hist[999]} en el paso 1.000, y después oscila cerca de cero (primera vez por debajo de 10 en el paso ${primeroMenor(10)}) hasta encontrar la solución en el paso ${num(hist.length)} de esa corrida.`,
  ),
);

c.push(H2("Discusión ampliada: viabilidad en casos reales"));
c.push(
  P(
    "Los resultados anteriores son sobre fórmulas generadas por nosotros. Conviene analizar con cuidado qué se puede y qué no se puede trasladar a problemas reales.",
  ),
  B(
    "**WalkSAT frente a los solvers CDCL modernos.** Los solvers de uso práctico (MiniSat, Glucose, Kissat y similares) son completos: aprenden cláusulas a partir de los conflictos, reinician y usan heurísticas de decisión, y pueden responder SAT o UNSAT. En la SAT Competition 2024, Kissat ganó las tres categorías de la pista principal. En este trabajo no se ejecutó ningún solver CDCL, por lo que no hay medición propia que compare ambos enfoques; la comparación de este informe es entre fuerza bruta y WalkSAT, y la fuerza bruta es una referencia deliberadamente ingenua.",
  ),
  B(
    "**Instancias aleatorias frente a estructuradas.** Nuestras fórmulas son aleatorias o planted. Las instancias que surgen de aplicaciones (verificación, planificación, horarios) tienen estructura: variables fuertemente relacionadas, cláusulas cortas de restricciones, simetrías. En esa clase, la bibliografía del área indica que los solvers CDCL suelen ser más efectivos que la búsqueda local, mientras que en k-SAT aleatorio cerca del umbral la búsqueda local es competitiva. Nuestros datos no permiten probar esa afirmación: solo muestran el comportamiento de WalkSAT en instancias aleatorias.",
  ),
  B(
    "**Incompletitud: UNKNOWN no es UNSAT.** Un horario sin conflictos, un plan o una verificación a veces requieren saber que no existe solución (por ejemplo, que un conjunto de restricciones es contradictorio). WalkSAT no puede probarlo: si devuelve UNKNOWN solo dice que no encontró nada. Para eso se necesita un algoritmo completo.",
  ),
  B(
    `**Tiempo y pasos como parámetros.** Que n = ${num(nEsc)} haya fallado se debe en parte a un tope de ${num(maxFlipsEsc)} pasos que fijamos nosotros. El resultado depende de ese tope, de p y de la semilla, así que no es una propiedad de WalkSAT sola. En una aplicación real hay que decidir cuánto esperar y qué hacer ante un UNKNOWN (más pasos, reinicios, otra semilla, o pasar a un solver completo).`,
  ),
  B(
    "**Límites de generalizar con 3 semillas.** En la escalabilidad se usaron 3 fórmulas por tamaño y por tipo. Con tan pocas observaciones, un “1 de 3” no permite estimar una probabilidad de éxito con precisión, y una diferencia entre planted y aleatoria puede ser parcialmente casualidad. Las conclusiones se leen como tendencias en este rango, no como leyes.",
  ),
  P(
    "En síntesis: WalkSAT es viable cuando alcanza con encontrar una solución, la instancia es satisfacible y el costo de un fallo es bajo; para el resto de los casos reales, lo habitual es usar un solver completo, a veces combinado con búsqueda local.",
  ),
);

// 9
c.push(H1("9. Comparación exacto vs. heurístico"));
c.push(
  P(
    `**Costo.** El exacto crece exponencialmente con n: el tiempo se multiplicó por ${num(razMedia, 1)} cada dos variables (n ≥ 18), llegó a ${seg(ultimo.t)} con n = ${ultimo.n} en una insatisfacible y con n = ${to32 ? to32.n : "?"} no terminó en 1 h (TIMEOUT). En las mismas fórmulas satisfacibles WalkSAT tardó como máximo ${seg(maxTw)}, y resolvió fórmulas planted de hasta n = 5.000 en unos ${seg(g5p.t)}.`,
  ),
  P(
    "**Garantías.** El exacto siempre responde correctamente y es el único que demuestra UNSAT. WalkSAT solo puede confirmar SAT; ante un fallo no puede concluir nada (UNKNOWN).",
  ),
  P(
    "**Calidad.** En las fórmulas satisfacibles WalkSAT alcanzó el óptimo (error_relativo = 0) en todas las corridas. En las insatisfacibles, su mejor asignación dejó una sola cláusula sin cumplir en todos los tamaños, incluido el que el exacto no pudo terminar.",
  ),
  P(
    "**Cuándo conviene cada uno.** El exacto para n pequeño o cuando hay que demostrar que no hay solución; WalkSAT cuando n es grande y basta con encontrar una asignación. En la práctica ambos son superados por solvers CDCL, que combinan completitud con buen desempeño en instancias estructuradas.",
  ),
  P(
    "**Diferencia con optimización.** En SAT la respuesta es sí o no, por eso el error de WalkSAT no se mide como un porcentaje de error clásico sino por si encontró la solución o no y, si no, qué fracción de cláusulas dejó sin cumplir (sección 8.2).",
  ),
);

// 10
c.push(H1("10. Aplicación en un contexto real"));
c.push(
  NOTA(
    "Rótulos usados en esta sección: [verificado] = se leyó la fuente citada; [no verificado] = dato que no se pudo comprobar; [hipotético, no documentado] = modelo propio sin evidencia de uso real.",
  ),
  H2(
    "10.1 Un caso real argentino contemporáneo: horarios en la Universidad Torcuato Di Tella (usa ILP, no SAT)",
  ),
  P(
    "Ferreiro y Marenco (2024, según metadatos del PDF) presentan un enfoque de programación entera para la programación de horarios de clases en la Universidad Torcuato Di Tella (Memorias de las 53 JAIIO, SIIIO, ISSN 2451-7496, pp. 375-382). [verificado por búsqueda: metadatos y resumen]",
  ),
  P(
    "**El problema.** Asignar un módulo horario (semana de 5 días por 7 bandas, 35 módulos) y un aula a cada clase de cada comisión, minimizando los tiempos muertos de docentes y estudiantes, con materias compartidas entre carreras. Restricciones duras: una comisión no repite módulo; un aula no se usa dos veces en el mismo módulo; disponibilidad de docentes y comisiones; clases consecutivas o con separación mínima; y cada paquete de estudiantes tiene una comisión por materia sin superposición. Restricciones elásticas (penalizadas): máximo diario por docente, cupo de aula y máximo diario por paquete.",
  ),
  P(
    "**El método.** Dos modelos de programación lineal entera (ILP) en secuencia: la Fase I reparte estudiantes entre comisiones; la Fase II elige un esquema de horario por comisión (variable binaria, exactamente un esquema por comisión), más un esquema secuencial por sub-instancias.",
  ),
  P(
    "**El impacto reportado** es solo cualitativo: soluciones de buena calidad en tiempos de ejecución aceptables, y se avanza en implementarlo con una herramienta de visualización y edición. [no verificado: solver usado, tamaños de instancia, tiempos y ahorros; el resumen consultado no los informa y este informe no los inventa.]",
  ),
  P(
    "**Aclaración importante: este caso usa ILP, no SAT.** Se lo presenta como pariente cercano porque el mismo problema se puede codificar en CNF: las restricciones “a lo sumo uno” y “exactamente uno” (un aula por clase y módulo, un esquema por comisión) son cláusulas, como se muestra en 10.3. SAT, ILP y coloreo de grafos son NP-completos y se reducen entre sí, así que la dificultad teórica estudiada en este trabajo es la misma.",
  ),
  H2("10.2 El puente hacia SAT: asignación de aulas como coloreo de grafos"),
  P(
    "Escalante, Chiesa y Fekete (SUMA 2019, FCEIA-UNR) plantean el problema de asignación de aulas en una institución con dos sedes como coloreo de grafos por listas, resuelto con programación lineal entera sobre datos reales de 2018. [verificado: resumen de una página, sin cifras]. Aunque usan ILP, el k-coloreo se codifica directamente en CNF:",
  ),
  B(
    "Variable x(v,c): el vértice v (una clase) recibe el color c (un módulo o un aula).",
  ),
  B(
    "Al menos un color por vértice: para cada v, (x(v,1) ∨ x(v,2) ∨ … ∨ x(v,k)).",
  ),
  B(
    "A lo sumo un color por vértice: para cada v y cada par de colores c < d, (¬x(v,c) ∨ ¬x(v,d)).",
  ),
  B(
    "Adyacentes no comparten color: para cada arista (u,v) y cada color c, (¬x(u,c) ∨ ¬x(v,c)).",
  ),
  P(
    "La fórmula resultante es satisfacible si y solo si el grafo se puede colorear con k colores. Con la lista de colores permitidos de cada vértice (por ejemplo, aulas con capacidad suficiente) basta agregar cláusulas unitarias (¬x(v,c)) para los colores no permitidos.",
  ),
  H2("10.3 Modelo propio para la FRCU (hipotético, no documentado)"),
  P(
    "[hipotético, no documentado] No se encontró evidencia pública de que la UTN Facultad Regional Concepción del Uruguay ni otra institución de Concepción del Uruguay o de Entre Ríos use SAT para armar horarios. Lo que sigue es un modelo propio, sin datos reales de la institución. Se asume que la FRCU publica los horarios de cursada como documentos PDF con comisiones, día y franja [VERIFICAR: comprobar esa afirmación y citar la fuente o URL antes de entregar].",
  ),
  P(
    "Variable x(c,f): la comisión c se dicta en la franja f. Restricciones como cláusulas:",
  ),
  B("Cada comisión tiene al menos una franja: (x(c,1) ∨ x(c,2) ∨ …)."),
  B(
    "Cada comisión tiene a lo sumo una franja: (¬x(c,f) ∨ ¬x(c,g)) para cada par de franjas.",
  ),
  B(
    "Conflictos de docente o de alumnos: dos comisiones en conflicto no pueden estar en la misma franja: (¬x(c₁,f) ∨ ¬x(c₂,f)).",
  ),
  B(
    "Disponibilidad: si un docente no puede en una franja, la cláusula unitaria (¬x(c,f)).",
  ),
  B(
    "Aulas y capacidad: se modelan agregando variables para el aula y cláusulas análogas. No se implementó.",
  ),
  P(
    "**Ejemplo de juguete (hipotético, no es el horario real).** 3 comisiones y 3 franjas, es decir 9 variables y 19 cláusulas (3 de “al menos una”, 9 de “a lo sumo una”, 6 de conflicto y 1 unitaria). C1 y C2 comparten docente, C2 y C3 comparten alumnos y el docente de C3 no puede en la franja 1. Numerando x(c,f) = 3(c−1) + f, el algoritmo exacto devolvió SAT tras evaluar 139 asignaciones, con C1 en la franja 2, C2 en la 1 y C3 en la 2. WalkSAT (semilla 1) devolvió SAT en 4 pasos, con C1 en la franja 1, C2 en la 2 y C3 en la 3. Ambas soluciones cumplen todas las cláusulas. (Estas corridas se reproducen con el script ejemplo_horarios.py, que usa sat.py.) La fórmula no es 3-SAT puro (tiene cláusulas de 1, 2 y 3 literales), pero se puede convertir a 3-SAT con el procedimiento de la sección 2.5.",
  ),
  H2("10.4 Impacto económico y logístico"),
  P(
    "**No existe, hasta donde se pudo verificar, una cifra pública atribuible específicamente al uso de SAT** para horarios o asignación de aulas. Este informe no inventa una.",
  ),
  P(
    "Como analogía medida (investigación operativa, no SAT): Kroon, Huisman y colaboradores (2009) describen el nuevo horario de los ferrocarriles holandeses, desarrollado con técnicas de investigación operativa, con un beneficio anual declarado de 50 millones de euros (Interfaces 39(1):6-17). [no verificado en el texto: la cifra se tomó del resumen que muestra un buscador]. Sirve para dimensionar que un buen modelo de horarios puede tener valor económico, no para atribuirlo a SAT.",
  ),
  P(
    "Para la FRCU, cualquier estimación de ahorro sería **hipotética**; no se propone ningún número. Lo que sí se puede plantear son indicadores a medir antes y después de una eventual herramienta: cantidad de conflictos de horario evitados, huecos de los docentes (tiempos muertos), ocupación de las aulas y cantidad de casos en que no existe una solución que cumpla todas las restricciones.",
  ),
  H2("10.5 Herramientas y usos documentados en la práctica"),
  B(
    "**Solvers CDCL:** MiniSat (Eén y Sörensson, 2003), Glucose (Audemard y Simon, IJCAI 2009), Kissat (ganó las tres categorías de la pista principal de la SAT Competition 2024) y el solver SMT Z3 (de Moura y Bjørner, 2008). OR-Tools CP-SAT (Google) es un solver de programación por restricciones que se apoya en técnicas de SAT y es de acceso libre.",
  ),
  B(
    "**Bounded model checking:** Biere, Cimatti, Clarke y Zhu (TACAS 1999) lo plantearon con solvers SAT; el trabajo recibió el ETAPS Test of Time Award 2017.",
  ),
  B(
    "**Dependencias de software:** Eclipse usó Sat4j para resolver dependencias (Le Berre y Parrain, JSAT 7, 2010); libsolv, la biblioteca de resolución de dependencias de openSUSE, usa un solver SAT.",
  ),
  B(
    "**Planificación:** Planning as Satisfiability (Kautz y Selman, ECAI 1992).",
  ),
  B(
    "**Horarios:** Asín Achá y Nieuwenhuis (2014) resuelven horarios de cursada (curriculum-based course timetabling) con SAT y MaxSAT.",
  ),
  P(
    "No se afirma en este informe que empresas específicas usen SAT internamente, porque no se verificó.",
  ),
);

// 11
c.push(H1("11. Reflexión crítica"));
c.push(
  H2("11.1 Limitaciones del experimento"),
  B(
    "Una sola fórmula satisfacible y una insatisfacible por tamaño en la comparación con el exacto, y una sola corrida de tiempo por fórmula. Los tiempos son ilustrativos de la tendencia, no estimaciones con intervalo de confianza.",
  ),
  B(
    "Las fórmulas planted son fáciles para WalkSAT; las insatisfacibles controladas contienen un núcleo pequeño insatisfacible, que no es la estructura típica de una fórmula difícil ni de una aplicación real.",
  ),
  B(
    "En la escalabilidad hay solo 3 semillas por tamaño y por tipo; el tope de 5.000.000 de pasos explica los fallos con n = 10.000 y fue fijado por nosotros.",
  ),
  B(
    "Se usó una única configuración de WalkSAT (p = 0,5, un reinicio). No se exploró p ni los reinicios.",
  ),
  B(
    "Todo está en Python puro: un lenguaje compilado llegaría a más variables con la misma curva exponencial, y los valores absolutos de tiempo no son comparables con solvers optimizados.",
  ),
  B(
    "No se corrió ningún solver CDCL ni benchmarks externos, aunque el código puede leer y escribir DIMACS.",
  ),
  H2("11.2 Validez"),
  P(
    "Validez interna: ambos algoritmos resuelven las mismas fórmulas (semillas y archivos DIMACS fijos), y WalkSAT solo declara SAT cuando no quedan cláusulas insatisfechas según sus contadores internos (no vuelve a recorrer la fórmula); como control aparte, los 120 modelos devueltos en el experimento exacto (12 fórmulas planted por 10 semillas) se volvieron a verificar con contar_insatisfechas (script verificar_modelos.py) y todos dieron 0 cláusulas insatisfechas. No produjo falsos positivos en ninguna de las 180 fórmulas de la prueba de densidad. Validez externa: acotada, por el tipo de instancias, el tamaño de la muestra y el hardware. Un tiempo medido es válido para esta máquina, esta versión de Python y esta carga del sistema; lo que se generaliza es la forma de la curva (exponencial en el exacto), no los valores.",
  ),
  H2("11.3 Qué cambiaríamos"),
  B(
    "Repetir cada medición del exacto varias veces y usar más fórmulas por tamaño para reportar dispersión.",
  ),
  B(
    "Agregar un solver CDCL (por ejemplo, MiniSat o Kissat sobre los archivos DIMACS ya generados) como tercera referencia, y probar instancias estructuradas, como la codificación de coloreo o de horarios de la sección 10.",
  ),
  B(
    "Barrer p y los reinicios, y fijar el tope de pasos con un criterio explícito en lugar de un valor único.",
  ),
  B(
    "Medir en n = 10.000 con topes mayores para distinguir “difícil” de “imposible con este tope”.",
  ),
  H2("11.4 Riesgos de usar una heurística incompleta en aplicaciones críticas"),
  P(
    "En verificación, seguridad o sistemas donde un error tiene consecuencias graves, un UNKNOWN no puede tratarse como “no hay problema”. Si una herramienta busca contraejemplos con una heurística incompleta y no los encuentra, no se ha demostrado su ausencia. Los riesgos concretos son: tomar un UNKNOWN como UNSAT, confiar en que “la última vez resolvió” cuando el resultado depende de la semilla y de los límites, y no registrar con qué parámetros se obtuvo un resultado. Para esos usos corresponde un algoritmo completo, idealmente con certificados de insatisfacibilidad verificables por separado.",
  ),
  H2("11.5 Honestidad sobre lo hipotético"),
  P(
    "El modelo de horarios para la FRCU es una construcción propia: no hay evidencia de que la institución lo use ni datos reales para evaluarlo. El caso de la Di Tella es real pero usa ILP, y de él solo se conoce el impacto de forma cualitativa. No hay ninguna cifra económica atribuible a SAT en este trabajo, y la cifra de los ferrocarriles holandeses es una analogía de investigación operativa, tomada de un resumen. Cada afirmación de la sección 10 está rotulada según su grado de verificación.",
  ),
  H2("11.6 Ética y reproducibilidad"),
  P(
    "Todos los números de este informe se generan leyendo los archivos de resultados que produjo el código, y el propio informe se regenera con un script (generadores/informe.js), de modo que no hay cifras copiadas a mano. Las semillas son fijas, las fórmulas están en instancias/ y el hardware se registra en hardware.txt, de modo que otra persona puede repetir las corridas; los tiempos absolutos cambiarán con el equipo. Los casos negativos (el TIMEOUT del exacto en n = 32 y los fallos de WalkSAT en n = 10.000) se informan tal cual, sin sustituirlos por estimaciones. Las referencias se limitan a fuentes cuyos metadatos se comprobaron con un buscador, y se marcó lo que no se verificó en el texto original.",
  ),
);

// 12
c.push(H1("12. Conclusiones"));
c.push(
  P(
    `**Exacto.** La fuerza bruta es correcta pero inviable al crecer n: en las fórmulas insatisfacibles el tiempo se multiplicó por ${num(razMedia, 1)} cada dos variables (el 4 que predice 2²), llegó a ${seg(ultimo.t)} con n = ${ultimo.n} y con n = ${to32 ? to32.n : "?"} no terminó en 1 h (TIMEOUT). Sigue siendo el único de los dos algoritmos que demuestra insatisfacibilidad.`,
  ),
  P(
    `**Heurístico.** WalkSAT resolvió todas las fórmulas satisfacibles de la comparación con un tiempo medio máximo de ${seg(maxTw)} y fórmulas de 5.000 variables en segundos, a cambio de perder garantías: si no encuentra una solución no puede probar que no exista, y con n = ${num(nEsc)} no la encontró dentro del tope de pasos fijado.`,
  ),
  P(
    "**Viabilidad.** La elección depende de si se necesita una demostración de insatisfacibilidad o alcanza con encontrar una solución, de cuánto cuesta un UNKNOWN y de la estructura de la instancia. Para casos reales, lo habitual es recurrir a solvers CDCL, que no se evaluaron aquí. La aplicación a horarios es plausible como modelo (hay casos reales cercanos con ILP y trabajos académicos con SAT y MaxSAT), pero no se verificó que se use en la FRCU.",
  ),
);

// Referencias
c.push(H1("Referencias"));
const refs = [
  "Asín Achá, R. y Nieuwenhuis, R. (2014). Curriculum-based course timetabling with SAT and MaxSAT. Annals of Operations Research, 218(1), 71-91. DOI 10.1007/s10479-012-1081-x.",
  "Audemard, G. y Simon, L. (2009). Predicting learnt clauses quality in modern SAT solvers. Proceedings of IJCAI 2009.",
  "Biere, A., Cimatti, A., Clarke, E. M. y Zhu, Y. (1999). Symbolic model checking without BDDs. TACAS 1999. (ETAPS Test of Time Award 2017: https://www.etaps.org/2017/test-of-time-award).",
  "Biere, A., Heule, M., van Maaren, H. y Walsh, T. (eds.) (2021). Handbook of Satisfiability, 2.ª ed. IOS Press.",
  "Cook, S. A. (1971). The complexity of theorem-proving procedures. Proceedings of the 3rd ACM Symposium on Theory of Computing (STOC), 151-158. DOI 10.1145/800157.805047.",
  "Davis, M. y Putnam, H. (1960). A computing procedure for quantification theory. Journal of the ACM, 7(3), 201-215.",
  "Davis, M., Logemann, G. y Loveland, D. (1962). A machine program for theorem-proving. Communications of the ACM, 5(7), 394-397.",
  "de Moura, L. y Bjørner, N. (2008). Z3: an efficient SMT solver. TACAS 2008.",
  "Eén, N. y Sörensson, N. (2003). An extensible SAT-solver. SAT 2003.",
  "Escalante, M. S., Chiesa, M. y Fekete, P. (2019). Problema de asignación de aulas en institución con dos sedes. SUMA 2019, FCEIA-UNR. https://www.union-matematica.org.ar/suma2019/files/10-8-025.pdf",
  "Ferreiro, C. y Marenco, J. (2024, según metadatos del PDF). Un enfoque de programación entera para la programación de horarios de clases en la Universidad Torcuato Di Tella. Memorias de las 53 JAIIO (SIIIO), ISSN 2451-7496, 375-382. https://revistas.unlp.edu.ar/JAIIO/article/download/18019/17687/81231",
  "Google. OR-Tools: CP-SAT solver. https://developers.google.com/optimization/cp/cp_solver",
  "Karp, R. M. (1972). Reducibility among combinatorial problems. En R. E. Miller y J. W. Thatcher (eds.), Complexity of Computer Computations, 85-103. Plenum Press. DOI 10.1007/978-1-4684-2001-2_9.",
  "Kautz, H. y Selman, B. (1992). Planning as satisfiability. ECAI 1992.",
  "Kroon, L. G., Huisman, D. et al. (2009). The new Dutch timetable: the OR revolution. Interfaces, 39(1), 6-17. https://ideas.repec.org/a/inm/orinte/v39y2009i1p6-17.html",
  "Le Berre, D. y Parrain, A. (2010). The Sat4j library, release 2.2. Journal on Satisfiability, Boolean Modeling and Computation (JSAT), 7.",
  "Levin, L. A. (1973). Universal sequential search problems. Problemy Peredachi Informatsii, 9(3), 115-116.",
  "libsolv (openSUSE). https://github.com/openSUSE/libsolv",
  "Marques-Silva, J. P. y Sakallah, K. A. (1999). GRASP: a search algorithm for propositional satisfiability. IEEE Transactions on Computers, 48(5), 506-521.",
  "SAT Competition 2024, medallas: https://cca.informatik.uni-freiburg.de/sat24medals",
  "Selman, B., Kautz, H. A. y Cohen, B. (1994). Noise strategies for improving local search. Proceedings of AAAI-94, 337-343.",
];
refs.forEach((r) => c.push(REF(r)));

// ---------------------------------------------------------------- documento
const doc = new Document({
  creator: "TP SAT",
  title: "SAT (Problema de Satisfacción Booleana): algoritmo exacto frente a la heurística WalkSAT",
  styles: {
    default: { document: { run: { font: FUENTE, size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 32, bold: true, font: FUENTE, color: "1F4E79" },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 26, bold: true, font: FUENTE, color: "2E75B6" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "viñetas",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 540, hanging: 270 } } },
          },
        ],
      },
      {
        reference: "numeros",
        levels: [
          {
            level: 0,
            format: LevelFormat.DECIMAL,
            text: "%1.",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 540, hanging: 360 } } },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  children: ["Página ", PageNumber.CURRENT],
                  size: 18,
                  color: "7F7F7F",
                }),
              ],
            }),
          ],
        }),
      },
      children: c,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  const salida = path.join(RAIZ, "SAT_Informe.docx");
  fs.writeFileSync(salida, buf);
  console.log("Escrito", salida, buf.length, "bytes");
  if (faltantes.length)
    console.warn("Gráficos faltantes:", faltantes.join(", "));
});
