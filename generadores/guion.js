// Genera SAT_Guion_Exposicion.docx con los datos reales de los CSV. Uso: node generadores/guion.js
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
  ShadingType,
  BorderStyle,
} = require("docx");
const D = require("./datos");
const { fn, ft, hw, razon, unsat, planted, porN } = D;

const AZUL = "1F4E79",
  NARANJA = "C55A11",
  GRIS = "666666",
  OSCURO = "16222E";
const A = "Camilo Escar",
  B = "Facundo Iglesias";

// ---- cifras reales
const u = (n) => porN(unsat, n),
  p = (n) => porN(planted, n);
const min1 = (s) => fn(s / 60, 1);
const n30 = u(30),
  n32u = u(32),
  n32p = p(32);
const todas32 = 2 ** 32;
const pctCorte = fn((n32u.asignaciones / todas32) * 100, 0);
const proyeccion32h = fn((n30.t_exacto_s * razon.gm) / 3600, 1); // proyeccion, no medicion
const nMax = D.nMax,
  e5000 = D.escala.find((e) => e.n === 5000);
const d = D.densidad;
const err = (x) => fn(x * 100, 1) + " %";
const walkUnsatMedio =
  D.walkUnsatT.reduce((a, b) => a + b) / D.walkUnsatT.length;
const ratioTxt = `${fn(razon.gm, 1)} (de ${fn(razon.min, 1)} a ${fn(razon.max, 1)} entre n=${razon.desde} y n=${razon.hasta})`;

// ---- horarios (misma distribucion que presentacion.js)
const T = [
  [1, "Portada", A, "0:00", "0:40"],
  [2, "El problema SAT y por qué importa", A, "0:40", "2:00"],
  [3, "Algoritmo exacto: fuerza bruta", A, "2:00", "3:10"],
  [4, "Equipo y metodología", A, "3:10", "4:00"],
  [5, "Resultado: tiempo de la fuerza bruta", A, "4:00", "5:50"],
  [6, "Heurística: WalkSAT", A, "5:50", "6:50"],
  [7, "Exacto vs. WalkSAT: mismas fórmulas", A, "6:50", "8:20"],
  [8, "Éxitos, error relativo y tiempos", A, "8:20", "10:00"],
  [9, "Escalabilidad de WalkSAT", B, "10:00", "11:10"],
  [10, "Aplicación real: horarios en Argentina", B, "11:10", "12:30"],
  [11, "Usos industriales y herramientas", B, "12:30", "13:40"],
  [12, "Modelo FRCU e impacto económico", B, "13:40", "14:50"],
  [13, "Viabilidad de la aplicación", B, "14:50", "16:00"],
  [14, "Reflexión crítica", B, "16:00", "17:10"],
  [15, "Conclusiones", B, "17:10", "18:40"],
  [16, "Referencias y cierre", B, "18:40", "20:00"],
];

// ---- helpers de documento
const run = (text, o = {}) =>
  new TextRun({
    text,
    font: "Arial",
    size: o.size || 28,
    bold: o.bold,
    italics: o.italics,
    color: o.color,
  });
const par = (text, o = {}) =>
  new Paragraph({
    spacing: { after: o.after ?? 140, before: o.before ?? 0, line: 300 },
    keepNext: o.keepNext,
    children: Array.isArray(text) ? text : [run(text, o)],
  });
const h1 = (t) => par(t, { size: 40, bold: true, color: OSCURO, after: 60 });
const h2 = (t) =>
  par(t, {
    size: 32,
    bold: true,
    color: AZUL,
    before: 260,
    after: 120,
    keepNext: true,
  });
const nota = (t) => par(`[${t}]`, { italics: true, color: GRIS, size: 24 });
const enganche = (t) =>
  par([
    run("ENGANCHE:  ", { bold: true, color: NARANJA }),
    run(t, { color: NARANJA }),
  ]);
const decir = (...ts) => ts.map((t) => par(t));
const borde = { style: BorderStyle.SINGLE, size: 4, color: "BBBBBB" };
const bordes = { top: borde, bottom: borde, left: borde, right: borde };
const celda = (t, w, o = {}) =>
  new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: bordes,
    shading: o.fill
      ? { type: ShadingType.CLEAR, fill: o.fill, color: "auto" }
      : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [
      new Paragraph({
        children: [run(t, { size: 24, bold: o.bold, color: o.color })],
      }),
    ],
  });
function tabla(cab, filas, anchos) {
  const head = new TableRow({
    tableHeader: true,
    children: cab.map((c, i) =>
      celda(c, anchos[i], { fill: AZUL, bold: true, color: "FFFFFF" }),
    ),
  });
  return new Table({
    width: { size: anchos.reduce((a, b) => a + b), type: WidthType.DXA },
    columnWidths: anchos,
    rows: [
      head,
      ...filas.map(
        (f) => new TableRow({ children: f.map((c, i) => celda(c, anchos[i])) }),
      ),
    ],
  });
}
const cab = (k) => {
  const [n, t, q, t0, t1] = T[k - 1];
  return h2(`DIAPOSITIVA ${n} · ${t}   ${q} · ${t0} a ${t1}`);
};

// ---- contenido
const c = [];
c.push(h1("Guion de exposición"));
c.push(
  par("SAT (3-SAT): algoritmo exacto frente a WalkSAT", {
    size: 32,
    bold: true,
    color: AZUL,
    after: 60,
  }),
);
c.push(
  par(
    "Teoría de la Computabilidad · Trabajo Final 2026 · 20 minutos · 2 personas",
    { color: GRIS, size: 24 },
  ),
);
c.push(h2("Cómo usar este guion"));
c.push(
  par(
    "Está escrito para leerse tal cual se dice, con frases cortas. Las frases en color con la palabra ENGANCHE son los puentes entre diapositivas. Lo que está entre corchetes en gris es una indicación, no se dice.",
  ),
);
c.push(
  par(
    `${A} expone las diapositivas 1 a 8 (0:00 a 10:00) y ${B} las diapositivas 9 a 16 (10:00 a 20:00). Todas las cifras salen de los archivos de resultados corridos en nuestro equipo (${hw.cpu}, ${hw.ram} de RAM, ${hw.so}, Python ${hw.python}, ${hw.fecha}). Si vuelven a correr los experimentos, regeneren este guion y la presentación con node generadores/guion.js y node generadores/presentacion.js: las cifras se actualizan solas.`,
  ),
);
c.push(h2("Reparto"));
c.push(
  tabla(
    ["Quién", "Diapositivas", "Minutos"],
    [
      [
        A,
        "1 a 8: problema, fuerza bruta, equipo, resultados del exacto, WalkSAT y comparación",
        "0:00 a 10:00",
      ],
      [
        B,
        "9 a 16: escalabilidad, casos reales, usos y herramientas, modelo FRCU, viabilidad, reflexión, conclusiones, referencias",
        "10:00 a 20:00",
      ],
    ],
    [1700, 6200, 1700],
  ),
);
c.push(h2("Cifras clave (para tener a mano)"));
c.push(
  tabla(
    ["Dato", "Valor"],
    [
      [
        "Equipo",
        `${hw.cpu}, ${hw.nucleosFisicos} núcleos (${hw.nucleosLogicos} lógicos), ${hw.ram} RAM, ${hw.so}, Python ${hw.python}. Límite ${fn(hw.limite)} s por corrida.`,
      ],
      [
        "Asignaciones con n variables",
        `2 elevado a la n (n=30: ${fn(n30.asignaciones)})`,
      ],
      [
        "Exacto, insatisfacible, n=20 / 26 / 30",
        `${ft(u(20).t_exacto_s)} / ${ft(u(26).t_exacto_s)} / ${ft(n30.t_exacto_s)} (${min1(n30.t_exacto_s)} min)`,
      ],
      [
        "Exacto, insatisfacible, n=32",
        `NO TERMINÓ en 1 h (cortado a ${fn(n32u.t_exacto_s, 0)} s; llevaba ${fn(n32u.asignaciones)} de ${fn(todas32)} asignaciones, ${pctCorte} %)`,
      ],
      ["Razón t(n+2) / t(n)", ratioTxt],
      [
        "Exacto con solución, n=32",
        `${fn(n32p.t_exacto_s)} s (frena al hallar la solución)`,
      ],
      [
        "WalkSAT con solución",
        `10 de 10 en todos los n (10 a 32); máximo ${ft(D.tmaxWalkSat)}; n=32: ${ft(n32p.t_walksat_medio_s)}`,
      ],
      [
        "WalkSAT en insatisfacibles",
        `0 de 10 siempre; ≈ ${ft(walkUnsatMedio)} (100.000 pasos); error relativo ${err(D.errUnsat.min)} a ${err(D.errUnsat.max)}`,
      ],
      [
        "Densidad, n=16",
        `razón m/n de ${fn(d.rMin)} a ${fn(d.rMax)}, ${d.porRazon} fórmulas cada una: satisfacibles ${d.satTotal} de ${d.instancias}; WalkSAT halló ${d.walkTotal}; falsos SAT: ${d.falsos}`,
      ],
      [
        "Escala",
        D.escala
          .map((e) => `n=${fn(e.n)}: ${e.ok}/${e.corridas}, ${ft(e.tMedio)}`)
          .join(" · "),
      ],
      [
        "n=10.000",
        `0 de ${nMax.corridas}; tope ${fn(nMax.flipsMax)} pasos (≈ ${ft(nMax.tMax)}); hasta ${fn(100 - nMax.pctMin, 2)} % de cláusulas sin cumplir`,
      ],
    ],
    [3200, 6400],
  ),
);

// PARTE 1
c.push(par("", { after: 0 }));
c.push(h2(`PARTE 1 · ${A} · diapositivas 1 a 8 · 0:00 a 10:00`));
c.push(cab(1));
c.push(
  ...decir(
    `Buenos días, profesor. Somos ${A} y ${B}.`,
    "Les traemos una pregunta: si tenemos una fórmula lógica con muchísimas variables, ¿se puede saber si existe alguna combinación de verdadero y falso que la haga verdadera?",
    "Ese es el problema SAT, y fue el primer problema que se demostró NP-completo.",
    "Vamos a ver un algoritmo exacto, que siempre acierta; una heurística, WalkSAT, que es muy rápida pero no da garantías; mediciones reales hechas en nuestro equipo; y dónde aparece SAT en la vida real, en particular en el armado de horarios.",
  ),
);
c.push(enganche("Empecemos por entender qué pregunta hace SAT."));

c.push(cab(2));
c.push(nota("Señalar la fórmula en la pantalla"));
c.push(
  ...decir(
    "SAT pregunta si una fórmula booleana se puede hacer verdadera. Miren esta fórmula: tiene 4 variables y 3 cláusulas.",
    "Cada cláusula es un OR de literales: una variable o su negación. Las cláusulas se unen con AND, o sea que tienen que cumplirse todas a la vez.",
    "En 3-SAT, cada cláusula tiene exactamente tres literales. Es el caso que usamos en el trabajo. Una variante es Max-SAT, donde no se pide cumplir todo, sino la mayor cantidad posible de cláusulas.",
    "Tres ideas clave. Primero: SAT es NP-completo, lo demostraron Cook y Levin, y 3-SAT también lo es. Segundo: verificar es fácil. Si me dan una asignación, recorro las cláusulas y compruebo que cada una tenga un literal verdadero: es tiempo polinomial, por eso está en NP. Tercero: buscar es difícil. Con n variables hay 2 elevado a la n asignaciones posibles. Ese es nuestro espacio de búsqueda.",
  ),
);
c.push(enganche("Veamos qué pasa si probamos todas esas asignaciones."));

c.push(cab(3));
c.push(nota("Señalar el pseudocódigo"));
c.push(
  ...decir(
    "La fuerza bruta hace exactamente eso. Cada número entre 0 y 2 a la n menos 1, escrito en binario, es una asignación: cada bit dice si la variable vale verdadero o falso.",
    "Para cada asignación recorremos las cláusulas. Si alguna no tiene ningún literal verdadero, descartamos la asignación y pasamos a la siguiente. Si una cumple todas, devolvemos SAT y la asignación. Si terminamos sin encontrar ninguna, devolvemos UNSAT. Y si se pasa del límite de tiempo, devolvemos TIMEOUT.",
    "La complejidad es O de 2 a la n por m por k: n variables, m cláusulas y k literales por cláusula.",
    "Dos cosas importantes. Es exacto: siempre da la respuesta correcta, y es el único que puede demostrar que una fórmula no tiene solución. Y es exponencial: cada variable que agregamos duplica el trabajo.",
  ),
);
c.push(
  enganche("¿Se cumple eso en la práctica? Primero, en qué equipo lo medimos."),
);

c.push(cab(4));
c.push(
  ...decir(
    `Todas las mediciones las hicimos en un mismo equipo: un ${hw.cpu}, con ${hw.nucleosFisicos} núcleos y ${hw.nucleosLogicos} hilos, ${hw.ram} de RAM, ${hw.so} y Python ${hw.python}. Se corrió el ${hw.fecha}. El límite por corrida del exacto fue de ${fn(hw.limite)} segundos, una hora.`,
    "Los dos algoritmos resolvieron exactamente las mismas fórmulas, guardadas en formato DIMACS y con semillas fijas, así que se puede repetir todo.",
    "Para cada n par entre 10 y 32 usamos dos fórmulas: una con solución garantizada, que llamamos planted, y una insatisfacible. A WalkSAT lo corrimos 10 veces por fórmula, con p igual a 0,5 y 100.000 pasos. Para escalabilidad, de 100 a 10.000 variables, tres fórmulas planted y tres aleatorias por tamaño, hasta 5 millones de pasos.",
  ),
);
c.push(enganche("Con eso, veamos qué le pasó al algoritmo exacto."));

c.push(cab(5));
c.push(
  nota(
    "Señalar el gráfico. Escala logarítmica en el eje vertical. La x roja es el TIMEOUT.",
  ),
);
c.push(
  ...decir(
    "Para medir el exacto usamos fórmulas insatisfacibles construidas a propósito. ¿Por qué? Porque obligan a recorrer todas las asignaciones: no hay solución donde frenar antes.",
    `Con 26 variables tardó ${ft(u(26).t_exacto_s)}. Con 28, ${ft(u(28).t_exacto_s)}. Con 30 variables, ${fn(n30.t_exacto_s)} segundos, unos ${min1(n30.t_exacto_s)} minutos. Evaluó exactamente ${fn(n30.asignaciones)} asignaciones, que es 2 a la 30.`,
    `Fíjense en el patrón: cada dos variables más, el tiempo se multiplica por cuatro, que es 2 al cuadrado. Calculado con nuestros datos, entre n=${razon.desde} y n=${razon.hasta}, la razón vale en promedio ${fn(razon.gm, 1)}, con valores entre ${fn(razon.min, 1)} y ${fn(razon.max, 1)}. En el gráfico logarítmico eso se ve como una recta ascendente: crecimiento exponencial.`,
    `Y con 32 variables, la x roja: el exacto NO terminó en una hora. Lo cortamos a los ${fn(n32u.t_exacto_s, 0)} segundos, cuando había evaluado el ${pctCorte} por ciento de las asignaciones. No tenemos un tiempo medido para ese caso; por la razón de cuatro, desde los ${min1(n30.t_exacto_s)} minutos de n=30 estimamos unas ${proyeccion32h} horas, pero eso es una proyección, no una medición.`,
    `Una aclaración: en las fórmulas con solución el tiempo depende de la suerte. Con n=32 tardó ${fn(n32p.t_exacto_s)} segundos, y con n=30 solo ${ft(p(30).t_exacto_s)}, menos que con n=28 (${ft(p(28).t_exacto_s)}). Frena apenas encuentra la primera solución.`,
  ),
);
c.push(
  enganche(
    "Entonces, ¿cómo se resuelven fórmulas con miles de variables? Cambiando de enfoque: una heurística.",
  ),
);

c.push(cab(6));
c.push(
  ...decir(
    `WalkSAT es una búsqueda local con azar. Empieza con una asignación al azar y mira qué cláusulas están insatisfechas. Elige una de ellas y da vuelta el valor de una de sus variables. Repite hasta que todas las cláusulas se cumplan, o hasta que se le acaben los pasos.`,
    "¿Qué variable da vuelta? La que rompe menos cláusulas que hoy están satisfechas. Y con probabilidad p, que en nuestras pruebas fue 0,5, da vuelta una cualquiera de la cláusula, para no quedarse atascada.",
    "Ahora la parte más importante. WalkSAT es incompleto. Si no encuentra solución, devuelve UNKNOWN. Y UNKNOWN no significa insatisfacible: significa que no la encontró. Solo el algoritmo exacto puede demostrar que no existe.",
  ),
);
c.push(enganche("Veamos cómo se comparan, sobre las mismas fórmulas."));

c.push(cab(7));
c.push(nota("Señalar el gráfico de tiempos"));
c.push(
  ...decir(
    `En las fórmulas con solución, WalkSAT la encontró las 10 veces que lo corrimos, en todos los tamaños, y nunca tardó más de ${ft(D.tmaxWalkSat)}. Con 32 variables tardó ${ft(n32p.t_walksat_medio_s)}; el exacto, ${fn(n32p.t_exacto_s)} segundos.`,
    `En las insatisfacibles, el exacto demostró UNSAT hasta n=30, en ${min1(n30.t_exacto_s)} minutos. Con n=32 no terminó en una hora: ahí no hay un punto medido, hay una marca de TIMEOUT.`,
  ),
);
c.push(enganche("Los números exactos, en la tabla."));

c.push(cab(8));
c.push(
  nota(
    "Señalar la tabla. Celdas de WalkSAT: éxitos sobre 10 y tiempo, o error relativo en las insatisfacibles",
  ),
);
c.push(
  ...decir(
    "En la tabla, con 20, 26, 30 y 32 variables. En las fórmulas con solución, WalkSAT acertó 10 de 10 siempre. En las insatisfacibles, 0 de 10: informó UNKNOWN.",
    `El error relativo es la fracción de cláusulas que la mejor asignación de WalkSAT dejó sin cumplir. En las insatisfacibles fue de ${err(D.errUnsat.min)} a ${err(D.errUnsat.max)}, y bajó al subir n porque hay más cláusulas. WalkSAT gastó siempre los 100.000 pasos, unos ${ft(walkUnsatMedio)}, y no puede decir nada más.`,
    `Hicimos además un experimento con 16 variables, 30 fórmulas aleatorias por cada razón de cláusulas sobre variables, de ${fn(d.rMin)} a ${fn(d.rMax)}. Al subir la razón, las fórmulas con solución cayeron de 30 a 1. De las ${d.satTotal} satisfacibles que detectó el exacto, WalkSAT encontró ${d.walkTotal}, y nunca dijo SAT sobre una insatisfacible.`,
    "Una aclaración honesta: usamos una sola fórmula de cada tipo por tamaño, y las planted resultaron fáciles. Por eso decimos que la comparación es ilustrativa.",
  ),
);
c.push(
  enganche(
    `Con 32 variables el exacto ya no llegó. ¿Qué pasa con tamaños reales? Eso lo sigue ${B}: le paso la palabra.`,
  ),
);

// PARTE 2
c.push(h2(`PARTE 2 · ${B} · diapositivas 9 a 16 · 10:00 a 20:00`));

c.push(cab(9));
c.push(nota("Señalar el gráfico y la tabla"));
c.push(
  ...decir(
    `Gracias, ${A}. Probamos WalkSAT con 100, 500, 1.000, 5.000 y 10.000 variables, tres fórmulas planted y tres aleatorias por tamaño. Hasta 1.000 variables resolvió las seis, con un tiempo medio de ${ft(D.escala.find((e) => e.n === 1000).tMedio)} en 1.000.`,
    `Con 5.000 resolvió ${e5000.ok} de 6: las tres planted y ${e5000.okAleat} de las 3 aleatorias, con un tiempo medio de ${ft(e5000.tMedio)}.`,
    `Con 10.000 variables no llegó a la solución en ninguna de las ${nMax.corridas} corridas. Pero atención: se detuvo por el tope de ${fn(nMax.flipsMax)} pasos que fijamos, unos ${ft(nMax.tMax)}, y dejó hasta ${fn(100 - nMax.pctMin, 2)} por ciento de las cláusulas sin cumplir. Eso no dice que no haya solución ni que WalkSAT no pueda: es el límite de nuestro experimento.`,
    "La dificultad no depende solo de n: las planted, con más cláusulas por variable, fueron más fáciles que las aleatorias.",
  ),
);
c.push(enganche("Ahora, ¿dónde se usa esto de verdad?"));

c.push(cab(10));
c.push(
  ...decir(
    "Empecemos por la Argentina. En la Universidad Torcuato Di Tella, Ferreiro y Marenco presentaron en las 53 JAIIO, en el SIIIO, un modelo de programación lineal entera, ILP, para armar los horarios de clases, con 35 módulos. El impacto que reportan es cualitativo; no tenemos una cifra.",
    "Aclaramos algo importante: eso no es SAT puro. Lo traemos porque es el caso argentino de horarios mejor documentado, y porque las restricciones de un horario, como que no choquen dos comisiones, son restricciones lógicas que se pueden codificar en CNF.",
    "El puente hacia SAT lo da el coloreo de grafos. En la Universidad Nacional de Rosario, Escalante, Chiesa y Fekete trabajaron este problema en SUMA 2019. Colorear un grafo con k colores se reduce a SAT: una variable dice que el vértice v lleva el color c, y se pide un color por vértice y colores distintos entre vecinos. Y armar horarios es, en esencia, colorear: las comisiones son vértices, los choques son aristas y las franjas son colores.",
  ),
);
c.push(enganche("Veamos qué usos y herramientas hay del lado de SAT."));

c.push(cab(11));
c.push(
  ...decir(
    "SAT se usa en verificación de hardware y software, con bounded model checking; en la resolución de dependencias de paquetes, por ejemplo en Eclipse con la biblioteca Sat4j y en libsolv; y en horarios, con SAT y MaxSAT: Asín Achá y Nieuwenhuis, en 2014, lo aplicaron a horarios universitarios.",
    "Las herramientas: MiniSat y Glucose son solvers clásicos; Kissat ganó la SAT Competition 2024; Z3 es un solver SMT; y OR-Tools tiene CP-SAT. Todos son libres y están pensados para instancias mucho más grandes que las de nuestro prototipo en Python.",
  ),
);
c.push(enganche("Ahora, qué haría falta para aplicarlo en nuestra facultad."));

c.push(cab(12));
c.push(nota("Insistir en la palabra HIPOTÉTICO"));
c.push(
  ...decir(
    `Para nuestra facultad, la UTN de Concepción del Uruguay, proponemos un modelo hipotético. La FRCU publica los horarios de cursada en PDF, con comisiones, día y franja. Una variable dice si la comisión c se dicta en la franja f. Las cláusulas piden al menos una franja por comisión, a lo sumo una, y que no coincidan comisiones que comparten docente o alumnos.`,
    "Hicimos un ejemplo mínimo, de 9 variables y 19 cláusulas, y ambos algoritmos encontraron un horario válido. Lo decimos con claridad: no hay evidencia pública de que la FRCU use SAT. Es un modelo hipotético.",
    "Sobre el impacto económico: no existe una cifra pública atribuible a SAT, y no vamos a inventarla. Como analogía, que no es SAT: el horario ferroviario neerlandés, optimizado con investigación operativa, se estimó en 50 millones de euros por año, según Kroon, Huisman y colaboradores, en Interfaces, 2009.",
  ),
);
c.push(enganche("Con eso, la viabilidad."));

c.push(cab(13));
c.push(
  ...decir(
    "Viabilidad en cuatro planos.",
    `Técnica: nuestra fuerza bruta no escala, con 32 variables ya pasa la hora, y WalkSAT no demuestra insatisfacibilidad. Para producción se usarían solvers como MiniSat o Kissat.`,
    "Datos: hoy tenemos PDFs. Harían falta docentes, aulas, capacidades y disponibilidades en formato usable.",
    "Económica: el costo es bajo, porque las herramientas son software libre. El ahorro no lo podemos cuantificar; es cualitativo.",
    "Del modelo: las preferencias, como evitar huecos o equilibrar la carga, son objetivos, no cláusulas duras. Eso lleva a MaxSAT o a ILP, no a SAT puro.",
  ),
);
c.push(enganche("Y ahora, con honestidad, los límites de lo que hicimos."));

c.push(cab(14));
c.push(
  ...decir(
    `Primero, los límites. Una sola fórmula y una sola corrida del exacto por tamaño. Las planted dan tiempos irregulares: con n=30 tardó ${ft(p(30).t_exacto_s)} y con n=28, ${ft(p(28).t_exacto_s)}. Una sola máquina y Python, que es lento. No ajustamos el parámetro p. En escala, tres fórmulas por punto. Y la aplicación es hipotética.`,
    "Segundo, qué haríamos distinto: repetir y promediar las corridas, comparar contra un solver como MiniSat, usar benchmarks estándar y estudiar el efecto de p.",
    "Y una reflexión de fondo: SAT es un caso donde un problema que en el peor caso es intratable se resuelve en la práctica, pero eso depende de la estructura de la instancia. Nuestros datos lo muestran: con 32 variables, WalkSAT halló la solución en milisegundos, y el exacto no terminó en una hora cuando no la había.",
  ),
);
c.push(enganche("Cerramos con las conclusiones."));

c.push(cab(15));
c.push(
  ...decir(
    "Resumimos en cuatro ideas.",
    `Uno. La fuerza bruta es correcta pero exponencial: con n=30 tardó ${min1(n30.t_exacto_s)} minutos en una fórmula insatisfacible, y con n=32 no terminó en una hora.`,
    `Dos. WalkSAT resolvió las fórmulas con solución en ${ft(D.tmaxWalkSat)} como máximo hasta n=32, y fórmulas de 5.000 variables en segundos.`,
    `Tres. A cambio, perdemos la garantía: UNKNOWN no es UNSAT. Y con 10.000 variables no encontró solución dentro del tope que fijamos.`,
    "Cuatro. Usamos el exacto cuando n es chico o necesitamos demostrar que no hay solución, y WalkSAT cuando n crece y alcanza con encontrar una. En la práctica real se usan solvers industriales.",
    "Como límites: una instancia por tamaño en la comparación, fórmulas planted fáciles, mediciones en Python, y una aplicación que es un modelo propuesto, no un caso verificado.",
  ),
);
c.push(enganche("Dejamos las referencias."));

c.push(cab(16));
c.push(
  nota("Dejar la diapositiva de referencias en pantalla durante las preguntas"),
);
c.push(
  ...decir(
    "Las fuentes que usamos están en pantalla: Cook, Selman, Kautz y Cohen, el trabajo de la UTDT, el de la UNR, Asín Achá y Nieuwenhuis, y Kroon y Huisman. No citamos nada que no hayamos podido verificar.",
    `Eso es todo de nuestra parte. Muchas gracias. Quedamos a disposición para las preguntas.`,
  ),
);

// Preguntas
const QA = [
  [
    "¿Por qué SAT está en NP?",
    "Porque dada una asignación se verifica en tiempo polinomial: se recorren las cláusulas y se comprueba que cada una tenga un literal verdadero.",
  ],
  [
    "¿Por qué 3-SAT es NP-completo?",
    "Porque cualquier fórmula SAT se transforma en una 3-SAT equivalente en tiempo polinomial, partiendo las cláusulas largas con variables nuevas. Y porque 3-SAT está en NP.",
  ],
  [
    "¿Diferencia entre SAT, 3-SAT y Max-SAT?",
    "SAT admite cláusulas de cualquier largo. 3-SAT, exactamente tres literales. Max-SAT busca cumplir la mayor cantidad posible de cláusulas: es de optimización.",
  ],
  [
    "¿Por qué WalkSAT falla en las fórmulas insatisfacibles?",
    `Porque ahí no hay ninguna solución que encontrar, y WalkSAT solo explora asignaciones guiado por el azar. Agotó los 100.000 pasos y devolvió UNKNOWN, con la mejor asignación dejando ${err(D.errUnsat.min)} a ${err(D.errUnsat.max)} de las cláusulas sin cumplir. Para demostrar que no existe solución hay que cubrir todo el espacio, como hace el exacto.`,
  ],
  [
    "¿Por qué midieron el exacto con fórmulas insatisfacibles?",
    "Porque obligan a recorrer las 2 a la n asignaciones. Con una fórmula satisfacible se frena al hallar la primera solución y el tiempo depende de la suerte; lo vimos: con n=30 tardó menos que con n=28.",
  ],
  [
    "¿Llegaron a una hora de ejecución?",
    `Superamos la hora: con 32 variables insatisfacible el exacto no terminó en una hora, lo cortamos a ${fn(n32u.t_exacto_s, 0)} segundos con el ${pctCorte} % de las asignaciones evaluadas. Lo último que sí terminó fue n=30, en ${fn(n30.t_exacto_s)} segundos (${min1(n30.t_exacto_s)} minutos). No tenemos un tiempo medido para n=32; la proyección con la razón de cuatro da unas ${proyeccion32h} horas, y lo decimos como proyección.`,
  ],
  [
    "¿Qué pasó con 10.000 variables?",
    `No encontró solución en ${fn(nMax.flipsMax)} pasos, el tope que fijamos (unos ${ft(nMax.tMax)} por corrida). Dejó hasta ${fn(100 - nMax.pctMin, 2)} % de las cláusulas sin cumplir. No significa que sea insatisfacible.`,
  ],
  [
    "¿Qué es una fórmula planted?",
    "Una fórmula construida a partir de una solución oculta: cada cláusula se corrige para que esa solución la cumpla. Garantiza que tiene solución, pero suele ser más fácil que una aleatoria.",
  ],
  [
    "¿Por qué p = 0,5?",
    "Es un valor típico de la literatura y no lo optimizamos. No estudiamos el efecto de p de forma sistemática; es trabajo futuro.",
  ],
  [
    "¿Por qué incluyeron un caso de ILP si el TP es de SAT?",
    "Porque es el caso argentino de horarios mejor documentado, y lo presentamos como tal: no es SAT puro. Las restricciones de un horario son lógicas y se codifican en CNF, y el coloreo de grafos de la UNR es el puente concreto hacia SAT. Lo usamos para mostrar que el problema existe y importa, no para afirmar que ese caso use SAT.",
  ],
  [
    "¿Por qué no hay datos de ahorro?",
    "Porque no hay ninguna cifra pública atribuible a SAT, y inventarla sería deshonesto. El impacto del caso de la UTDT es cualitativo. La única cifra, 50 millones de euros por año del horario ferroviario neerlandés, es una analogía de optimización con otros métodos, y la rotulamos así.",
  ],
  [
    "¿La UTN usa SAT para los horarios?",
    "No encontramos evidencia pública. Documentamos que publican horarios con comisiones y franjas, y proponemos un modelo hipotético.",
  ],
  [
    "¿Un solver industrial lo resolvería?",
    "Seguramente mucho más rápido: MiniSat o Kissat están hechos para eso. Pero no los medimos; nuestro trabajo compara un exacto y una heurística en Python, desarrollados con asistencia de una herramienta de IA (Claude) y revisados y ejecutados por nosotros.",
  ],
  [
    "¿Qué es DIMACS?",
    "Un formato de texto estándar para fórmulas CNF. El código del trabajo lo lee y lo escribe, así se pueden usar benchmarks externos.",
  ],
];
c.push(h2("PREGUNTAS PROBABLES"));
c.push(
  par(
    "Respuestas cortas para decir. Cada pregunta indica quién responde, según la diapositiva de la que viene el tema; el otro complementa.",
    { italics: true, color: GRIS, size: 24 },
  ),
);
// quién responde, según la diapositiva de origen del tema (1-8: A, 9-16: B), mismo orden que QA
const RESP = [[A, 2], [A, 2], [A, 2], [A, 8], [A, 5], [A, 5], [B, 9], [A, 4], [A, 6], [B, 10], [B, 12], [B, 12], [B, 11], [A, 4]];
for (const [i, [q, a]] of QA.entries()) {
  c.push(par(`${q}  (Responde ${RESP[i][0]}, diap. ${RESP[i][1]})`, { bold: true, color: AZUL, after: 40, keepNext: true }));
  c.push(par(a));
}
c.push(h2("Si se traban"));
c.push(
  par(
    "Para ganar tiempo: «Es una buena pregunta. Lo que sabemos con certeza es esto...» y repetir lo que ya está en el guion.",
  ),
);
c.push(
  par(
    "Si no saben algo: «Eso no lo medimos en nuestro trabajo. Lo que sí podemos decir es...». No inventen datos.",
  ),
);
c.push(
  par(
    "Si sobra tiempo: mostrar la fórmula de la diapositiva 2 y decir una asignación que la cumple. Si falta: acortar la diapositiva 11 a «MiniSat, Glucose, Kissat, Z3 y OR-Tools» y saltar a la 12.",
  ),
);

const doc = new Document({
  styles: { default: { document: { run: { font: "Arial", size: 28 } } } },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 },
        },
      },
      children: c,
    },
  ],
});
const salida = path.join(D.RAIZ, "SAT_Guion_Exposicion.docx");
Packer.toBuffer(doc).then((b) => {
  fs.writeFileSync(salida, b);
  console.log("OK", salida);
});
