// Genera GUIA_DE_ESTUDIO.docx (material de estudio para defender el TP de SAT).
// Uso: node generadores/estudio.js   -> escribe GUIA_DE_ESTUDIO.docx en la raiz del proyecto.
// Las cifras salen de los CSV/TXT reales via ./datos.js. Marcado inline: **negrita**, §codigo§.
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
  AlignmentType,
  LevelFormat,
  Footer,
  PageNumber,
  PageBreak,
} = require("docx");
const D = require("./datos");
const { fn, ft, hw, razon, unsat, planted, porN } = D;

const AZUL = "1F4E79",
  NARANJA = "C55A11",
  VERDE = "2E7D32",
  GRIS = "666666",
  OSCURO = "16222E";
const CAMILO = "Camilo",
  FACU = "Facundo";

// ---------- cifras reales
const u = (n) => porN(unsat, n),
  p = (n) => porN(planted, n);
const n30 = u(30),
  n32u = u(32),
  n32p = p(32),
  n28p = p(28),
  n30p = p(30);
const esc = (n) => D.escala.find((e) => e.n === n);
const d = D.densidad;
const rg = fn(razon.gm, 2),
  rmin = fn(razon.min, 2),
  rmax = fn(razon.max, 2);
const pctCorte = fn((n32u.asignaciones / 2 ** 32) * 100, 1);
const proy32 = fn((n30.t_exacto_s * razon.gm) / 3600, 1);
const velCorte = n32u.asignaciones / n32u.t_exacto_s;
const proy32v = fn(2 ** 32 / velCorte / 3600, 1);
const walkUnsatMed = D.walkUnsatT.reduce((a, b) => a + b) / D.walkUnsatT.length;
const hist = D.historial;
const err = (x) => fn(x * 100, 1) + " %";
const aprox = (a, n) => fn((a / 2 ** n) * 100, 1) + " %";

// ---------- helpers de documento
const c = [];
function runs(text, o = {}) {
  const size = o.size || 22;
  const out = [];
  text.split(/(\*\*[^*]+\*\*)/).filter((s) => s !== "").forEach((seg) => {
    const negrita = seg.startsWith("**") && seg.endsWith("**") && seg.length > 4;
    const cuerpo = negrita ? seg.slice(2, -2) : seg;
    cuerpo.split(/(§[^§]+§)/).filter((s) => s !== "").forEach((s) => {
      if (s.startsWith("§") && s.endsWith("§") && s.length > 2)
        out.push(new TextRun({ text: s.slice(1, -1), font: "Consolas", size: size - 2, bold: negrita || o.bold, color: "7A2E00" }));
      else out.push(new TextRun({ text: s, font: "Arial", size, bold: negrita || o.bold, color: o.color, italics: o.italics }));
    });
  });
  return out;
}
const mk = (text, o = {}) =>
  new Paragraph({
    spacing: { after: o.after ?? 120, before: o.before ?? 0, line: 276 },
    keepNext: o.keepNext,
    pageBreakBefore: o.pageBreakBefore,
    alignment: o.align,
    children: runs(text, o),
  });
const P = (t, o) => c.push(mk(t, o));
const H1 = (t) =>
  c.push(
    mk(t, {
      size: 38,
      bold: true,
      color: OSCURO,
      after: 160,
      pageBreakBefore: true,
      keepNext: true,
    }),
  );
const H2 = (t) =>
  c.push(
    mk(t, {
      size: 30,
      bold: true,
      color: AZUL,
      before: 280,
      after: 120,
      keepNext: true,
    }),
  );
const H3 = (t) =>
  c.push(
    mk(t, {
      size: 24,
      bold: true,
      color: NARANJA,
      before: 200,
      after: 80,
      keepNext: true,
    }),
  );
const B = (t, nivel = 0) =>
  c.push(
    new Paragraph({
      numbering: { reference: "vi", level: nivel },
      spacing: { after: 70, line: 276 },
      children: runs(t),
    }),
  );
const BS = (...ts) => ts.forEach((t) => B(t));
const CODE = (txt) => {
  const lineas = txt.replace(/^\n/, "").replace(/\n$/, "").split("\n");
  lineas.forEach((l, i) =>
    c.push(
      new Paragraph({
        spacing: { after: 0, line: 240 },
        shading: { type: ShadingType.CLEAR, fill: "F1F3F5", color: "auto" },
        indent: { left: 200, right: 200 },
        keepNext: i < lineas.length - 1,
        keepLines: true,
        border:
          i === 0
            ? {
                top: {
                  style: BorderStyle.SINGLE,
                  size: 4,
                  color: "CCCCCC",
                  space: 4,
                },
              }
            : i === lineas.length - 1
              ? {
                  bottom: {
                    style: BorderStyle.SINGLE,
                    size: 4,
                    color: "CCCCCC",
                    space: 4,
                  },
                }
              : undefined,
        children: [
          new TextRun({ text: l === "" ? " " : l, font: "Consolas", size: 18 }),
        ],
      }),
    ),
  );
  c.push(mk("", { after: 60 }));
};
const borde = { style: BorderStyle.SINGLE, size: 4, color: "BBBBBB" };
const bordes = { top: borde, bottom: borde, left: borde, right: borde };
const celda = (t, w, o = {}) =>
  new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: bordes,
    shading: o.fill
      ? { type: ShadingType.CLEAR, fill: o.fill, color: "auto" }
      : undefined,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    children: String(t)
      .split("\n")
      .map(
        (l) =>
          new Paragraph({
            spacing: { after: 30 },
            children: runs(l, {
              size: o.size || 19,
              bold: o.bold,
              color: o.color,
            }),
          }),
      ),
  });
function TABLA(cab, filas, anchos, o = {}) {
  const tot = anchos.reduce((a, b) => a + b);
  c.push(
    new Table({
      width: { size: tot, type: WidthType.DXA },
      columnWidths: anchos,
      rows: [
        new TableRow({
          tableHeader: true,
          cantSplit: true,
          children: cab.map((x, i) =>
            celda(x, anchos[i], {
              fill: AZUL,
              bold: true,
              color: "FFFFFF",
              size: o.size,
            }),
          ),
        }),
        ...filas.map(
          (f, k) =>
            new TableRow({
              cantSplit: true,
              children: f.map((x, i) =>
                celda(x, anchos[i], {
                  fill: k % 2 ? "F7F9FB" : undefined,
                  size: o.size,
                }),
              ),
            }),
        ),
      ],
    }),
  );
  c.push(mk("", { after: 100 }));
}
// caja de color: tipo OJO | MEMORIA | HONESTIDAD | EJEMPLO
function CAJA(tipo, texto) {
  const col = {
    OJO: ["FDF0E6", NARANJA],
    MEMORIA: ["E8F0F8", AZUL],
    HONESTIDAD: ["EAF4EA", VERDE],
    EJEMPLO: ["F3F0F8", "5B3F8C"],
  }[tipo];
  const partes = Array.isArray(texto) ? texto : [texto];
  c.push(
    new Table({
      width: { size: 9026, type: WidthType.DXA },
      columnWidths: [9026],
      rows: [
        new TableRow({
          cantSplit: true,
          children: [
            new TableCell({
              width: { size: 9026, type: WidthType.DXA },
              shading: { type: ShadingType.CLEAR, fill: col[0], color: "auto" },
              borders: {
                top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                left: { style: BorderStyle.SINGLE, size: 24, color: col[1] },
              },
              margins: { top: 80, bottom: 80, left: 160, right: 140 },
              children: partes.map(
                (t, i) =>
                  new Paragraph({
                    spacing: { after: 50, line: 270 },
                    children:
                      i === 0
                        ? [
                            new TextRun({
                              text: tipo + ".  ",
                              font: "Arial",
                              size: 20,
                              bold: true,
                              color: col[1],
                            }),
                            ...runs(t, { size: 20 }),
                          ]
                        : runs(t, { size: 20 }),
                  }),
              ),
            }),
          ],
        }),
      ],
    }),
  );
  c.push(mk("", { after: 100 }));
}
const QA = (q, a) => {
  c.push(mk("P: " + q, { bold: true, color: AZUL, after: 40, keepNext: true }));
  (Array.isArray(a) ? a : [a]).forEach((t, i) =>
    c.push(mk(i === 0 ? "R: " + t : t, { after: 100 })),
  );
};

// =====================================================================
// PORTADA
// =====================================================================
c.push(
  mk("Teoría de la Computabilidad · Trabajo Final 2026", {
    size: 24,
    color: GRIS,
    after: 400,
    before: 1400,
  }),
);
c.push(
  mk("Guía de estudio", { size: 64, bold: true, color: OSCURO, after: 80 }),
);
c.push(
  mk("SAT (Problema de Satisfacción Booleana): algoritmo exacto frente a la heurística WalkSAT", {
    size: 34,
    bold: true,
    color: AZUL,
    after: 240,
  }),
);
c.push(
  mk("Material para entender, defender y exponer el trabajo", {
    size: 26,
    after: 600,
  }),
);
c.push(
  mk(
    "Para: Camilo Escar (diapositivas 1 a 8) y Facundo Iglesias (diapositivas 9 a 16)",
    { size: 24, after: 60 },
  ),
);
c.push(
  mk(
    "Materia: Teoría de la Computabilidad · Prof. Pascal · exposición de 20 minutos (10 + 10)",
    { size: 24, after: 400 },
  ),
);
CAJA("HONESTIDAD", [
  "Esta guía existe porque el código y el informe los generó una IA y ustedes tienen que poder defenderlos como propios. Eso solo se logra entendiendo cada línea, cada número y cada decisión. Léanla con el código abierto al lado y ejecuten los ejemplos a mano.",
  "Todas las cifras salen de los archivos del repositorio (CSV, hardware.txt, informe, presentación y guion) y coinciden con ellos. Lo que no está medido en el TP lo rotulo como [LITERATURA], [PROYECCIÓN], [HIPOTÉTICO] o [VERIFICAR].",
]);
P("Contenido", { bold: true, size: 26, color: AZUL, before: 200 });
BS(
  "**Parte 0** · Cómo usar la guía, mapa de la consigna y del repositorio",
  "**Parte 1** · Fundamentos que los dos tienen que dominar",
  "**Parte 2** · Cómo está hecho el código, función por función",
  "**Parte 3** · Registro de decisiones (qué se decidió, qué se descartó y por qué)",
  "**Parte 4** · Guía de la exposición: una sección por persona, diapositiva por diapositiva",
  "**Parte 5** · Resultados explicados: cómo leer cada gráfico y la tabla",
  "**Parte 6** · Banco de preguntas difíciles de la cátedra, con respuestas modelo",
  "**Parte 7** · Glosario, autoevaluación, checklist del día, qué NO decir y referencias a verificar",
  "**Anexo** · Inconsistencias a revisar (cosas del código o de los documentos que hay que corregir o saber)",
);

// =====================================================================
// PARTE 0
// =====================================================================
H1("Parte 0 · Cómo usar esta guía y mapa del TP");
H2("0.1 Cómo estudiar con esta guía");
P(
  "La exposición dura 20 minutos, pero la defensa oral (las preguntas) es donde se nota si entendieron o no. Un profesor de Computabilidad no se conforma con «WalkSAT es una heurística»: pregunta por qué no puede demostrar UNSAT, por qué 3-SAT es NP-completo, de dónde sale el 4,26. La guía está armada para que puedan contestar eso con tranquilidad.",
);
P("**Orden recomendado (cuatro sesiones):**");
BS(
  "**Sesión 1 (los dos juntos, ~2 h):** Parte 0 y Parte 1 completas. Frenen en cada ejemplo y háganlo a mano en un papel. Si algo no cierra, no sigan.",
  "**Sesión 2 (los dos juntos, ~2 h):** Parte 2 con el archivo §sat.py§ abierto. Corran los ejemplos chicos en una consola de Python. La meta: que cada uno pueda explicar §fuerza_bruta§ y §walksat§ sin mirar.",
  "**Sesión 3 (cada uno por su lado, ~2 h):** Parte 4, la sección propia. Ensayen en voz alta con el cronómetro. Después lean la sección del compañero a nivel básico.",
  "**Sesión 4 (juntos):** Parte 5 y Parte 6. Uno hace de profesor y le tira preguntas al otro (sin mirar las respuestas). Cierren con la autoevaluación y el checklist de la Parte 7.",
);
H3("Cómo leer los rótulos");
TABLA(
  ["Rótulo", "Qué significa", "Cómo se dice en voz alta"],
  [
    [
      "[MEDIDO]",
      "Salió de una corrida real del código en el equipo del TP (CSV).",
      "«Medimos…»",
    ],
    [
      "[DERIVADO]",
      "Cuenta hecha a partir de números medidos (una razón, un promedio, un porcentaje).",
      "«Calculado con nuestros datos…»",
    ],
    [
      "[PROYECCIÓN]",
      "Extrapolación que NO se midió (por ejemplo, el tiempo de n=32 si hubiera terminado).",
      "«Estimamos, pero no lo medimos…»",
    ],
    [
      "[LITERATURA]",
      "Resultado conocido y publicado, que NO se probó ni se midió en este TP (por ejemplo, el umbral 4,26).",
      "«Según la bibliografía…»",
    ],
    [
      "[HIPOTÉTICO]",
      "Modelo propio sin datos ni evidencia de uso real (el modelo para la FRCU).",
      "«Proponemos, como hipótesis…»",
    ],
    [
      "[VERIFICAR]",
      "Dato que se tomó de un resumen o buscador y hay que abrir la fuente antes de exponer.",
      "No decirlo hasta verificarlo",
    ],
  ],
  [1500, 4800, 2726],
);

H2("0.2 Qué pide la consigna y dónde está cada punto");
CAJA(
  "OJO",
  "La tabla sigue los 9 puntos de la consigna de la cátedra, en su orden. Aclaración de encuadre: el problema elegido (punto 1) es SAT, con instancias restringidas a 3-SAT. La fuerza bruta (punto 2) es la solución exacta de referencia y WalkSAT (punto 5) es la heurística elegida, de búsqueda local. WalkSAT no es el tema del trabajo.",
);
TABLA(
  ["#", "Punto de la consigna", "Dónde está", "Diapositiva"],
  [
    [
      "1",
      "Elegir un problema (SAT): presentación y relevancia",
      "Informe sec. 1, 2 y 3; README",
      "1, 2",
    ],
    [
      "2",
      "Algoritmo exacto de fuerza bruta, parámetro mínimo n, por qué es costoso",
      "Informe sec. 4 y sec. 6; §sat.py§ → §fuerza_bruta§",
      "3",
    ],
    [
      "3",
      "Pruebas desde n=10 hasta superar 1 hora, tiempos registrados, condiciones controladas",
      "Informe sec. 7 y sec. 8.1; §experimentos.py§; §hardware.txt§",
      "4, 5",
    ],
    [
      "4",
      "Gráfico n vs. tiempo, interpretación y relación con la complejidad teórica",
      "Informe sec. 8.1 (Gráfico 1); §graficos/1_tiempo_fuerza_bruta.png§",
      "5",
    ],
    [
      "5",
      "Heurística (búsqueda local): funcionamiento, ventajas y limitaciones",
      "Informe sec. 5; §sat.py§ → §walksat§",
      "6",
    ],
    [
      "6",
      "Mismos casos del punto 3: calidad, error relativo y tiempo",
      "Informe sec. 8.2, sec. 9; §tabla_comparacion.md§",
      "7, 8",
    ],
    [
      "7",
      "Heurística con n grande (n ≥ 1000): escalabilidad y viabilidad",
      "Informe sec. 8.4; §resultados_escala.csv§",
      "9",
    ],
    [
      "8",
      "Aplicación concreta en una empresa o institución de la región",
      "Informe sec. 10 y discusión ampliada",
      "10, 11, 12, 13",
    ],
    [
      "9",
      "Informe: introducción, algoritmos, resultados, conclusiones y reflexión crítica",
      "Informe completo (sec. 1 a 12, Referencias)",
      "14, 15, 16",
    ],
  ],
  [450, 4000, 3200, 1376],
);

H2("0.3 Qué es cada archivo del repositorio y en qué orden se corre");
TABLA(
  ["Archivo / carpeta", "Qué es", "Quién lo produce"],
  [
    [
      "§sat.py§",
      "El corazón: generadores de fórmulas, lector/escritor DIMACS, verificador, §fuerza_bruta§ y §walksat§. Se importa desde los demás. Corrido solo (§python sat.py§) hace una demo mínima.",
      "Escrito (por la IA)",
    ],
    [
      "§experimentos.py§",
      "Tres experimentos: §exacto§ (fuerza bruta vs. WalkSAT, n=10 a 32), §escala§ (WalkSAT con n grande) y §densidad§ (variar m/n con n=16). Guarda CSV, instancias y §hardware.txt§.",
      "Escrito",
    ],
    [
      "§graficos.py§",
      "Lee los CSV y genera los 6 gráficos en §graficos/§ y la tabla §tabla_comparacion.md§.",
      "Escrito",
    ],
    [
      "§resultados_exacto.csv§",
      "24 filas (12 tamaños × 2 tipos): tiempos del exacto, éxitos de WalkSAT sobre 10, error relativo, etc.",
      "§experimentos.py exacto§",
    ],
    [
      "§resultados_escala.csv§",
      "30 filas: WalkSAT con n=100 a 10.000 (3 semillas × 2 tipos × 5 tamaños).",
      "§experimentos.py escala§",
    ],
    [
      "§resultados_densidad.csv§",
      "6 filas: para cada m/n de 2 a 6, cuántas de 30 fórmulas son satisfacibles y cuántas halló WalkSAT.",
      "§experimentos.py densidad§",
    ],
    [
      "§historial_n500.csv§",
      "Cláusulas insatisfechas en cada paso de una corrida de WalkSAT (n=500). Alimenta el gráfico 6.",
      "§experimentos.py escala§",
    ],
    [
      "§hardware.txt§",
      "CPU, RAM, SO, Python, LIMITE y fecha del equipo de medición.",
      "§experimentos.py exacto§",
    ],
    [
      "§instancias/§",
      "Las 24 fórmulas del experimento exacto en formato DIMACS (§planted_nN_mM.cnf§ e §insat_nN_mM.cnf§).",
      "§experimentos.py exacto§",
    ],
    ["§graficos/§", "Gráficos 1 a 6 (PNG).", "§graficos.py§"],
    [
      "§tabla_comparacion.md§",
      "Tabla completa exacto vs. WalkSAT para pegar en el informe.",
      "§graficos.py§",
    ],
    [
      "§generadores/§",
      "Scripts de node que arman los documentos leyendo los CSV: §informe.js§, §presentacion.js§, §guion.js§, §estudio.js§ (esta guía) y §datos.js§ (lee los CSV y calcula las cifras).",
      "Escritos",
    ],
    [
      "§SAT_Informe.docx§, §SAT_Presentacion.pptx§, §SAT_Guion_Exposicion.docx§",
      "Los entregables. Salen de los generadores, así que las cifras nunca están copiadas a mano.",
      "Generadores",
    ],
  ],
  [2600, 5000, 1426],
);
H3("Orden de ejecución (de cero)");
CODE(`
pip install -r requirements.txt                       # solo matplotlib
LIMITE=3600 python experimentos.py exacto 10 12 14 16 18 20 22 24 26 28 30 32
python experimentos.py escala
python experimentos.py densidad                       # opcional (n=16)
python graficos.py                                    # gráficos + tabla
node generadores/informe.js                           # regenera el informe
node generadores/presentacion.js                      # regenera las diapositivas
node generadores/guion.js                             # regenera el guion
node generadores/estudio.js                           # regenera esta guía`);
CAJA("OJO", [
  "§LIMITE§ se lee de una variable de entorno y por defecto vale **300 s**, no 3600. Si corren §experimentos.py exacto§ sin poner §LIMITE=3600§, el límite será de 5 minutos y las fórmulas de n=28 en adelante darán TIMEOUT. Los datos del TP se obtuvieron con §LIMITE=3600§ (lo confirma §hardware.txt§).",
  "El exacto completo (n=10 a 32) tardó unas **1 h 50 min** [DERIVADO: suma de los tiempos del CSV; §hardware.txt§ marca el inicio a las 19:45:26 y el CSV se terminó de escribir a las 21:37].",
]);

// ---- fin chunk 1
// =====================================================================
// PARTE 1
// =====================================================================
H1("Parte 1 · Fundamentos que los dos tienen que dominar");
P(
  "Esta parte explica desde cero todo lo teórico que aparece en el TP. Si dominan esto, pueden contestar casi cualquier pregunta conceptual. Los dos deben poder explicar cada punto con sus palabras, sin leer.",
);

H2("1.1 Problema de decisión");
P(
  "Un **problema de decisión** es una pregunta con respuesta sí o no sobre una entrada. Ejemplos: «¿este número es primo?», «¿este grafo se puede colorear con 3 colores?», «¿esta fórmula lógica se puede hacer verdadera?». Es importante porque la teoría de la complejidad (P, NP, NP-completo) está definida para problemas de decisión. Los problemas de optimización («¿cuál es el mejor coloreo?») se pasan a decisión preguntando «¿existe uno con al menos tal calidad?».",
);
P(
  "SAT es un problema de decisión: **entrada** = una fórmula booleana; **pregunta** = ¿existe una asignación de verdadero/falso a sus variables que la haga verdadera? Respuesta: SÍ (satisfacible, SAT) o NO (insatisfacible, UNSAT).",
);
CAJA(
  "MEMORIA",
  "Decidir ≠ encontrar. Decidir es responder sí/no. Encontrar es dar la asignación. Si se sabe decidir rápido, encontrar la solución también se puede hacer rápido (probando variable por variable); el TP trabaja con las dos cosas, pero la teoría habla de decidir.",
);

H2("1.2 P, NP, NP-difícil y NP-completo");
TABLA(
  ["Clase", "Definición", "Ejemplo"],
  [
    [
      "**P**",
      "Problemas que una máquina de Turing determinística resuelve en tiempo polinomial en el tamaño de la entrada.",
      "Ordenar, camino mínimo, 2-SAT.",
    ],
    [
      "**NP**",
      "Problemas cuya solución propuesta (un «certificado») se puede **verificar** en tiempo polinomial. Equivalente: los que resuelve una máquina de Turing **no determinística** en tiempo polinomial.",
      "SAT, 3-SAT, coloreo, clique, viajante (versión de decisión).",
    ],
    [
      "**NP-difícil**",
      "Todo problema de NP se reduce a él en tiempo polinomial. Es «al menos tan difícil como cualquiera de NP». Puede no estar en NP.",
      "Max-SAT (optimización), problema de la parada.",
    ],
    [
      "**NP-completo**",
      "Está en NP **y** es NP-difícil. Son los más difíciles de NP.",
      "SAT, 3-SAT, coloreo con k ≥ 3 colores, subset-sum.",
    ],
  ],
  [1500, 5000, 2526],
);
P(
  "**Qué se sabe y qué no.** P ⊆ NP (si puedo resolver en tiempo polinomial, también puedo verificar). Lo que **nadie sabe** es si P = NP: es un problema abierto (uno de los siete del Clay Institute, con premio de un millón de dólares). Lo que se cree, pero no está demostrado, es que P ≠ NP. Nadie demostró que SAT necesite tiempo exponencial: solo se conocen algoritmos exponenciales y ninguno polinomial.",
);
CAJA(
  "OJO",
  "Error muy común: decir «NP significa no polinomial». Es FALSO. NP = «Nondeterministic Polynomial time». Los problemas de P también están en NP. Es una trampa clásica en las orales.",
);
H3("¿Qué pasaría si P = NP?");
P(
  "Si alguien encontrara un algoritmo polinomial para SAT (o para cualquier problema NP-completo), entonces **todos** los problemas de NP se resolverían en tiempo polinomial, porque se reducen a ese. Consecuencias: se podrían resolver eficientemente muchos problemas de optimización, planificación y verificación hoy intratables, y se rompería buena parte de la criptografía de clave pública, que se apoya en que ciertos problemas son difíciles. La mayoría de los especialistas cree que P ≠ NP, pero no hay demostración. En este TP: la fuerza bruta crece exponencialmente, pero **eso no prueba nada sobre P vs NP** (mide un algoritmo concreto, no todos los posibles).",
);

H2("1.3 Reducciones");
P(
  "Una **reducción polinomial** (de Karp) de un problema A a un problema B es una función que convierte, en tiempo polinomial, cualquier entrada x de A en una entrada f(x) de B tal que: la respuesta de A sobre x es SÍ **si y solo si** la respuesta de B sobre f(x) es SÍ. Se escribe A ≤ₚ B. Idea: «si sé resolver B, resuelvo A traduciendo».",
);
P("Dos usos de las reducciones:");
BS(
  "**Para demostrar que algo es difícil:** si A es NP-difícil y A ≤ₚ B, entonces B es NP-difícil. Así se demuestra que 3-SAT, coloreo, etc. son NP-completos: se parte de un problema ya conocido como NP-completo.",
  "**Para resolver algo con una herramienta existente:** traducir mi problema a SAT y usar un solver. Es lo que se propone con el coloreo de grafos y los horarios (diapositivas 10 y 12).",
);
CAJA(
  "EJEMPLO",
  "**Coloreo a SAT (resumen; el detalle está en la Parte 6).** ¿Se puede colorear un grafo con k colores? Variables x(v,c) = «el vértice v lleva el color c». Cláusulas: cada vértice tiene al menos un color; cada vértice tiene a lo sumo un color; dos vértices vecinos no comparten color. La fórmula es satisfacible si y solo si el grafo es k-coloreable, y su tamaño es polinomial en el grafo. Eso es una reducción polinomial.",
);

H2("1.4 El teorema de Cook-Levin");
P(
  "**Enunciado:** SAT es NP-completo (Cook 1971; Levin 1973, de forma independiente). Fue el **primer** problema que se demostró NP-completo; todos los demás se demuestran después por reducción desde SAT (o desde otro ya demostrado).",
);
P(
  "**Idea de la demostración** (para poder explicarla sin todos los detalles técnicos):",
);
BS(
  "SAT está en NP: dada una asignación, se verifica en tiempo polinomial recorriendo las cláusulas.",
  "SAT es NP-difícil: sea L cualquier lenguaje de NP. Existe una máquina de Turing **no determinística** M que lo decide en tiempo polinomial p(n). Dada una entrada w, se construye una fórmula F_w que «describe» una ejecución aceptadora de M sobre w.",
  "La ejecución de M se representa como una tabla (un «tableau») de p(n) × p(n) celdas: cada fila es la configuración de la cinta en un instante, y hay p(n) instantes. Se crean variables booleanas del tipo «la celda (i,j) contiene el símbolo s».",
  "Se escriben cláusulas que dicen: cada celda tiene exactamente un símbolo; la primera fila es la configuración inicial con la entrada w; cada fila se obtiene de la anterior por una transición válida de M (esto es una condición local, sobre ventanas chicas de celdas); y en alguna fila aparece el estado de aceptación.",
  "F_w tiene tamaño polinomial (≈ p(n)² celdas por una cantidad constante de símbolos), se construye en tiempo polinomial y es satisfacible **si y solo si** M acepta w. Entonces L ≤ₚ SAT.",
);
P(
  "**Qué hay que llevarse:** el no determinismo («existe una elección de transiciones que acepta») se codifica como «existe una asignación de variables que cumple todo». Por eso SAT captura exactamente la idea de «existe un certificado». Karp (1972) usó SAT para demostrar que otros 21 problemas combinatorios son NP-completos, y desde entonces se demostró para miles más.",
);

H2("1.5 Por qué SAT importa");
BS(
  "**Teóricamente:** es el problema NP-completo de referencia. Cualquier avance sobre SAT es un avance sobre toda NP.",
  "**Prácticamente:** muchos problemas se pueden traducir a SAT y resolverlos con solvers modernos (CDCL), que manejan instancias con millones de variables cuando tienen estructura. Usos documentados en el informe: verificación de hardware y software (bounded model checking), planificación como satisfacibilidad, resolución de dependencias de paquetes (Eclipse con Sat4j; libsolv) y horarios con SAT/MaxSAT [LITERATURA; las referencias a abrir están en la Parte 7].",
  "**Paradoja interesante:** en el peor caso es intratable, pero en la práctica los solvers resuelven instancias enormes. La dificultad depende de la estructura de la instancia. Esa es la idea de fondo de la reflexión crítica (diapositiva 14).",
);

H2("1.6 Fórmulas booleanas, literales, cláusulas y CNF");
P(
  "Una **variable booleana** vale verdadero (V, 1) o falso (F, 0). Un **literal** es una variable (x₃) o su negación (¬x₃). Una **cláusula** es una disyunción (OR, ∨) de literales: se cumple si **al menos uno** de sus literales es verdadero. Una fórmula en **CNF** (forma normal conjuntiva) es una conjunción (AND, ∧) de cláusulas: se cumple si **todas** las cláusulas se cumplen a la vez.",
);
P(
  "Ejemplo del enunciado (4 variables, 3 cláusulas), que aparece en la diapositiva 2 y en §sat.py§:",
);
CODE(`
F = (x1 ∨ ¬x2 ∨ x3) ∧ (¬x1 ∨ x2 ∨ x4) ∧ (x2 ∨ ¬x3 ∨ ¬x4)

en el código:  [[1, -2, 3], [-1, 2, 4], [2, -3, -4]]
               (3 significa x3,  -3 significa ¬x3)`);
P(
  "**Verificarlo a mano.** Probemos la asignación con todas las variables falsas (x1=F, x2=F, x3=F, x4=F):",
);
BS(
  "Cláusula 1: x1=F, ¬x2=V (porque x2 es F), x3=F → hay un literal verdadero → **se cumple**.",
  "Cláusula 2: ¬x1=V → **se cumple**.",
  "Cláusula 3: ¬x3=V → **se cumple**. Las tres se cumplen, entonces la fórmula es **satisfacible** y esa es una solución. De hecho es la primera que prueba la fuerza bruta (el número 0): al correr §python sat.py§ devuelve «SAT» tras evaluar una sola asignación.",
);
P(
  "**¿Por qué CNF?** Porque cualquier fórmula booleana se puede convertir a CNF manteniendo la satisfacibilidad (transformación de Tseitin: agrega variables auxiliares y crece solo linealmente; en cambio, convertir aplicando la distributiva puede explotar exponencialmente). Y porque los solvers están diseñados para CNF. Contraste: SAT para fórmulas en **DNF** (disyunción de conjunciones) es trivial, porque alcanza con ver si alguna conjunción no contiene un literal y su negación. Por eso lo difícil es CNF.",
);

H2("1.7 k-SAT y 3-SAT");
BS(
  "**SAT:** las cláusulas pueden tener cualquier cantidad de literales.",
  "**k-SAT:** todas las cláusulas tienen exactamente k literales.",
  "**2-SAT** está en P (se resuelve en tiempo lineal con el grafo de implicaciones y componentes fuertemente conexas). **3-SAT** ya es NP-completo. Por eso 3 es el salto de «fácil» a «difícil».",
  "**3-SAT es NP-completo.** (a) Está en NP porque es un caso particular de SAT. (b) Es NP-difícil porque SAT ≤ₚ 3-SAT: cada cláusula larga se parte con variables nuevas.",
);
CODE(`
Cláusula larga:     (a ∨ b ∨ c ∨ d)
Se reemplaza por:   (a ∨ b ∨ y) ∧ (¬y ∨ c ∨ d)        y = variable nueva

  - Si la original se cumple por a o b:  tomo y = F  -> (a∨b∨y) vale, (¬y∨c∨d) vale por ¬y.
  - Si la original se cumple por c o d:  tomo y = V  -> (a∨b∨y) vale por y, (¬y∨c∨d) vale por c o d.
  - Recíproco: si las dos nuevas valen y y = V, de la 2ª sale (c∨d); si y = F, de la 1ª sale (a∨b).
    En ambos casos la original vale.

Cláusula de k literales  ->  k-2 cláusulas de 3 literales y k-3 variables nuevas.
Cláusulas cortas:  (a)  se completa con  (a∨y∨z)(a∨y∨¬z)(a∨¬y∨z)(a∨¬y∨¬z).`);
P(
  "El informe (sec. 2.5) lo explica igual. Lo importante si preguntan: **la transformación es polinomial y preserva la satisfacibilidad**.",
);

H2("1.8 SAT vs. UNSAT, certificados y por qué importa para WalkSAT");
P(
  "Una fórmula es **satisfacible (SAT)** si existe una asignación que la cumple, e **insatisfacible (UNSAT)** si no existe ninguna. Hay una asimetría fundamental:",
);
TABLA(
  ["", "Respuesta SAT", "Respuesta UNSAT"],
  [
    [
      "**Certificado**",
      "Una asignación (n bits). Se verifica en tiempo polinomial: O(m·k).",
      "No se conoce un certificado corto en general. UNSAT está en coNP, y no se sabe si NP = coNP.",
    ],
    [
      "**Cómo se prueba**",
      "Alcanza con mostrar la asignación.",
      "Hay que cubrir todo el espacio de búsqueda (o producir una prueba, por ejemplo por resolución). Solo un algoritmo **completo** puede hacerlo.",
    ],
    [
      "**Lo que hace WalkSAT**",
      "Puede encontrarla y confirmarla.",
      "No puede probarlo. Si no encuentra solución devuelve UNKNOWN.",
    ],
  ],
  [1700, 3400, 3926],
);
CAJA(
  "MEMORIA",
  "UNKNOWN ≠ UNSAT. Es la frase central del TP (diapositivas 6 y 15). WalkSAT dice «no encontré»; solo el exacto puede decir «no hay».",
);

H2("1.9 DIMACS");
P(
  "**DIMACS CNF** es el formato de texto estándar para fórmulas CNF que leen los solvers (MiniSat, Kissat, etc.). Las reglas:",
);
BS(
  "Las líneas que empiezan con §c§ son comentarios.",
  "La línea §p cnf n m§ declara n variables y m cláusulas.",
  "Cada cláusula es una línea de enteros distintos de cero (positivo = variable, negativo = negada) y termina en §0§.",
);
CODE(`
c ejemplo del enunciado
p cnf 4 3
1 -2 3 0
-1 2 4 0
2 -3 -4 0`);
P(
  "El TP guarda las 24 fórmulas del experimento exacto en DIMACS (carpeta §instancias/§). Sirve para dos cosas: que **los dos algoritmos resuelvan exactamente la misma fórmula** y que cualquiera pueda repetir los resultados o probar las fórmulas con un solver externo.",
);

H2("1.10 Fuerza bruta y por qué 2ⁿ");
P(
  "**Fuerza bruta** (búsqueda exhaustiva): probar todas las asignaciones posibles. Cada una de las n variables tiene 2 valores posibles, independientes entre sí, así que hay 2 × 2 × … × 2 = **2ⁿ** asignaciones. Se duplican con cada variable nueva.",
);
TABLA(
  ["n", "2ⁿ asignaciones", "Comentario"],
  [
    ["10", "1.024", "Instantáneo."],
    [
      "20",
      "1.048.576",
      "Un millón: en el TP, 1,68 s en una insatisfacible [MEDIDO].",
    ],
    ["30", "1.073.741.824", "Mil millones: 1.768 s (29,5 min) [MEDIDO]."],
    ["32", "4.294.967.296", "Cuatro mil millones: no terminó en 1 h [MEDIDO]."],
    [
      "100",
      "≈ 1,27 × 10³⁰",
      "Aunque se evaluara una asignación por nanosegundo, tardaría ≈ 4 × 10¹³ años [DERIVADO].",
    ],
  ],
  [800, 2600, 5626],
);
P(
  "**Costo por asignación.** Para cada asignación hay que revisar hasta m cláusulas de k literales: O(m·k). Total en el peor caso: **O(2ⁿ · m · k)**. Espacio: O(m·k) para guardar la fórmula más O(n) para la asignación en curso (el código la representa con un único entero).",
);
CAJA(
  "OJO",
  "El O(2ⁿ·m·k) es una cota del PEOR caso. En las corridas reales el código evalúa en promedio solo ≈ 7,9 cláusulas por asignación en las insatisfacibles (ver sección 2.6), porque una asignación al azar rompe una cláusula aleatoria con probabilidad 1/8 y entonces el algoritmo corta rápido. Eso no cambia el 2ⁿ, solo la constante.",
);

H2("1.11 Heurística, búsqueda local, completitud e incompletitud");
BS(
  "**Algoritmo exacto (completo):** si existe solución, la encuentra; si no existe, lo demuestra. Siempre responde correctamente (dado tiempo suficiente).",
  "**Heurística:** un método que usa una regla práctica para buscar buenas soluciones rápido, sin garantías de encontrarlas. «Heurística» no significa «mala»: significa «sin garantía».",
  "**Búsqueda local:** se parte de una solución candidata completa (aquí, una asignación entera) y se la va mejorando con cambios chicos (aquí, dar vuelta una variable), mirando solo los «vecinos». No recorre todo el espacio: camina por él.",
  "**Completo vs. incompleto:** un algoritmo es **completo** si es capaz de responder en todos los casos (incluyendo UNSAT). WalkSAT es **incompleto**: puede no encontrar una solución que existe y nunca puede probar que no hay.",
  "**Correcto (sound) vs. completo:** WalkSAT es correcto en lo que afirma (cuando dice SAT, la asignación es solución) pero incompleto (puede no decir nada). Esa distinción responde «¿puede WalkSAT dar un falso SAT?» (ver Parte 6).",
  "**Mínimo local:** una asignación donde cualquier cambio chico empeora las cosas pero que no es solución. Es la razón de agregar ruido: con probabilidad p se hace un movimiento al azar para salir del pozo.",
);

H2("1.12 El umbral de fase en 3-SAT aleatorio (m/n ≈ 4,26)");
CAJA(
  "HONESTIDAD",
  "El 4,26 es un resultado conocido de la LITERATURA, no algo que se probó en este TP. Lo que el TP sí midió es cuántas fórmulas de n=16 son satisfacibles para distintos m/n (informe sec. 8.3). No se puede decir «verificamos que el umbral es 4,26».",
);
P(
  "Tomemos fórmulas 3-SAT **aleatorias** con n variables y m = α·n cláusulas, donde α = m/n es la «densidad». Con pocas cláusulas por variable (α chico) casi todas las fórmulas son satisfacibles: hay muchísimas soluciones. Con muchas cláusulas (α grande) casi todas son insatisfacibles: las restricciones se contradicen. En el medio hay una **transición de fase**: al aumentar n, el pasaje de «casi siempre SAT» a «casi siempre UNSAT» se vuelve abrupto y ocurre cerca de **α ≈ 4,26**. [LITERATURA: Mitchell, Selman y Levesque 1992; Crawford y Auton 1993; Kirkpatrick y Selman 1994; existencia de un umbral agudo: Friedgut 1999; valor ≈ 4,267 por física estadística. VERIFICAR antes de citar con precisión.] Cerca del umbral las instancias son, en promedio, las más difíciles, porque están «al borde»: ni sobran soluciones ni es evidente la contradicción.",
);
P("**Qué midió el TP al respecto** [MEDIDO, n=16, 30 fórmulas por densidad]:");
TABLA(
  ["m/n", "m", "Satisfacibles (de 30)", "WalkSAT encontró", "Falsos SAT"],
  d.filas.map((r) => [
    fn(r.razon_m_n, 2),
    String(r.m),
    String(r.satisfacibles_segun_exacto),
    String(r.walksat_encontro_solucion),
    String(r.walksat_dijo_SAT_en_UNSAT),
  ]),
  [1400, 1400, 2300, 2100, 1826],
);
P(
  `La fracción de satisfacibles cae de 30/30 a 1/30 al subir la densidad: esa es la tendencia. Pero con n=16 la transición todavía es «suave» (a 4,26 hay ${d.filas[3].satisfacibles_segun_exacto} de 30, o sea 60 %). Para ver el umbral nítido en 4,26 hacen falta n mucho más grandes. Total: ${d.satTotal} satisfacibles de ${d.instancias}; WalkSAT halló ${d.walkTotal}; falsos SAT: ${d.falsos}.`,
);
P(
  "**Cómo se usa en el TP:** m = round(4,26·n) en la comparación exacto vs. WalkSAT (la zona difícil «de manual»). Pero ojo: las fórmulas de esa comparación no son aleatorias puras (son planted o llevan un núcleo insatisfacible), así que lo de «zona difícil» aplica con reservas. En la escala se usó 4,26 para las planted y **4,0** para las aleatorias (por debajo del umbral: probablemente satisfacibles, pero el TP no lo sabe con certeza).",
);

H2("1.13 Instancias planted e insatisfacibles construidas");
P(
  "Para medir algoritmos hay que conocer la respuesta correcta de antemano o poder calcularla. El TP usa tres tipos de fórmulas:",
);
TABLA(
  ["Tipo", "Cómo se genera", "Qué garantiza", "Para qué sirve", "Trampa"],
  [
    [
      "**Aleatoria**",
      "Cada cláusula: 3 variables distintas al azar, cada una con signo al azar.",
      "Nada: puede ser SAT o UNSAT.",
      "Escala (m/n=4,0) y densidad.",
      "En la escala no se sabe si es SAT: un UNKNOWN no prueba nada.",
    ],
    [
      "**Planted**",
      "Se elige una solución oculta; se genera cada cláusula al azar y, si la solución no la cumple, se da vuelta un literal.",
      "SAT garantizada (la solución oculta).",
      "Medir qué tan rápido se encuentra una solución.",
      "Suele ser más fácil que una aleatoria: sesga hacia la solución oculta.",
    ],
    [
      "**Insatisfacible**",
      "8 cláusulas con todas las combinaciones de signos sobre 3 variables + cláusulas aleatorias hasta llegar a m.",
      "UNSAT garantizada.",
      "Forzar al exacto a recorrer las 2ⁿ asignaciones.",
      "Tiene un núcleo insatisfacible chico y obvio; no es una instancia «difícil» típica.",
    ],
  ],
  [1300, 2500, 1500, 1900, 1826],
  { size: 18 },
);

H2("1.14 Max-SAT y el «error relativo»");
P(
  "**Max-SAT:** dada una fórmula CNF, encontrar la asignación que cumple la **mayor cantidad posible** de cláusulas. Es un problema de **optimización** NP-difícil. Si la fórmula es satisfacible, el óptimo es m (todas). Si no, el óptimo es menor que m.",
);
P(
  "En el TP se usa Max-SAT solo como **vara de medir**: como SAT es de decisión (cumplir el 99 % de las cláusulas no es una respuesta correcta), se define §error_relativo = (cláusulas insatisfechas de la mejor asignación de WalkSAT) / m§. Vale 0 en las satisfacibles resueltas y crece cuanto más lejos queda. Es una medida de **cercanía**, no de corrección. Ver Parte 3 y Parte 5.",
);
P(
  "[LITERATURA] Dato de color: una asignación al azar cumple en promedio 7/8 de las cláusulas de un 3-SAT (cada cláusula falla con prob. 1/8), y se sabe que no se puede garantizar más de 7/8 en tiempo polinomial salvo que P = NP (Håstad 2001). VERIFICAR antes de usarlo en voz alta.",
);
// =====================================================================
// PARTE 2
// =====================================================================
H1("Parte 2 · Cómo está hecho el código");
P(
  "Recorrido guiado de los tres archivos de Python. Abran cada archivo al lado y sigan con el dedo. Para cada función: qué hace, pseudocódigo, un ejemplo chico a mano y qué cosas raras tiene.",
);
CAJA(
  "HONESTIDAD",
  "El código lo escribió una IA y no tiene tests automáticos. La confianza en que es correcto viene de cuatro evidencias (sección 2.12), no de una suite de pruebas. Si el profesor pregunta «¿cómo saben que está bien?», contesten con esas evidencias.",
);

H2("2.1 Cómo se representa todo");
BS(
  "**Fórmula:** lista de cláusulas; cada cláusula es una lista de enteros con signo. §3§ es x₃ y §-3§ es ¬x₃. Ejemplo: §[[1, -2, 3], [-1, 2, 4], [2, -3, -4]]§.",
  "**Asignación en WalkSAT y en los verificadores:** lista §asig§ de largo n+1 donde §asig[i]§ es True/False para la variable i. La posición 0 no se usa (vale None). Se hace así para que el índice coincida con el número de variable, que empieza en 1.",
  "**Asignación en la fuerza bruta:** un entero §bits§. El bit número (i−1) vale 1 si x_i es verdadera. Se extrae con §(bits >> (i-1)) & 1§. Recorrer todos los enteros de 0 a 2ⁿ−1 recorre todas las asignaciones.",
  "**Evaluar un literal §lit§ bajo una asignación:** el literal es verdadero si §(lit > 0) == valor_de_la_variable§. Es decir: literal positivo y variable verdadera, o literal negativo y variable falsa.",
);
CAJA("EJEMPLO", [
  "Decodificar §bits = 5§ con n=4. En binario 5 = 0101 (el bit 0 es el de más a la derecha).",
  "bit 0 = 1 → x1 = V;  bit 1 = 0 → x2 = F;  bit 2 = 1 → x3 = V;  bit 3 = 0 → x4 = F.",
  "Entonces el entero 5 es la asignación (V, F, V, F). Como x1 es el bit menos significativo, **x1 cambia en cada paso del contador y x_n es la que cambia más lento**. Esto importa para entender por qué el tiempo de las planted no es monótono (sección 5.4).",
]);

H2("2.2 §generar_aleatoria(n, m, k=3, semilla=0)§");
CODE(`
azar = Random(semilla)
repetir m veces:
    variables = k variables DISTINTAS elegidas al azar de 1..n
    cláusula  = cada variable con signo + o - con probabilidad 1/2
agregar la cláusula a la fórmula`);
BS(
  "§random.Random(semilla)§ crea un generador propio con semilla fija: **misma semilla → misma fórmula**, siempre. Eso da reproducibilidad.",
  "§azar.sample(range(1, n+1), k)§ elige k variables **sin repetición**, así una cláusula nunca tiene la misma variable dos veces (ni x ∨ ¬x, que sería trivialmente verdadera).",
  "§v if azar.random() < 0.5 else -v§ pone el signo: cara o ceca.",
  "No garantiza nada sobre la satisfacibilidad. Con m/n cerca de 4,26 salen mezcladas.",
);

H2("2.3 §generar_planted(n, m, k=3, semilla=0)§: satisfacible garantizada");
CODE(`
solucion = asignación oculta al azar (n valores V/F)
repetir m veces:
    cláusula = generar como en aleatoria
    si la solución oculta NO cumple la cláusula (ninguno de sus literales es verdadero):
        dar vuelta el signo de UN literal elegido al azar   # ahora la solución sí la cumple
    agregar cláusula
devolver (fórmula, solucion)`);
BS(
  "**Por qué es satisfacible:** cada cláusula quedó cumplida por §solucion§, y entonces §solucion§ cumple todas. La función devuelve también la solución oculta, aunque los experimentos no la usan.",
  "**Por qué suele ser fácil:** la corrección no es neutral. Una cláusula al azar la cumple la solución oculta con prob. 7/8; en el 1/8 restante se «arregla» a la fuerza. El resultado es que las fórmulas están sesgadas hacia la solución oculta (más literales «alineados» con ella que en una fórmula aleatoria). Esto hace que WalkSAT las encuentre en decenas de pasos. En el informe se avisa: «las planted resultaron fáciles».",
  "**Detalle:** las primeras n llamadas a §azar.random()§ generan la solución oculta. Esto es la clave de la trampa de las semillas (sección 2.11).",
);

H2("2.4 §generar_insatisfacible(n, m, semilla=0)§: insatisfacible garantizada");
CODE(`
tres = 3 variables distintas al azar
formula = las 8 cláusulas que salen de TODAS las combinaciones de signos sobre esas 3 variables
formula += (m - 8) cláusulas aleatorias   (con una semilla distinta: semilla + 7919)
mezclar el orden de las cláusulas`);
P(
  "**Por qué esas 8 cláusulas obligan a UNSAT.** Llamemos a, b, c a las tres variables. Las 8 cláusulas son:",
);
CODE(`
(a ∨ b ∨ c)   (a ∨ b ∨ ¬c)   (a ∨ ¬b ∨ c)   (a ∨ ¬b ∨ ¬c)
(¬a ∨ b ∨ c)  (¬a ∨ b ∨ ¬c)  (¬a ∨ ¬b ∨ c)  (¬a ∨ ¬b ∨ ¬c)`);
P(
  "Una cláusula (l₁ ∨ l₂ ∨ l₃) solo es falsa cuando **los tres literales son falsos**. Para cada una de las 8 asignaciones posibles de (a, b, c) hay **exactamente una** cláusula, entre las 8, cuyos tres literales son falsos: la que tiene los signos opuestos a la asignación. Ejemplo: con a=V, b=F, c=V, la única cláusula falsa es (¬a ∨ b ∨ ¬c). Entonces **toda** asignación viola al menos una cláusula, y la fórmula completa es insatisfacible sin importar qué otras cláusulas se agreguen.",
);
BS(
  "En el código: §mascara§ va de 0 a 7; el bit i de la máscara decide si la variable i-ésima va positiva o negada: §tres[i] if (mascara >> i) & 1 else -tres[i]§.",
  "Las §m−8§ cláusulas restantes son aleatorias. La semilla §semilla + 7919§ es una forma de **no reutilizar la misma semilla** que la usada para elegir las 3 variables (7919 es el primo número 1000; no tiene ningún significado teórico). Si se reutilizara la misma, las cláusulas aleatorias estarían correlacionadas con la elección de §tres§.",
  "§azar.shuffle(formula)§ mezcla las cláusulas para que el núcleo insatisfacible no quede al principio. Importa para WalkSAT y para la fuerza bruta: sin mezclar, el núcleo estaría siempre al principio del recorrido y la fuerza bruta descartaría casi todas las asignaciones rápido (corta en la primera cláusula falsa).",
  "**Consecuencia sobre el error relativo:** como cualquier asignación viola al menos 1 cláusula, el mejor valor posible es 1 insatisfecha. Eso explica que en TODAS las insatisfacibles §error_relativo = 1/m§ (por ejemplo 1/43 = 0,02326 en n=10). WalkSAT llegó al óptimo de Max-SAT, solo que el programa no lo puede afirmar (por eso §n.a.§).",
  "**Cuidado al hablar:** el núcleo de 8 cláusulas es una estructura obvia para un solver moderno (CDCL lo descubriría enseguida, aprendiendo cláusulas). La fuerza bruta no usa estructura, por eso tarda 2ⁿ. Es un buen ejemplo de por qué los tiempos del exacto miden la ingenuidad del algoritmo, no la dificultad intrínseca de la fórmula. [Esto último es mi razonamiento, no una medición del TP.]",
);

H2("2.5 §escribir_dimacs§ y §leer_dimacs§");
BS(
  "§escribir_dimacs(ruta, formula, n, comentarios)§: escribe las líneas §c ...§ de comentario, la línea §p cnf n m§ y una línea por cláusula terminada en §0§.",
  "§leer_dimacs(ruta)§: salta líneas vacías y comentarios; en la línea §p§ lee n; para el resto, acumula enteros hasta ver un §0§ y entonces cierra la cláusula. Devuelve §(formula, n)§. Soporta cláusulas partidas en varias líneas (como el formato estándar permite).",
  "Los experimentos **escriben** las instancias del exacto pero no las vuelven a **leer** (generan y guardan). §leer_dimacs§ está para usar benchmarks externos. Yo verifiqué que leer los 24 archivos de §instancias/§ y compararlos con regenerar las fórmulas con las mismas semillas da fórmulas idénticas.",
);

H2(
  "2.6 §contar_insatisfechas(formula, asignacion)§ y §fuerza_bruta(formula, n, limite_s)§",
);
P(
  "§contar_insatisfechas§ recorre las cláusulas y cuenta las que no tienen ningún literal verdadero. Es un verificador independiente. **Ninguno de los experimentos lo llama**; solo lo usa el control aparte §verificar_modelos.py§ (ver sección 2.12).",
);
H3("fuerza_bruta, línea por línea");
CODE(`
inicio = reloj()
para bits desde 0 hasta 2^n - 1:
    asignaciones += 1
    cumple_todas = verdadero
    para cada cláusula cl de la fórmula:
        clausulas += 1
        cumple = falso
        para cada literal lit de cl:
            valor = bit número (|lit|-1) de bits
            si (lit > 0 y valor == 1) o (lit < 0 y valor == 0):
                cumple = verdadero ; cortar el for de literales
        si no cumple:
            cumple_todas = falso ; cortar el for de cláusulas   # esta asignación ya falló
    si cumple_todas:
        devolver SAT, modelo = bits convertido a lista V/F
    cada 16.384 asignaciones: si pasó más de limite_s -> devolver TIMEOUT
devolver UNSAT`);
BS(
  "**Tres cortes tempranos:** sale del bucle de literales apenas hay uno verdadero; sale del bucle de cláusulas apenas hay una falsa; y sale de todo apenas encuentra una asignación que cumple todo. El tercero es la razón de que las fórmulas satisfacibles tarden menos que las insatisfacibles.",
  "**Salida:** un diccionario con §estado§ (SAT, UNSAT o TIMEOUT), §modelo§, y los contadores §asignaciones§ y §clausulas§ (evaluadas) y §tiempo§. Los contadores son los que aparecen en el CSV y en el gráfico 5.",
  "**TIMEOUT:** el reloj se mira solo cuando §(asignaciones & 0x3FFF) == 0§, o sea cada 16.384 asignaciones, para no gastar tiempo en llamar al reloj todo el tiempo. Por eso el tiempo de corte real fue 3.600,02 s y no exactamente 3.600.",
  "§asignaciones§ es el número de asignaciones **probadas**. Si devuelve SAT, es la posición (contando desde 1) de la primera solución en el orden 0, 1, 2, … Si UNSAT, vale exactamente 2ⁿ. Eso permite comprobar en el CSV que cada UNSAT evaluó 2ⁿ (por ejemplo 1.073.741.824 en n=30).",
);
CAJA("EJEMPLO", [
  "**Ejemplo chico a mano.** Con la fórmula del enunciado y n=4: bits=0 es (F,F,F,F) y ya vimos que cumple todo. La función evalúa 3 cláusulas (una por una; cada una cumple con un solo literal) y devuelve SAT con §asignaciones = 1§ y §clausulas = 3§. Es lo que imprime §python sat.py§.",
  "**Ejemplo UNSAT chico.** Con las 8 cláusulas sobre 3 variables (n=3): cada una de las 8 asignaciones (0 a 7) falla en una cláusula distinta. Devuelve UNSAT con §asignaciones = 8 = 2³§.",
]);
H3(
  "Por qué 2ⁿ·m·k y por qué en la práctica el costo por asignación es ≈ 8 cláusulas",
);
P(
  "En el peor caso cada asignación revisa las m cláusulas (cumple casi todas) y cada cláusula k literales: O(m·k) por asignación, O(2ⁿ·m·k) en total. Pero en las fórmulas del TP casi todas las asignaciones fallan **rápido**: una cláusula 3-SAT es falsa para una asignación al azar con probabilidad (1/2)³ = 1/8, así que, en promedio, la primera cláusula falsa aparece en la 8.ª cláusula revisada (distribución geométrica). Con los datos del CSV:",
);
TABLA(
  [
    "Fórmula",
    "Cláusulas evaluadas",
    "Asignaciones",
    "Cláusulas por asignación",
  ],
  [
    [
      "insatisfacible n=30",
      fn(n30.clausulas_evaluadas),
      fn(n30.asignaciones),
      fn(n30.clausulas_evaluadas / n30.asignaciones, 2),
    ],
    [
      "planted n=32",
      fn(n32p.clausulas_evaluadas),
      fn(n32p.asignaciones),
      fn(n32p.clausulas_evaluadas / n32p.asignaciones, 2),
    ],
  ],
  [2300, 2400, 2200, 2126],
);
P(
  "[DERIVADO] Es decir ≈ 8 y no m = 128 o 136. Así que el costo real es ≈ 2ⁿ · 8 · (hasta 3 literales); el «m» de la fórmula teórica es el peor caso. No contradice nada: sigue siendo exponencial. Si el profesor pregunta por qué el tiempo parece no depender de m, es por esto.",
);
P(
  `**Velocidad medida:** n=30 insatisfacible: ${fn(n30.asignaciones / n30.t_exacto_s, 0)} asignaciones/s; n=32 hasta el corte: ${fn(velCorte, 0)} asignaciones/s [DERIVADO]. Python puro, un solo hilo. Un programa en C haría 100 o 1000 veces más por segundo, pero con la misma curva exponencial: solo se desplazaría el n al que se llega.`,
);

H2("2.7 §walksat§: parámetros y estructuras de datos");
P(
  "Firma: §walksat(formula, n, max_flips=100000, p=0.5, reinicios=1, semilla=0, limite_s=None, guardar_historial=False)§.",
);
TABLA(
  ["Parámetro", "Qué es", "Valor en los experimentos"],
  [
    [
      "§max_flips§",
      "Tope de pasos (cada paso = dar vuelta una variable) por reinicio.",
      "100.000 en la comparación; 5.000.000 en la escala; 20.000 en densidad; 200.000 en el historial.",
    ],
    [
      "§p§",
      "Probabilidad de «ruido»: hacer un movimiento aleatorio en lugar del mejor.",
      "0,5 siempre (no se ajustó).",
    ],
    [
      "§reinicios§",
      "Cuántas veces se vuelve a empezar desde una asignación nueva al azar.",
      "1 siempre (no se usó).",
    ],
    [
      "§semilla§",
      "Semilla del generador de WalkSAT (independiente de la de la fórmula).",
      "0 a 9 en la comparación; semilla+1000 en escala y densidad.",
    ],
    [
      "§limite_s§",
      "Tiempo máximo en segundos (opcional).",
      "Solo en escala: 120 s (no llegó a alcanzarse).",
    ],
    [
      "§guardar_historial§",
      "Si es True, anota cuántas insatisfechas hay en cada paso.",
      "Solo para el gráfico 6.",
    ],
  ],
  [1800, 4100, 3126],
);
H3("Las estructuras que hacen que cada paso sea barato");
BS(
  "**§aparece§:** diccionario §literal → lista de índices de cláusulas donde aparece§. Se arma una vez al principio. Sirve para no recorrer toda la fórmula en cada paso: al dar vuelta una variable solo hay que tocar las cláusulas donde figura.",
  "**§verdaderos[c]§:** para cada cláusula c, **cuántos de sus literales son verdaderos** bajo la asignación actual. Cláusula insatisfecha ⇔ §verdaderos[c] == 0§. Cláusula «frágil» ⇔ §verdaderos[c] == 1§ (depende de un solo literal).",
  "**§insat§:** la lista de las cláusulas insatisfechas, y **§pos§:** diccionario que guarda, para cada cláusula, su posición dentro de §insat§. Con ambos se puede (a) elegir una insatisfecha al azar en O(1) con §insat[randrange(len(insat))]§, (b) agregar en O(1) con §append§ y (c) **eliminar en O(1)** con el truco «intercambiar con el último y hacer pop». Sin §pos§ habría que buscar la cláusula en la lista (O(len)).",
  "**§mejor_insat§, §mejor_asig§:** la mejor asignación vista (la que dejó menos cláusulas insatisfechas). Es lo que se devuelve como §modelo§ en caso de UNKNOWN y de donde sale §pct§ y el error relativo.",
);

H2("2.8 §walksat§: el algoritmo paso a paso");
CODE(`
para cada reinicio (aquí 1):
    asig = asignación inicial al azar
    verdaderos[c] = cantidad de literales verdaderos en cada cláusula c
    insat = [c : verdaderos[c] == 0]
    repetir hasta max_flips veces:
        si insat está vacía: salir (¡solución!)
        si len(insat) < mejor_insat: guardar esta asignación como la mejor
        c = una cláusula de insat elegida AL AZAR
        para cada literal lit de c:
            romper[lit] = cuántas cláusulas con verdaderos == 1 contienen el literal
                          que HOY es verdadero para esa variable   # "break count"
        si min(romper) == 0:                       # movimiento gratis
            elegir la primera variable con romper == 0
        sino con probabilidad p:                    # ruido
            elegir UNA variable cualquiera de c
        sino:
            elegir la variable con menor romper
        dar vuelta la variable elegida y actualizar verdaderos / insat
si insat quedó vacía: devolver SAT (modelo = asig)
sino: devolver UNKNOWN (modelo = mejor_asig)`);
H3("Los conceptos, uno por uno");
BS(
  "**Asignación inicial al azar:** cada variable V o F con prob. 1/2. Es el punto de partida de la caminata.",
  "**Cláusula insatisfecha:** ninguno de sus literales es verdadero. El objetivo es que no quede ninguna. Elegir una insatisfecha **y arreglarla** asegura que cada paso tenga una razón: hay algo que corregir.",
  "**Por qué una variable DE ESA cláusula:** dar vuelta cualquiera de sus tres variables hace verdadero uno de sus literales, así que esa cláusula pasa a estar satisfecha. El problema es qué otras cláusulas se rompen.",
  "**Break count (cantidad que rompe):** si doy vuelta la variable x, el literal que hoy es verdadero para x pasa a falso. Las cláusulas que contienen ese literal pierden un literal verdadero; las que tenían **exactamente uno** (verdaderos == 1) quedan insatisfechas. La cantidad de esas cláusulas es el «break count» de x. En el código: §sum(1 for cc in aparece.get(lit_verdadero, []) if verdaderos[cc] == 1)§.",
  "**Free move (movimiento libre):** si alguna variable de la cláusula tiene break count 0, darla vuelta no rompe nada y arregla la cláusula: es estrictamente bueno y se elige siempre, sin tirar ninguna moneda.",
  "**Ruido p:** si todas las opciones rompen algo, con probabilidad p (0,5) se elige una variable cualquiera de la cláusula (aunque rompa más) y con 1−p la que menos rompe. El ruido permite **salir de mínimos locales**: sin él, el algoritmo sería puramente goloso y podría quedar atrapado.",
  "**Reinicios:** volver a empezar con otra asignación al azar. El parámetro existe pero se usó 1 (un solo intento de max_flips pasos).",
  "**Límite de pasos:** si se agotan los §max_flips§ sin llegar a cero insatisfechas, devuelve UNKNOWN. Es una decisión de **tiempo**, no una prueba de nada.",
  "**Dar vuelta y actualizar:** al cambiar x, el literal que era verdadero (§era_verdadero§) pasa a falso: para cada cláusula que lo contiene, §verdaderos -= 1§ y, si llega a 0, entra en §insat§. El literal que era falso (§ahora_verdadero§) pasa a verdadero: para cada cláusula que lo contiene, §verdaderos += 1§ y, si llega a 1 (estaba insatisfecha), sale de §insat§. El costo es proporcional a cuántas cláusulas contienen a x: unas pocas decenas.",
);
CAJA(
  "MEMORIA",
  "Resumen en una frase: «WalkSAT agarra una cláusula que está mal, y da vuelta la variable de esa cláusula que menos daño hace (o, a veces, una al azar para no quedar trabado), y repite hasta que nada esté mal o se le acaben los pasos.»",
);
H3("Ejemplo a mano 1: un free move");
P(
  "Con la fórmula del enunciado F = [[1,-2,3], [-1,2,4], [2,-3,-4]] partimos de x1=F, x2=V, x3=F, x4=F:",
);
BS(
  "Cláusula 1 (x1 ∨ ¬x2 ∨ x3): x1=F, ¬x2=F, x3=F → **insatisfecha**. Cláusula 2: ¬x1=V, x2=V → 2 verdaderos. Cláusula 3: x2=V, ¬x3=V, ¬x4=V → 3 verdaderos. §insat = [cláusula 1]§.",
  "Se elige la cláusula 1 (la única). Break counts: dar vuelta x1 rompe 0 (el literal que hoy es verdadero para x1 es ¬x1, que está en la cláusula 2, que tiene 2 verdaderos); x2 rompe 0; x3 rompe 0.",
  "Hay un free move: se elige la primera variable con break 0, o sea x1. Nuevo estado: x1=V, x2=V, x3=F, x4=F. La cláusula 1 se cumple por x1; la 2 por x2; la 3 por x2. §insat§ queda vacía → **SAT en 1 paso**.",
);
H3("Ejemplo a mano 2: sin free move y con ruido");
P(
  "Fórmula H = [[4,1,3], [3,1,-4], [2,4,1], [-4,-1,2], [2,-3,-4]] y estado x1=F, x2=F, x3=F, x4=V:",
);
CODE(`
cláusula  literales    verdaderos
   0      (x4 ∨ x1 ∨ x3)      1   (x4)
   1      (x3 ∨ x1 ∨ ¬x4)     0   <- INSATISFECHA
   2      (x2 ∨ x4 ∨ x1)      1   (x4)
   3      (¬x4 ∨ ¬x1 ∨ x2)    1   (¬x1)
   4      (x2 ∨ ¬x3 ∨ ¬x4)    1   (¬x3)

Se elige la cláusula 1 (única insatisfecha). Break counts de sus variables:
   x3: hoy el literal verdadero de x3 es ¬x3  -> está en la cláusula 4 (verdaderos=1) -> rompe 1
   x1: hoy el literal verdadero de x1 es ¬x1  -> está en la cláusula 3 (verdaderos=1) -> rompe 1
   x4: hoy el literal verdadero de x4 es  x4  -> está en las cláusulas 0 y 2 (ambas =1) -> rompe 2
min(romper) = 1 > 0: no hay free move.
   - con prob. p = 0,5: se elige una variable cualquiera de la cláusula (x3, x1 o x4 al azar)
   - con prob. 0,5:     se elige la que menos rompe (x3, la primera con romper = 1)
Si se elige x3: x3 pasa a V -> la cláusula 1 se cumple, pero la 4 queda con 0 verdaderos -> insat = [4].
Y el proceso continúa con la cláusula 4.`);
P(
  "Este ejemplo muestra por qué hace falta el ruido: arreglar la cláusula 1 rompió la 4; sin movimientos aleatorios ocasionales, el algoritmo podría ciclar entre pocas asignaciones.",
);
H3("Complejidad y comportamiento");
BS(
  "**Un paso** cuesta O(k·c), con c ≈ cuántas cláusulas contienen un literal dado. En 3-SAT con m/n ≈ 4,26 cada literal aparece en ≈ 3·4,26/2 ≈ 6,4 cláusulas en promedio, así que un paso son unas pocas decenas de operaciones: independiente del tamaño de la fórmula. Medido: ≈ 3,8 µs por paso en n ≤ 32 (100.000 pasos en 0,38 s) y ≈ 5,2 µs en n=10.000 (5.000.000 pasos en ≈ 26 s) [DERIVADO].",
  "**Espacio:** O(m·k) (la fórmula, §aparece§ y §verdaderos§).",
  "**Cuántos pasos necesita** es lo que no se puede acotar: depende de la fórmula y del azar. En las planted de n ≤ 32 hicieron falta entre 6,4 y 381,8 pasos en promedio; en n=5.000, millones.",
);
CAJA("OJO", [
  "**Detalle fino del código (no cambia ningún resultado del TP):** la mejor asignación se anota al **principio** de cada iteración, antes de dar vuelta. Si el último paso del tope mejora la situación, ese estado final no se compara con el mejor. Efecto: en UNKNOWN el «mejor» podría estar un paso atrasado. En las insatisfacibles de la comparación el mejor es siempre 1 insatisfecha, que es el óptimo, así que no afecta nada.",
  "**Otro detalle:** el docstring de §walksat§ dice «con probabilidad p da vuelta una variable cualquiera; si no, la que rompe menos (si alguna rompe 0, esa)». El código en realidad mira **primero** si hay un free move y solo después tira la moneda de p (es la versión estándar de Selman-Kautz-Cohen, y así lo describe el informe sec. 5). El docstring es una simplificación imprecisa.",
]);

H2("2.9 §experimentos.py§");
H3("Funciones auxiliares");
BS(
  "§guardar(nombre, filas)§: escribe una lista de diccionarios como CSV (columnas = claves, en orden de aparición).",
  "§correr_walksat_varias_veces(formula, n, veces=10, max_flips=100000)§: corre WalkSAT 10 veces con semillas 0 a 9 y devuelve (éxitos, tiempo medio, pasos medios, **mejor** porcentaje de cláusulas satisfechas entre las 10 corridas).",
  "§_cpu()§, §_ram_gb()§, §escribir_hardware()§: obtienen CPU (con wmic o PowerShell en Windows), RAM y demás, y escriben §hardware.txt§. Se llama **al iniciar** el comando §exacto§ (se pisa cada vez).",
  "§LIMITE§ (variable de entorno, por defecto 300 s) y §RAZON = 4.26§.",
);
H3("§exacto(tamanios)§");
CODE(`
para cada n pedido:
    m = round(4.26 * n)
    planted = generar_planted(n, m, semilla = n)
    insat   = generar_insatisfacible(n, m, semilla = n)
    para cada fórmula (planted, insat):
        guardar la fórmula en instancias/ (DIMACS)
        e = fuerza_bruta(fórmula, n, LIMITE)
        walksat 10 veces (semillas 0..9, 100.000 pasos, p=0,5)
        fila = tipo, n, m, estado_exacto, tiempo, asignaciones, cláusulas evaluadas,
               éxitos de 10, tiempo medio, pasos medios, mejor %, error_relativo, optimo_alcanzado
        guardar en resultados_exacto.csv (acumulando: reemplaza la fila si ya existía)`);
BS(
  "**error_relativo** = (100 − mejor_pct)/100 = fracción de cláusulas insatisfechas de la mejor asignación entre las 10 corridas.",
  "**optimo_alcanzado:** «n.a.» si el exacto dio TIMEOUT o si la fórmula es insatisfacible; si no, «sí» cuando el mejor porcentaje fue 100, «no» en otro caso. En el CSV el «sí» aparece mal codificado (ver Anexo, inconsistencia 4).",
  "**Se puede correr por tandas:** si el CSV existe, lo lee y reemplaza solo las filas de los n pedidos.",
);
H3("§escala()§");
BS(
  "Para n en 100, 500, 1.000, 5.000, 10.000 y para dos tipos: **planted** (m/n = 4,26) y **aleatoria** (m/n = 4,0), con semillas de fórmula 1, 2 y 3. Total 5 × 2 × 3 = **30 corridas**.",
  "WalkSAT con 5.000.000 de pasos máximos, p=0,5, **semilla = semilla_de_la_fórmula + 1000** y límite de 120 s.",
  "Al final corre una vez WalkSAT sobre la aleatoria n=500, m=2.000, semilla 1 (semilla de WalkSAT 1001) con historial, y escribe §historial_n500.csv§ (22.500 filas, una por paso).",
);
H3("§densidad()§");
BS(
  "n=16 fijo. Para m/n en 2,0; 3,0; 4,0; 4,26; 5,0; 6,0 genera 30 fórmulas aleatorias (semillas 0 a 29), las resuelve con el exacto (sin límite) y con WalkSAT (20.000 pasos, semilla+1000).",
  "Cuenta: cuántas son satisfacibles según el exacto, cuántas de ellas encontró WalkSAT y cuántas veces dijo SAT cuando el exacto dijo UNSAT (§walk_mal§, «no debería pasar nunca»).",
  "Tarda ≈ 20 s. **Volví a correrlo con el código actual y da exactamente el mismo CSV**, así que el CSV de densidad es reproducible aunque no se haya regenerado.",
);

H2("2.10 §graficos.py§");
TABLA(
  ["Salida", "Qué grafica", "De qué CSV"],
  [
    [
      "§tabla_comparacion.md§",
      "Tabla n × tipo con estado, tiempo del exacto, éxitos, tiempo de WalkSAT y error relativo.",
      "exacto",
    ],
    [
      "Gráfico 1",
      "Tiempo de la fuerza bruta vs. n (escala log). Insatisfacible (círculos) y satisfacible (cuadrados). La x roja marca el TIMEOUT.",
      "exacto",
    ],
    [
      "Gráfico 2",
      "Tiempo medio de WalkSAT vs. n (escala log-log), promedio de las 3 instancias por punto.",
      "escala",
    ],
    [
      "Gráfico 3",
      "Exacto vs. WalkSAT sobre las mismas fórmulas (tiempos, escala log).",
      "exacto",
    ],
    [
      "Gráfico 4",
      "% de corridas de WalkSAT que encontraron solución según n (escala).",
      "escala",
    ],
    [
      "Gráfico 5",
      "Asignaciones evaluadas por la fuerza bruta (escala log).",
      "exacto",
    ],
    [
      "Gráfico 6",
      "Cláusulas insatisfechas en cada paso de WalkSAT (n=500).",
      "historial",
    ],
  ],
  [1900, 5200, 1926],
);
BS(
  "Las filas con TIMEOUT **se excluyen** de las curvas de tiempo (no son un tiempo medido) y se dibujan aparte como una x roja. Antes se graficaban como si fueran un tiempo: es el error que se corrigió.",
  "Los gráficos 1, 3, 5 usan eje y logarítmico: ahí una recta significa crecimiento exponencial (sección 5.2).",
  "Los gráficos 1, 2, 3 están en la presentación (diapositivas 5, 9 y 7); los 4, 5 y 6 solo en el informe.",
);

H2("2.11 Las semillas: qué se fija y por qué");
TABLA(
  ["Qué", "Semilla", "Por qué así"],
  [
    [
      "Fórmulas del exacto (planted e insat.)",
      "= n",
      "Una fórmula distinta por cada n, pero siempre la misma para un n dado: reproducible.",
    ],
    [
      "WalkSAT en el exacto",
      "0 a 9 (10 corridas)",
      "10 corridas independientes sobre la misma fórmula, siempre las mismas.",
    ],
    ["Fórmulas de la escala", "1, 2, 3", "Tres fórmulas por tipo y por n."],
    [
      "WalkSAT en la escala y densidad",
      "semilla de la fórmula + 1000",
      "**Para que WalkSAT no arranque en la solución oculta** (ver más abajo).",
    ],
    ["Fórmulas de densidad", "0 a 29", "30 fórmulas por densidad."],
    [
      "Cláusulas aleatorias dentro de la insatisfacible",
      "semilla + 7919",
      "No reutilizar la semilla que eligió las 3 variables del núcleo.",
    ],
  ],
  [3200, 2100, 3726],
);
CAJA("OJO", [
  "**La trampa de la semilla.** Las primeras n llamadas a §random()§ de un generador con semilla s producen, en §generar_planted§, la solución oculta; y en §walksat§, la asignación inicial. Si ambos usaran la misma semilla, **WalkSAT arrancaría exactamente en la solución oculta** y resolvería todo en 0 pasos. Lo verifiqué: §Random(10)§ genera la misma secuencia en los dos casos.",
  "En la **escala** se evita sumando 1000. En el **exacto** se evita de casualidad (o por diseño): las fórmulas planted usan semilla n ≥ 10 y WalkSAT usa semillas 0 a 9, que no se superponen. Si alguien cambiara el rango de n a empezar en 5, aparecería la trampa. Conviene saber esto por si preguntan «¿por qué +1000?».",
]);

H2("2.12 ¿Cómo sabemos que el código es correcto?");
P(
  "No hay una batería de tests. Hay cuatro evidencias, de más débil a más fuerte. Si les preguntan «¿cómo saben que fuerza_bruta es correcta?», usen estas:",
);
BS(
  "**Coherencia entre algoritmos independientes:** el exacto y WalkSAT son dos programas distintos. En las 180 fórmulas de densidad coinciden en todas las satisfacibles (111 de 111) y WalkSAT nunca dijo SAT cuando el exacto dijo UNSAT.",
  "**Instancias con respuesta conocida por construcción:** las 12 planted dan SAT y las 11 insatisfacibles que terminaron dan UNSAT, justo lo que se construyó. Un exacto que falle en una de esas tendría que equivocarse justo al revés.",
  "**Conteos exactos:** en cada UNSAT, §asignaciones§ es exactamente 2ⁿ (de 1.024 a 1.073.741.824). Prueba que recorrió todo el espacio.",
  `**Verificación independiente de los modelos (script §verificar_modelos.py§, reproducido):** re-corre WalkSAT sobre las 12 instancias planted de §instancias/§ con las 10 semillas y comprueba con §contar_insatisfechas§ el modelo de las 120 corridas exitosas: dio 0 insatisfechas en todas; y sobre el modelo de la fuerza bruta en las planted de n ≤ 22. Por lo tanto, cuando dicen SAT, la asignación realmente cumple la fórmula.`,
);
CAJA(
  "OJO",
  "Qué hace realmente WalkSAT: declara SAT cuando la lista §insat§ queda vacía, es decir, cuando no quedan cláusulas insatisfechas según sus contadores internos §verdaderos§, que se actualizan de forma incremental; **no vuelve a recalcular** con §contar_insatisfechas§. El informe ya está corregido para decir esto. Como control aparte, §verificar_modelos.py§ confirmó los 120 modelos. Ver Anexo.",
);
// =====================================================================
// PARTE 3
// =====================================================================
H1("Parte 3 · Registro de decisiones");
P(
  "Un trabajo no es solo lo que se hizo: es lo que se decidió **no** hacer y por qué. Cada fila de la tabla responde a la pregunta «¿por qué hicieron esto así?». Si el profesor la formula, la respuesta tiene que salir sin dudar.",
);
CAJA("HONESTIDAD", [
  "Las decisiones las tomó la IA que armó el TP, con la aprobación de ustedes. Las razones de la columna «Por qué» son las que están documentadas en el informe, el README o en el propio código, o las que se deducen directamente de ellos; donde agrego una razón que es mi razonamiento y no está escrita en ningún documento del TP, lo marco como [razonamiento propio].",
  "Para cada decisión pregúntense: «¿yo podría defenderla con mis palabras?». Si no, vuelvan a la sección de la Parte 2 que la explica.",
]);
H2("3.1 Tabla de decisiones");
TABLA(
  [
    "Decisión",
    "Alternativa descartada",
    "Por qué",
    "Consecuencia",
    "Quién lo puede preguntar",
  ],
  [
    [
      "**Elegir SAT / 3-SAT**",
      "Otros NP-completos (coloreo, viajante, mochila).",
      "Es el primer NP-completo (Cook-Levin), tiene formato estándar (DIMACS), solvers reales para comparar, y generar y verificar fórmulas es simple. [Si el tema lo asignó la cátedra, decirlo así.]",
      "Todo el vocabulario (CNF, cláusulas) y el umbral 4,26 aplican. Los casos reales vienen por reducción (coloreo, horarios).",
      "Camilo (diap. 2)",
    ],
    [
      "**Fuerza bruta en Python puro**",
      "C/C++, numpy, bitmasks vectorizados, o un DPLL/CDCL.",
      "Es la referencia más simple de entender y de verificar; la consigna pide un algoritmo exacto y una heurística (no necesariamente el mejor exacto). [razonamiento propio]",
      "Lento (~600.000 asignaciones/s): por eso el límite de 1 h se alcanza en **n=32**. Con C se llegaría a más variables (n≈40 es una estimación mía, no medida), con la misma curva exponencial.",
      "Camilo (diap. 3, 5)",
    ],
    [
      "**WalkSAT como heurística**",
      "GSAT (anterior, más lento), Novelty y variantes, recocido simulado, o un solver CDCL.",
      "Es la búsqueda local clásica para SAT (Selman-Kautz-Cohen 1994): simple, pocas líneas, muy eficaz en 3-SAT aleatorio. Es incompleta por diseño, lo que ilustra el tema del TP.",
      "Resuelve planted en ms y hasta 5.000 variables en segundos. No puede probar UNSAT. No se comparó con ninguna variante.",
      "Camilo (diap. 6)",
    ],
    [
      "**Dos tipos de fórmula para el exacto**",
      "Fórmulas aleatorias puras para todo.",
      "Hay que **saber** la respuesta: planted = SAT segura; insatisfacible construida = UNSAT segura. Y para medir el peor caso del exacto hay que forzarlo a recorrer 2ⁿ (UNSAT).",
      "Los resultados dependen de dos familias especiales (fáciles / con núcleo obvio). La comparación es ilustrativa, no representativa.",
      "Camilo (diap. 4, 5, 7)",
    ],
    [
      "**m = round(4,26·n)**",
      "m/n fijo en otro valor (3, 5…).",
      "Es la densidad donde, en 3-SAT aleatorio, está el umbral de fase [LITERATURA]: la zona más difícil.",
      "Las fórmulas del exacto son las «difíciles de manual», pero planted y con núcleo, no aleatorias puras: no hay que decir que se probó el umbral.",
      "Camilo (diap. 4); Facundo (si preguntan por 4,0)",
    ],
    [
      "**Límite de 3.600 s y resultado TIMEOUT**",
      "Dejarlo correr sin límite; o el defecto del código (300 s).",
      "La consigna pide mostrar un caso que supere la hora. Con límite 3.600 se mide hasta n=30 y n=32 queda cortado. Un TIMEOUT no se grafica como tiempo.",
      "n=32 insatisfacible: **no hay tiempo medido**; solo se sabe que superó 1 h y llevaba 42 %. La estimación (≈ 2 h) es proyección.",
      "Camilo (diap. 5)",
    ],
    [
      "**Una sola corrida por punto en el exacto**",
      "Repetir y promediar varias veces.",
      "Cada punto de n=30 tarda media hora; repetir 12 tamaños × 2 tipos × varias veces es inviable en esta máquina. [razonamiento propio]",
      "No hay dispersión ni intervalos. La razón ≈ 4 sale de 6 pares consecutivos (n=18 a 30); la irregularidad en planted es esperable.",
      "Camilo (diap. 5, 14 también)",
    ],
    [
      "**10 semillas de WalkSAT**",
      "1 sola corrida o 100.",
      "WalkSAT es aleatorio; 10 corridas con semillas 0..9 dan una idea de éxitos/10 sin costar casi nada (ms).",
      "«10/10» es un conteo; con 10 corridas no se estima una probabilidad fina.",
      "Camilo (diap. 7, 8)",
    ],
    [
      "**100.000 pasos en la comparación; 5.000.000 en la escala**",
      "El mismo tope en los dos.",
      "n ≤ 32 se resuelve en decenas de pasos: 100.000 sobra. En n grande hacen falta millones.",
      "El fallo en n=10.000 (0/6) es un efecto del tope; no sabemos qué pasaría con más pasos.",
      "Camilo (diap. 6); Facundo (diap. 9)",
    ],
    [
      "**Error relativo = fracción de cláusulas insatisfechas**",
      "(valor − óptimo)/óptimo, como en optimización.",
      "SAT es de decisión; para tener una medida de «cuán cerca» se mira Max-SAT. Con óptimo 0 (satisfacible) la fórmula clásica daría una división por cero.",
      "Es una medida de cercanía, no de corrección. En UNSAT da 1/m porque el óptimo Max-SAT es 1 (derivado de la construcción).",
      "Camilo (diap. 8)",
    ],
    [
      "**«n.a.» en optimo_alcanzado**",
      "Poner «sí» en las insatisfacibles.",
      "El programa no calcula el óptimo Max-SAT de una insatisfacible (sólo prueba que no hay asignación perfecta) ni tiene respuesta en TIMEOUT.",
      "El informe agrega, por razonamiento, que 1/m es el óptimo, pero la columna queda n.a.",
      "Camilo (diap. 8)",
    ],
    [
      "**Experimento de densidad con n=16**",
      "n más grande.",
      "Hay que correr el **exacto** sobre 180 fórmulas; con n=16 tarda ≈ 20 s en total.",
      "Muestra la tendencia, no el umbral nítido. El CSV es reproducible con el código actual (lo verifiqué), aunque no se regeneró.",
      "Camilo (diap. 8)",
    ],
    [
      "**No comparar exacto y heurístico con n ≥ 1.000**",
      "Comparar en todos los tamaños.",
      "El exacto no puede llegar (2¹⁰⁰⁰ asignaciones); no hay referencia contra qué comparar la respuesta.",
      "En la escala solo se mide a WalkSAT (éxito, tiempo, % de cláusulas).",
      "Facundo (diap. 9)",
    ],
    [
      "**Generar los documentos con scripts (generadores/)**",
      "Escribir los números a mano en Word/PowerPoint.",
      "Evita copiar mal: toda cifra sale de los CSV. Si se vuelve a correr, se regeneran informe, diapositivas, guion y guía.",
      "Hay que correr §node generadores/*.js§ después de cualquier cambio de datos. Si alguien edita el .docx a mano, el próximo regenerado lo pisa.",
      "Ambos",
    ],
    [
      "**Caso UTDT (ILP, no SAT)**",
      "Un caso real que use SAT.",
      "No se encontró un caso argentino contemporáneo documentado con SAT; el de la UTDT es el mejor documentado de horarios y es pariente (problema NP-completo; restricciones codificables en CNF).",
      "Hay que decir siempre «no es SAT puro». No se infiere nada de SAT a partir de él.",
      "Facundo (diap. 10)",
    ],
    [
      "**Puente coloreo → SAT (UNR)**",
      "Afirmar que la UNR usa SAT.",
      "Escalante, Chiesa y Fekete usan ILP para coloreo por listas. El puente a SAT lo da la reducción estándar de k-coloreo a CNF, que es teoría, no un resultado de ellos.",
      "La traducción horarios → coloreo → SAT es nuestro razonamiento, no algo que ellos hayan hecho.",
      "Facundo (diap. 10)",
    ],
    [
      "**Modelo FRCU hipotético**",
      "Presentarlo como la solución de la FRCU.",
      "No hay evidencia pública de que la FRCU use SAT ni datos reales. El ejemplo de 9 variables y 19 cláusulas sirve para ilustrar la codificación.",
      "Siempre rotulado HIPOTÉTICO. No se puede afirmar nada de su desempeño real.",
      "Facundo (diap. 12, 13)",
    ],
    [
      "**Sin cifra de ahorro económico**",
      "Estimar un ahorro (por ejemplo, en horas docentes).",
      "No existe una cifra pública atribuible a SAT. Inventar una sería deshonesto. La única cifra (50 M€ por año, ferrocarril neerlandés) es una analogía de investigación operativa, no SAT.",
      "Hay indicadores propuestos a medir (conflictos evitados, huecos, ocupación de aulas), no resultados.",
      "Facundo (diap. 12, 13)",
    ],
    [
      "**p = 0,5 y un solo reinicio, sin ajustar**",
      "Barrer p y reinicios.",
      "Es un valor típico y alcanzó para estos tamaños. [razonamiento propio]",
      "No se sabe si con otro p o con reinicios se resolvería n=10.000. Está declarado como limitación y trabajo futuro.",
      "Camilo (diap. 6); Facundo (diap. 14)",
    ],
    [
      "**Aleatorias con m/n = 4,0 en la escala (y planted con 4,26)**",
      "Aleatorias con 4,26.",
      "Con 4,26 la mitad serían insatisfacibles y WalkSAT no podría resolverlas; con 4,0 hay buena chance de ser satisfacibles. [razonamiento propio]",
      "No se sabe con certeza que las aleatorias de la escala sean satisfacibles: un UNKNOWN no prueba nada.",
      "Facundo (diap. 9)",
    ],
    [
      "**n par de 10 a 32**",
      "Todos los n.",
      "Cada +2 multiplica por 4 el trabajo; con pasos de 2 se ve bien la tendencia y se reduce el total de corridas.",
      "Permite calcular la razón t(n+2)/t(n).",
      "Camilo (diap. 4, 5)",
    ],
    [
      "**Tiempo con time.perf_counter(), sin la generación de la fórmula**",
      "Medir el programa entero (con generación y lectura).",
      "Se mide solo el algoritmo, que es lo que se compara.",
      "Los tiempos de WalkSAT incluyen armar sus tablas internas (§aparece§, §verdaderos§).",
      "Camilo (diap. 4)",
    ],
  ],
  [1500, 1500, 2600, 2300, 1126],
  { size: 16 },
);

H2("3.2 Qué es medido, qué es proyección y qué es hipotético");
P(
  "La honestidad es parte de la nota. Si reparten cada afirmación en uno de estos tres cajones, nadie los puede acusar de inventar:",
);
TABLA(
  ["Cajón", "Qué entra", "Ejemplos en el TP"],
  [
    [
      "**MEDIDO**",
      "Cualquier cosa que salió de una corrida real y está en un CSV.",
      `Tiempos del exacto (n=30: ${fn(n30.t_exacto_s, 1)} s); 10/10 de WalkSAT; TIMEOUT de n=32 con ${fn(n32u.asignaciones)} asignaciones evaluadas (${pctCorte} %); densidad 111/111; escala (4/6 en n=5.000; 0/6 en n=10.000).`,
    ],
    [
      "**DERIVADO**",
      "Cuentas sobre números medidos.",
      `Razón t(n+2)/t(n) ≈ ${rg} (de ${rmin} a ${rmax}, n=18 a 30); asignaciones por segundo; ≈ 8 cláusulas por asignación; error relativo = 1/m en UNSAT.`,
    ],
    [
      "**PROYECCIÓN**",
      "Extrapolaciones no medidas.",
      `Tiempo que habría tardado n=32 insatisfacible: ≈ ${proy32} h (con la razón 4) o ≈ ${proy32v} h (con la velocidad observada). Es solo una referencia coherente con el TIMEOUT.`,
    ],
    [
      "**LITERATURA**",
      "Resultados conocidos que el TP cita pero no mide.",
      "Cook-Levin; P vs NP abierto; umbral ≈ 4,26; Kissat ganó la SAT Competition 2024 [VERIFICAR]; casos de Eclipse/Sat4j, libsolv, BMC.",
    ],
    [
      "**HIPOTÉTICO**",
      "Modelos propios sin datos reales.",
      "El modelo de horarios para la FRCU; el ejemplo de 9 variables; los indicadores a medir; cualquier idea de ahorro.",
    ],
  ],
  [1700, 2800, 4526],
);

H2("3.3 Referencias: cuáles son sólidas y cuáles hay que verificar");
P(
  "Esto es lo que dice el propio informe: «las referencias se limitan a fuentes cuyos metadatos se comprobaron con un buscador, y se marcó lo que no se verificó en el texto original». Por eso, **antes de exponer, abran las fuentes** de la columna «Verificar» (la lista completa de pasos está en la Parte 7).",
);
TABLA(
  ["Grado", "Referencias", "Qué respaldan"],
  [
    [
      "**Clásicas, muy firmes**",
      "Cook 1971 (STOC); Levin 1973; Karp 1972; Davis-Putnam 1960; Davis-Logemann-Loveland 1962; Selman-Kautz-Cohen 1994 (AAAI-94); Marques-Silva y Sakallah 1999 (GRASP); Eén y Sörensson 2003 (MiniSat); Handbook of Satisfiability (Biere et al. 2021).",
      "NP-completitud, DPLL, CDCL, WalkSAT, MiniSat. Son referencias de manual.",
    ],
    [
      "**Firmes pero con datos a chequear**",
      "Biere-Cimatti-Clarke-Zhu 1999 (BMC); Kautz y Selman 1992; Le Berre y Parrain 2010 (Sat4j); Audemard y Simon 2009 (Glucose); de Moura y Bjørner 2008 (Z3).",
      "Usos y herramientas. Chequear año, congreso y que digan lo que se les atribuye.",
    ],
    [
      "**Verificadas solo por resumen o metadatos**",
      "Ferreiro y Marenco (53 JAIIO/SIIIO, UTDT); Escalante, Chiesa y Fekete (SUMA 2019, FCEIA-UNR); Asín Achá y Nieuwenhuis 2014; Kroon, Huisman et al. 2009 (cifra de 50 M€); SAT Competition 2024 (Kissat); libsolv; ETAPS Test of Time 2017.",
      "El caso argentino, el puente de coloreo, el uso de SAT en horarios, la analogía económica, la afirmación sobre Kissat. **Hay que abrirlas y leerlas.**",
    ],
  ],
  [1900, 4300, 2826],
);
CAJA(
  "OJO",
  "La cifra de 50 millones de euros (Kroon et al.) está marcada en el propio informe como «no verificada en el texto: tomada del resumen de un buscador». Si van a decirla en voz alta, abran el artículo (o su abstract) y confirmen el número, el año y a qué se refiere exactamente (beneficio anual de qué). Si no pueden confirmarlo, no la digan.",
);
// =====================================================================
// PARTE 4
// =====================================================================
function DIAP(o) {
  H2(`Diapositiva ${o.n} · ${o.titulo}  (${o.tiempo})`);
  P("**Qué dice la diapositiva.** " + o.dice);
  P("**Qué hay que entender detrás.**", { after: 40, keepNext: true });
  BS(...o.detras);
  P("**Qué decir (puntos clave, no leer):**", {
    after: 40,
    keepNext: true,
    before: 80,
  });
  BS(...o.decir);
  CAJA("MEMORIA", "Números para saber de memoria: " + o.memoria);
  CAJA("OJO", "Error típico: " + o.error);
  o.preg.forEach(([q, a]) => QA(q, a));
}
const t = {
  exacto30: `${fn(n30.t_exacto_s, 0)} s`,
  min30: fn(n30.t_exacto_s / 60, 1),
};

H1("Parte 4 · Guía de la exposición");
P(
  "Una sección por persona. En cada diapositiva encuentran: qué dice, qué hay que entender detrás, qué decir (puntos clave, **no** para leer), el número que tienen que saber de memoria, el error típico a evitar y 2 o 3 preguntas probables con respuesta modelo. Los tiempos son los del guion (§SAT_Guion_Exposicion.docx§).",
);
CAJA(
  "MEMORIA",
  "Reparto: Camilo = diapositivas 1 a 8 (0:00 a 10:00). Facundo = diapositivas 9 a 16 (10:00 a 20:00). Numeración confirmada leyendo el .pptx real (16 diapositivas).",
);
TABLA(
  ["Diap.", "Título", "Quién", "Tiempo"],
  [
    ["1", "Portada", CAMILO, "0:00 a 0:40"],
    ["2", "El problema SAT y por qué importa", CAMILO, "0:40 a 2:00"],
    ["3", "Algoritmo exacto: fuerza bruta", CAMILO, "2:00 a 3:10"],
    ["4", "Equipo y metodología", CAMILO, "3:10 a 4:00"],
    ["5", "Resultado: tiempo de la fuerza bruta", CAMILO, "4:00 a 5:50"],
    ["6", "Heurística: WalkSAT (búsqueda local)", CAMILO, "5:50 a 6:50"],
    ["7", "Exacto vs. WalkSAT: mismas fórmulas", CAMILO, "6:50 a 8:20"],
    ["8", "Éxitos, error relativo y tiempos", CAMILO, "8:20 a 10:00"],
    ["9", "Escalabilidad de WalkSAT", FACU, "10:00 a 11:10"],
    ["10", "Aplicación real: horarios en Argentina", FACU, "11:10 a 12:30"],
    ["11", "Usos industriales y herramientas", FACU, "12:30 a 13:40"],
    ["12", "Modelo FRCU e impacto económico", FACU, "13:40 a 14:50"],
    ["13", "Viabilidad de la aplicación", FACU, "14:50 a 16:00"],
    ["14", "Reflexión crítica", FACU, "16:00 a 17:10"],
    ["15", "Conclusiones", FACU, "17:10 a 18:40"],
    ["16", "Referencias y cierre", FACU, "18:40 a 20:00"],
  ],
  [800, 4500, 1700, 2026],
);

// ---------------------------------------------------------------- CAMILO
H2("4.1 Camilo · diapositivas 1 a 8 (0:00 a 10:00)");
P(
  "**Tu misión:** que el tribunal entienda el problema, vea cómo funciona la fuerza bruta, se convenza con los números de que es exponencial y comprenda qué es WalkSAT y qué gana y qué pierde. Tu mitad es la **teórica y experimental**: la que más preguntas de fundamentos recibe. Tenés que dominar la Parte 1, la Parte 2 y la Parte 5.",
);

DIAP({
  n: 1,
  titulo: "Portada",
  tiempo: "0:00 a 0:40",
  dice: "Título «SAT: Problema de Satisfacción Booleana», subtítulo (solución exacta de fuerza bruta frente a la heurística WalkSAT, una búsqueda local; instancias 3-SAT), integrantes, materia (Teoría de la Computabilidad, Prof. Pascal) y fecha. La fecha todavía dice «[completar fecha de exposición]».",
  detras: [
    "Es el momento de plantear la pregunta del trabajo en una frase: ¿se puede saber si una fórmula lógica con muchas variables tiene solución, y a qué costo?",
    "Tenés que poder resumir todo el TP en 20 segundos: «Desarrollamos, con asistencia de una herramienta de IA, un algoritmo exacto de fuerza bruta y una heurística, WalkSAT; los comparamos sobre las mismas fórmulas; medimos hasta dónde llega cada uno; y analizamos una aplicación posible, los horarios.»",
  ],
  decir: [
    "Saludo, presentación de ambos.",
    "Dejar claro que el tema es el problema SAT, no WalkSAT: «si tenemos una fórmula con muchísimas variables, ¿existe alguna combinación de verdadero y falso que la haga verdadera?». Es SAT, el primer problema que se demostró NP-completo (Cook-Levin, 1971). En clase SAT se nombró como ejemplo de NP-completo pero no se desarrolló: por eso lo explicamos desde cero.",
    "Qué vamos a ver sobre SAT, en cuatro cosas: un algoritmo exacto de fuerza bruta que siempre acierta; una heurística de búsqueda local, WalkSAT, rápida pero sin garantías; mediciones reales en nuestro equipo; y dónde aparece en la vida real, en particular en horarios.",
    "Anunciar el reparto: «yo hasta la diapositiva 8, Facundo desde la 9».",
  ],
  memoria:
    "20 minutos; 16 diapositivas; vos 1 a 8 (0:00 a 10:00), Facundo 9 a 16 (10:00 a 20:00).",
  error:
    "leer la portada, o arrancar con «vamos a hablar de SAT» sin plantear una pregunta. Arrancá con la pregunta.",
  preg: [
    [
      "¿Por qué eligieron SAT?",
      "Porque es el problema NP-completo de referencia (el primero que se demostró, con el teorema de Cook-Levin), tiene un formato estándar (DIMACS), existen solvers reales con los que compararse y es fácil de generar y de verificar. Además aparece en aplicaciones concretas. [Si el tema lo asignó la cátedra, decir: «el tema fue asignado, y nos pareció adecuado porque…».]",
    ],
    [
      "¿Qué hizo cada uno de ustedes en el trabajo?",
      [
        "Respondan con la verdad. Mi recomendación: decir que se repartieron la exposición y, sobre el trabajo, explicar con precisión qué entendieron y cómo se prepararon (revisaron el código, corrieron los experimentos, verificaron las referencias). **Si utilizaron herramientas de IA para generar código y texto, y la cátedra lo permite, díganlo; si no está claro si se permite, consulten antes de la exposición. No digan «lo programamos nosotros línea por línea» si no es cierto.** Ver también la pregunta 30 del banco (Parte 6).",
      ],
    ],
  ],
});

DIAP({
  n: 2,
  titulo: "El problema SAT y por qué importa",
  tiempo: "0:40 a 2:00",
  dice: "La pregunta de SAT; la fórmula F = (x1 ∨ ¬x2 ∨ x3) ∧ (¬x1 ∨ x2 ∨ x4) ∧ (x2 ∨ ¬x3 ∨ ¬x4); tres ideas: NP-completo (Cook 1971, Levin; 3-SAT también), verificar es fácil (está en NP, tiempo polinomial) y buscar es difícil (2ⁿ asignaciones).",
  detras: [
    "Qué es una fórmula, un literal, una cláusula, CNF, 3-SAT (sección 1.6 y 1.7). En la fórmula de la diapositiva hay **4 variables y 3 cláusulas**; el espacio de búsqueda es 2⁴ = 16 asignaciones.",
    "Por qué verificar es polinomial: se recorren las m cláusulas y se comprueba que cada una tenga un literal verdadero: O(m·k). Por qué buscar parece exponencial: no se conoce nada mejor que explorar un espacio de 2ⁿ (con refinamientos).",
    "NP-completo = está en NP y es NP-difícil (sección 1.2). Cook-Levin: SAT es NP-completo (sección 1.4). 3-SAT también, por reducción (sección 1.7).",
    "Max-SAT: variante de optimización (máximo de cláusulas cumplidas). Se usa después como vara del error relativo.",
  ],
  decir: [
    "Mostrar la fórmula: «cuatro variables, tres cláusulas. Cada cláusula es un OR de literales; las cláusulas van unidas por AND: hay que cumplir todas a la vez».",
    "«La consigna habla de SAT. Nosotros restringimos las instancias a 3-SAT, con exactamente tres literales por cláusula: sigue siendo NP-completo y deja generar fórmulas de forma uniforme.»",
    "Las tres ideas en orden: NP-completo → verificar es fácil (polinomial) → buscar es difícil (2ⁿ).",
    "Mencionar Max-SAT en una frase, para dejarla plantada.",
  ],
  memoria:
    "4 variables, 3 cláusulas, 2⁴ = 16 asignaciones; Cook 1971 y Levin (independiente); verificar = O(m·k).",
  error:
    "decir «NP quiere decir no polinomial» o «Karp demostró que SAT es NP-completo» (fue Cook; Karp usó SAT para demostrar otros 21 problemas). Tampoco digas que «SAT no se puede resolver rápido»: nadie lo demostró (P vs NP está abierto).",
  preg: [
    [
      "¿Por qué SAT está en NP?",
      "Porque una solución propuesta, es decir una asignación, se verifica en tiempo polinomial: recorro las m cláusulas y compruebo que cada una tenga un literal verdadero. Eso lleva O(m·k), polinomial en el tamaño de la fórmula.",
    ],
    [
      "¿Por qué 3-SAT es NP-completo?",
      "Está en NP porque es un caso particular de SAT. Y es NP-difícil porque SAT se reduce a 3-SAT en tiempo polinomial: cada cláusula larga se parte con variables nuevas, por ejemplo (a∨b∨c∨d) se reemplaza por (a∨b∨y)∧(¬y∨c∨d); y las cortas se rellenan. La fórmula original es satisfacible si y solo si la nueva lo es.",
    ],
    [
      "¿Cuál es la diferencia entre SAT, 3-SAT y Max-SAT?",
      "SAT admite cláusulas de cualquier largo; 3-SAT, exactamente tres literales; ambos son problemas de decisión. Max-SAT es de optimización: busca la asignación que cumple la mayor cantidad posible de cláusulas. Dato extra: 2-SAT se resuelve en tiempo polinomial.",
    ],
  ],
});

DIAP({
  n: 3,
  titulo: "Algoritmo exacto: fuerza bruta",
  tiempo: "2:00 a 3:10",
  dice: "Pseudocódigo: para cada asignación de 0 a 2ⁿ−1, para cada cláusula, si ningún literal es verdadero se descarta la asignación; si cumple todas, devuelve SAT; si termina, UNSAT. Complejidad O(2ⁿ·m·k). «Exacta: si no hay solución, lo demuestra. Cada variable más duplica el trabajo.»",
  detras: [
    "Cada entero de 0 a 2ⁿ−1, en binario, es una asignación: el bit i−1 dice si x_i es verdadera (sección 2.1). El contador recorre todas las asignaciones sin repetir.",
    "Los cortes tempranos: se corta una cláusula en cuanto un literal es verdadero y una asignación en cuanto una cláusula es falsa. Por eso la cota teórica m·k por asignación es el peor caso; en la práctica son ≈ 8 cláusulas (sección 2.6).",
    "Tres resultados posibles en el código: SAT, UNSAT y TIMEOUT (si pasa del límite de tiempo; se chequea cada 16.384 asignaciones).",
    "**Exacto = completo y correcto**: es el único de los dos que puede demostrar UNSAT.",
  ],
  decir: [
    "Explicar el mapeo entero ↔ asignación con un ejemplo de 3 bits (por ejemplo 5 = 101 = V, F, V).",
    "Recorrer el pseudocódigo en voz alta, con el cursor sobre cada línea.",
    "Complejidad: «O de 2ⁿ por m por k: 2ⁿ asignaciones, y por cada una hasta m cláusulas de k literales».",
    "Las dos ideas: es exacto (único que demuestra UNSAT) y es exponencial (cada variable duplica el trabajo).",
  ],
  memoria: `O(2ⁿ · m · k); n=30 → ${fn(n30.asignaciones)} asignaciones; tres resultados: SAT, UNSAT, TIMEOUT.`,
  error:
    "decir que la fuerza bruta es «el algoritmo exacto» (hay exactos mucho mejores: DPLL, CDCL; la fuerza bruta es una referencia deliberadamente ingenua) o omitir m·k en la complejidad.",
  preg: [
    [
      "¿Existen algoritmos exactos mejores que la fuerza bruta?",
      "Sí. DPLL (backtracking con propagación de unidades) y, sobre todo, los solvers CDCL, que aprenden cláusulas a partir de los conflictos. Siguen siendo exponenciales en el peor caso, pero resuelven en la práctica instancias con millones de variables si tienen estructura. Para 3-SAT existen además algoritmos de peor caso del orden de 1,3ⁿ [LITERATURA, VERIFICAR]. No usamos ninguno: la fuerza bruta es nuestra referencia, y lo declaramos como limitación.",
    ],
    [
      "¿Cómo representa el código una asignación y qué costo tiene evaluarla?",
      "Como un entero cuyos bits son los valores de las variables. Evaluar una asignación recorre las cláusulas y, para cada literal, extrae su bit con un desplazamiento y una máscara. En el peor caso O(m·k); en la práctica corta en la primera cláusula falsa.",
    ],
    [
      "¿Qué quiere decir que el algoritmo es «correcto» y «completo»?",
      "Correcto: lo que responde es verdad (si dice SAT hay solución; si dice UNSAT no hay). Completo: responde en todos los casos (dado tiempo). La fuerza bruta es las dos cosas; WalkSAT es correcto pero incompleto.",
    ],
  ],
});

DIAP({
  n: 4,
  titulo: "Equipo y metodología",
  tiempo: "3:10 a 4:00",
  dice: `Equipo: AMD Ryzen 5 5600X, 6 núcleos y 12 lógicos, RAM 31,9 GB, Windows 11, Python 3.12.1, límite 3.600 s, corrido el 04/11/2026. Metodología: mismas fórmulas (DIMACS, semillas fijas) para los dos algoritmos; n=10 a 32 de a 2, una planted y una insatisfacible; WalkSAT 10 corridas, p=0,5, 100.000 pasos; escala: n=100 a 10.000, 3 planted + 3 aleatorias, hasta 5.000.000 de pasos.`,
  detras: [
    "Por qué un equipo único y qué información da §hardware.txt§ (CPU, RAM, SO, Python, LIMITE, fecha). Es lo que permite interpretar los tiempos: un tiempo sin equipo no significa nada.",
    "«6 núcleos, 12 lógicos» porque el procesador tiene hyper-threading; **el programa usa un solo hilo**. No hay paralelismo.",
    "«Mismas fórmulas»: se generan una vez, se guardan en DIMACS y las dos algoritmos las resuelven. Eso hace justa la comparación.",
    "planted vs. insatisfacible (sección 1.13) y por qué las dos: la primera mide cuán rápido se encuentra una solución, la segunda fuerza al exacto a recorrer todo.",
    "Semillas fijas = reproducibilidad: cualquiera que corra el código obtiene exactamente las mismas fórmulas (los tiempos cambiarán con el equipo).",
  ],
  decir: [
    "Equipo en una oración. Remarcar «un solo hilo».",
    "«Los dos algoritmos resolvieron exactamente las mismas fórmulas, guardadas en DIMACS y con semillas fijas.»",
    "Para cada n par de 10 a 32 dos fórmulas: planted (con solución garantizada) e insatisfacible. WalkSAT, 10 corridas por fórmula, p = 0,5, 100.000 pasos.",
    "Para escalabilidad: de 100 a 10.000 variables, tres planted y tres aleatorias, hasta 5.000.000 de pasos.",
  ],
  memoria:
    "Ryzen 5 5600X; 6/12 núcleos; 31,9 GB; Windows 11; Python 3.12.1; límite 3.600 s; 04/11/2026; n = 10 a 32; 10 corridas; p = 0,5; 100.000 pasos; escala 100 a 10.000; 5.000.000 de pasos.",
  error:
    "decir «usamos los 12 hilos», o afirmar que las condiciones estaban «controladas» como un hecho: el README las pide (enchufado, alto rendimiento, sin otros programas), pero no hay registro de que se cumplieron. Decí «el procedimiento pide» o «se corrió en un mismo equipo».",
  preg: [
    [
      "¿Por qué las fórmulas planted? ¿No son demasiado fáciles?",
      "Son fáciles, y lo reconocemos en la reflexión crítica. Se usan porque garantizan que hay solución, así sabemos que la respuesta correcta es SAT y podemos medir cuán rápido cada algoritmo la encuentra. Para no depender solo de casos fáciles agregamos las insatisfacibles, las aleatorias de la escala (m/n=4,0) y el experimento de densidad.",
    ],
    [
      "¿Los resultados se repiten en otra máquina?",
      "Los números absolutos no: dependen de la CPU, de Python y de la carga. La forma de la curva sí: en el exacto, el tiempo se multiplica por ≈ 4 cada dos variables. Las fórmulas y los conteos (éxitos, asignaciones evaluadas) sí se repiten exactos porque están fijados por las semillas.",
    ],
    [
      "¿Por qué un límite de 3.600 s?",
      "Para poder mostrar un caso que supera una hora y a la vez que la corrida termine. Con ese límite, n=30 insatisfacible termina (1.768 s) y n=32 se corta.",
    ],
  ],
});

DIAP({
  n: 5,
  titulo: "Resultado: tiempo de la fuerza bruta",
  tiempo: "4:00 a 5:50",
  dice: `Gráfico de tiempo vs. n (escala log) con una x roja (TIMEOUT). Textos: «Insatisfacible: recorre las 2ⁿ. La x roja es TIMEOUT: no es un tiempo medido»; n=30: 1.768 s (29,5 min); n=32: no terminó en 1 h; t(n+2)/t(n) ≈ 4,0 (de 3,8 a 4,3, n=18 a 30) ≈ 4 = 2²; n=28: 465 s; n=26: 113 s; con solución n=32: 631 s.`,
  detras: [
    "**Por qué se mide con insatisfacibles:** obligan a recorrer las 2ⁿ asignaciones. Con una satisfacible se frena al encontrar la primera solución y el tiempo depende de dónde caiga (sección 5.4).",
    "**La razón ≈ 4:** si t(n) ≈ k·2ⁿ, t(n+2)/t(n) = 2² = 4. Con los datos: media geométrica 4,01 entre n=18 y n=30 (sección 5.3).",
    "**Una recta en escala log = crecimiento exponencial** (sección 5.2).",
    "**La x roja:** n=32 insatisfacible llegó a 3.600 s sin terminar. No hay un tiempo medido. Llevaba 1.805.975.552 de 4.294.967.296 asignaciones (42 %).",
    "**Proyección vs. medición:** extrapolando, n=32 habría tardado ≈ 2,0 h. Es una estimación, nunca un resultado.",
  ],
  decir: [
    "«Medimos con insatisfacibles construidas a propósito porque obligan a recorrer todo.»",
    "Dar los tres puntos: n=26 → 113 s, n=28 → 465 s, n=30 → 1.768 s, que son 29,5 minutos y exactamente 2³⁰ = 1.073.741.824 asignaciones.",
    "El patrón: «cada dos variables más, el tiempo se multiplica por cuatro, que es 2 al cuadrado». En el gráfico log es una recta.",
    "La x roja: «con 32 variables, no terminó en una hora; lo cortamos a los 3.600 s. No tenemos un tiempo medido». Si hablan de la proyección, decirla como **proyección**.",
    "Aclaración: en las satisfacibles el tiempo es irregular: n=32 631 s; n=30 solo 9,83 s (menos que n=28: 42,5 s).",
  ],
  memoria: `n=26: 113 s · n=28: 465 s · n=30: 1.768 s (29,5 min; ${fn(n30.asignaciones)} asignaciones) · n=32: TIMEOUT (> 1 h; 42 %) · razón ≈ ${rg} (${rmin} a ${rmax}) · planted n=32: 631 s.`,
  error: `presentar «2 horas» como lo que tardó n=32. No lo midieron. Tampoco confundir la x roja con un punto de la curva. Y no decir «tardó una hora» (se cortó a la hora; no se sabe cuánto habría tardado).`,
  preg: [
    [
      "¿Llegaron a una hora de ejecución?",
      `Sí, en el sentido de que el exacto con 32 variables insatisfacible superó la hora: lo cortamos a los 3.600 segundos con el ${pctCorte} % de las asignaciones evaluadas. Lo último que terminó fue n=30, en 1.768 s (29,5 minutos). No tenemos un tiempo medido para n=32. Con la razón de cuatro, la proyección da ≈ ${proy32} h; con la velocidad observada, ≈ ${proy32v} h. Lo decimos como proyección.`,
    ],
    [
      "¿Por qué usaron fórmulas insatisfacibles para medir el exacto?",
      "Porque en una insatisfacible no hay atajo: el exacto tiene que recorrer las 2ⁿ asignaciones antes de poder responder UNSAT. Con una satisfacible se detiene al encontrar la primera solución, y el tiempo depende de la suerte: lo vimos con n=30, que tardó menos que n=28.",
    ],
    [
      "¿Cómo saben que el crecimiento es exponencial y no otra cosa?",
      `Por tres evidencias: (1) la cantidad de asignaciones evaluadas es exactamente 2ⁿ en cada insatisfacible (Gráfico 5); (2) la razón t(n+2)/t(n) es ≈ 4 para n ≥ 18 (entre ${rmin} y ${rmax}); (3) en escala log la curva es una recta. Con un crecimiento polinomial n^c, la razón t(n+2)/t(n) iría tendiendo a 1, no se mantendría en 4.`,
    ],
  ],
});

DIAP({
  n: 6,
  titulo: "Heurística: WalkSAT (búsqueda local)",
  tiempo: "5:50 a 6:50",
  dice: "Cómo funciona: 1) asignación inicial al azar; 2) elegir una cláusula insatisfecha; 3) dar vuelta una variable de esa cláusula: la que rompe menos, o una al azar con probabilidad p; 4) repetir hasta cumplir todo o agotar los pasos. Es incompleta: si no encuentra devuelve UNKNOWN, y eso NO significa insatisfacible. Parámetros: p = 0,5 · 100.000 pasos (n ≤ 32) o 5.000.000 (n ≥ 100) · semillas fijas.",
  detras: [
    "Todo el algoritmo está explicado en las secciones 2.7 y 2.8. Tenés que poder hacer el **ejemplo a mano 1** sin mirar y explicar el **ejemplo 2** (cuando no hay free move).",
    "«La que rompe menos»: el **break count** (cuántas cláusulas hoy satisfechas pasarían a insatisfechas). Si hay una con break 0 se elige siempre (free move); si no, con prob. p una al azar y con prob. 1−p la de menor break.",
    "Por qué el ruido: para salir de mínimos locales.",
    "**Incompleto:** puede no encontrar una solución que existe y no puede demostrar UNSAT. Resultado posible: SAT (con asignación) o UNKNOWN (con la mejor asignación vista).",
  ],
  decir: [
    "«WalkSAT es una búsqueda local con azar»: ni recorre todo el espacio ni es puramente al azar.",
    "Los cuatro pasos de la diapositiva con el ejemplo de la fórmula chica si hay tiempo.",
    "El ruido en una oración: «con probabilidad p da vuelta una cualquiera, para no quedarse atascada».",
    "**Lo más importante:** UNKNOWN no es UNSAT. «Solo el exacto puede demostrar que no existe».",
  ],
  memoria:
    "p = 0,5; 100.000 pasos (n ≤ 32); 5.000.000 pasos (n ≥ 100); 1 reinicio; UNKNOWN ≠ UNSAT; paso ≈ 4 µs en Python.",
  error:
    "decir que WalkSAT «siempre encuentra la solución», que «elige la variable al azar» (elige la que menos rompe, salvo ruido), o que «p es la probabilidad de éxito» (es la probabilidad de movimiento aleatorio).",
  preg: [
    [
      "¿Por qué incluye azar? ¿No sería mejor siempre elegir la mejor variable?",
      "Si siempre se eligiera la que menos rompe, el algoritmo es puramente goloso y puede quedar atrapado en un mínimo local, o ciclar entre pocas asignaciones. El ruido permite salir. Además hace que corridas con semillas distintas exploren caminos distintos.",
    ],
    [
      "¿Cuál es la diferencia con GSAT?",
      "GSAT elige entre **todas** las variables de la fórmula la que más mejora la cantidad de cláusulas satisfechas: cada paso es caro. WalkSAT primero elige una cláusula insatisfecha y sólo considera sus k variables: es mucho más barato y enfocado. Selman, Kautz y Cohen (1994) mostraron que con ruido funciona mejor [LITERATURA].",
    ],
    [
      "¿Por qué p = 0,5? ¿Probaron otros valores?",
      "No lo ajustamos: es un valor razonable, típico, y alcanzó para estos tamaños. Está declarado como limitación y trabajo futuro (barrer p y reinicios). No podemos decir que sea el óptimo.",
    ],
  ],
});

DIAP({
  n: 7,
  titulo: "Exacto vs. WalkSAT: mismas fórmulas",
  tiempo: "6:50 a 8:20",
  dice: "Gráfico de tiempos de exacto y WalkSAT. «Con solución: WalkSAT 10/10 en todos los n, como máximo 1,4 ms. n=32: exacto 631 s; WalkSAT 1,4 ms. Insatisfacible: exacto demuestra UNSAT hasta n=30; con n=32 no terminó en 1 h.»",
  detras: [
    `Hay **tres curvas** en el gráfico: fuerza bruta en satisfacibles, fuerza bruta en insatisfacibles y WalkSAT en satisfacibles. **No hay curva de WalkSAT en las insatisfacibles** porque su tiempo ahí no mide nada útil: gasta los 100.000 pasos (≈ 392 ms) y devuelve UNKNOWN.`,
    `Cociente en n=32: ${fn(n32p.t_exacto_s, 1)} s contra ${ft(n32p.t_walksat_medio_s)}: ≈ 4,4 × 10⁵ veces [DERIVADO]. Pero es una sola fórmula planted, fácil. No se generaliza.`,
    "Planted: WalkSAT 10/10 en los 12 tamaños (120/120 corridas), tiempo medio máximo 1,45 ms.",
    "La comparación es **justa** en el sentido de que ambos resuelven las mismas fórmulas (DIMACS), pero no es una comparación de dos algoritmos óptimos: el exacto es ingenuo y está en Python.",
  ],
  decir: [
    "Con solución, WalkSAT la encontró las 10 veces en todos los tamaños y nunca tardó más de 1,4 ms. En n=32: 1,4 ms contra 631 s.",
    "Con las insatisfacibles, solo el exacto contesta (UNSAT hasta n=30, 29,5 minutos). En n=32 no hay punto: hay una marca de TIMEOUT.",
    "Cerrar con el enganche: los números exactos están en la tabla.",
  ],
  memoria: `WalkSAT planted: 10/10 en todos los n, ≤ 1,4 ms · n=32: exacto ${fn(n32p.t_exacto_s, 0)} s, WalkSAT ${ft(n32p.t_walksat_medio_s)} · UNSAT: exacto hasta n=30.`,
  error:
    "decir «WalkSAT es 400.000 veces más rápido» a secas. Es en una fórmula con solución, fácil, de n=32; no es una ley. Y no compares el tiempo de WalkSAT en UNSAT con el del exacto: WalkSAT ahí no resuelve nada.",
  preg: [
    [
      "¿Es justa la comparación?",
      "Es justa en que ambos resuelven exactamente las mismas fórmulas, guardadas en DIMACS. No es una comparación entre los mejores algoritmos: el exacto es fuerza bruta en Python y WalkSAT también está en Python. Y las fórmulas planted son fáciles. Por eso decimos que la comparación es ilustrativa.",
    ],
    [
      "¿Por qué WalkSAT es tan rápido en las satisfacibles?",
      "Porque no explora el espacio de 2ⁿ: camina por él guiado por las cláusulas insatisfechas. En las planted hay muchas soluciones «cerca» y cada paso mejora la situación; necesitó entre 6 y 380 pasos en promedio. Cada paso cuesta microsegundos.",
    ],
    [
      "¿Qué pasa con el tiempo de WalkSAT en las insatisfacibles?",
      "Gasta siempre los 100.000 pasos (≈ 392 ms) y devuelve UNKNOWN. Ese tiempo no mide la dificultad de la fórmula sino el tope que le pusimos. Por eso no se compara con el exacto.",
    ],
  ],
});

DIAP({
  n: 8,
  titulo: "Éxitos, error relativo y tiempos",
  tiempo: "8:20 a 10:00",
  dice: "Tabla de n = 20, 26, 30 y 32: exacto SAT, exacto UNSAT, WalkSAT SAT (éxitos/10 y tiempo), WalkSAT UNSAT (0/10 y error relativo). Texto: «En UNSAT agotó los 100.000 pasos (≈ 392 ms) y dejó 0,7 % a 2,3 % de cláusulas sin cumplir. n=16, razón m/n de 2 a 6, 30 fórmulas c/u: halló la solución en 111 de 111 satisfacibles y dijo SAT falsamente 0 veces.»",
  detras: [
    "**Error relativo** = fracción de cláusulas que la mejor asignación de WalkSAT dejó sin cumplir (sección 1.14). Baja al subir n porque la fórmula tiene más cláusulas y siempre queda **una** insatisfecha: 1/43 = 2,3 % en n=10; 1/136 = 0,7 % en n=32.",
    "**Por qué siempre 1 y no 0:** la fórmula es UNSAT por construcción, así que ninguna asignación cumple todo; el mínimo posible es 1 insatisfecha (sección 2.4). WalkSAT llegó al óptimo de Max-SAT.",
    "**«n.a.»** en optimo_alcanzado: no aplica, porque el programa no calcula el óptimo Max-SAT de una insatisfacible ni tiene respuesta en un TIMEOUT.",
    "**Densidad (n=16):** al subir m/n de 2 a 6 las satisfacibles caen de 30/30 a 1/30. WalkSAT encontró solución en las 111 satisfacibles y nunca dijo SAT sobre una insatisfacible (0 de 69 insatisfacibles).",
  ],
  decir: [
    "Leer la tabla en voz alta solo en los puntos clave: con solución WalkSAT acertó 10/10; en insatisfacibles 0/10 (informó UNKNOWN).",
    "Definir el error relativo y por qué baja con n.",
    "El experimento de densidad en una frase y el resultado: 111 de 111, 0 falsos.",
    "Aclaración honesta: una sola fórmula por tamaño, planted fáciles; la comparación es ilustrativa.",
  ],
  memoria: `n=20: exacto 322 ms / 1,68 s; WalkSAT 10/10 0,11 ms / 0/10 y 1,2 % · n=26: 3,22 s / 113 s; 0,19 ms; 0,9 % · n=30: 9,83 s / 1.768 s; 1,1 ms; 0,8 % · n=32: 631 s / no terminó; 1,4 ms; 0,7 % · error 0,7 % a 2,3 % · 392 ms · densidad: 111/111, 0 falsos.`,
  error:
    "interpretar el error relativo como «probabilidad de que WalkSAT se equivoque» o decir que «WalkSAT falló en las insatisfacibles» (no hay nada que encontrar: no es un fallo). Y no digas que el experimento de densidad «confirma el umbral 4,26».",
  preg: [
    [
      "¿Qué es el error relativo y por qué lo definieron así?",
      "Es la fracción de cláusulas que quedan sin cumplir en la mejor asignación que encontró WalkSAT, dividida por m. En un problema de optimización se usaría (valor − óptimo)/óptimo, pero SAT es de decisión y el óptimo de las satisfacibles es 0, así que esa fórmula no sirve. Miramos Max-SAT como vara. Vale 0 si cumplió todo. Es una medida de cercanía, no de correctitud.",
    ],
    [
      "¿Qué significa «n.a.» y por qué WalkSAT «falla» en las insatisfacibles?",
      "«n.a.» es «no aplica»: no conocemos el óptimo de Max-SAT con los datos del experimento. Y WalkSAT no falla: en una fórmula insatisfacible no existe ninguna solución que encontrar. Hace lo que debe: devuelve UNKNOWN con la mejor asignación que vio, que deja una sola cláusula sin cumplir.",
    ],
    [
      "¿Puede WalkSAT dar un falso SAT?",
      "No. Declara SAT únicamente cuando la lista de cláusulas insatisfechas queda vacía, es decir, cuando la asignación cumple todas. Lo comprobamos: 0 falsos SAT en las 180 fórmulas de densidad y, para esta guía, verificamos con un verificador independiente los 120 modelos de WalkSAT de la comparación. Matiz: lo que impide el falso SAT es que los contadores internos sean correctos; el código no revalida con un chequeo aparte.",
    ],
  ],
});

H3("Transición Camilo → Facundo (10:00)");
CAJA("EJEMPLO", [
  "**Camilo cierra la diapositiva 8:** «…en síntesis, con 32 variables el exacto ya no llegó, y WalkSAT sigue en milisegundos. La pregunta natural es: ¿hasta dónde llega WalkSAT? Eso lo sigue Facundo.»",
  "**Facundo arranca la diapositiva 9:** «Gracias, Camilo. Probamos WalkSAT con 100, 500, 1.000, 5.000 y 10.000 variables…»",
  "Ensayen este pase para que sea limpio y no se pisen. Camilo no vuelve a hablar hasta las preguntas, salvo que Facundo le ceda una respuesta de la primera mitad.",
]);

H2("4.2 Lo que Facundo tiene que saber de la mitad de Camilo");
P(
  "Si el tribunal le hace una pregunta de las diapositivas 1 a 8 a Facundo, tiene que poder contestar a nivel **básico**. Mínimo imprescindible:",
);
BS(
  "SAT: pregunta de decisión sobre una fórmula CNF; 3-SAT = tres literales por cláusula; NP-completo (Cook-Levin); verificar es polinomial; buscar tiene 2ⁿ asignaciones.",
  "Fuerza bruta: prueba las 2ⁿ asignaciones; O(2ⁿ·m·k); exacta; es la única que demuestra UNSAT.",
  `Equipo: Ryzen 5 5600X, 31,9 GB, Python 3.12.1, un hilo, límite 3.600 s.`,
  `Resultados del exacto: n=30 → ${t.exacto30} (${t.min30} min); n=32 no terminó en 1 h; razón ≈ 4 (= 2²) cada +2 variables.`,
  "WalkSAT: búsqueda local; elige una cláusula insatisfecha, da vuelta la variable que menos rompe, con ruido p=0,5; incompleta: UNKNOWN ≠ UNSAT.",
  "Comparación: WalkSAT 10/10 en todas las planted, ≤ 1,4 ms; en UNSAT 0/10 con error 0,7 a 2,3 %; el exacto n=32: 631 s en planted.",
);

// ---------------------------------------------------------------- FACUNDO
H2("4.3 Facundo · diapositivas 9 a 16 (10:00 a 20:00)");
P(
  "**Tu misión:** mostrar hasta dónde llega WalkSAT, y después llevar el problema a la realidad: casos concretos, herramientas, un modelo propio, viabilidad y una evaluación honesta. Tu mitad es la de la **honestidad intelectual**: casi todas las preguntas que te harán son del tipo «¿esto lo midieron o lo suponen?». Tenés que dominar la Parte 1 (en especial reducciones y Max-SAT), la sección 3.2 (medido / proyección / hipotético) y la lista de «qué NO decir» (Parte 7).",
);

DIAP({
  n: 9,
  titulo: "Escalabilidad de WalkSAT",
  tiempo: "10:00 a 11:10",
  dice: `Tabla: n | éxitos | t medio: 100 → 6/6, 2,5 ms; 500 → 6/6, 205 ms; 1.000 → 6/6, 337 ms; 5.000 → 4/6, 16,0 s; 10.000 → 0/6, 26,1 s. «n=10.000: 0 de 6 llegó a la solución en 5.000.000 pasos (≈ 27,4 s); quedó hasta 0,06 % sin cumplir. Es el tope del experimento, no prueba de UNSAT. n=5.000: planted 3/3, aleatoria 1/3.»`,
  detras: [
    "Seis corridas por n: 3 planted (m/n = 4,26) + 3 aleatorias (m/n = 4,0). Éxito = encontró solución dentro de 5.000.000 de pasos y 120 s.",
    "**n=10.000:** las 6 corridas terminaron en UNKNOWN por el **tope de pasos** (5.000.000 ≈ 24,5 a 27,4 s por corrida); el límite de 120 s no se alcanzó. La mejor asignación dejó entre 1 y 24 cláusulas sin cumplir (≥ 99,94 % satisfechas). Eso significa «cerca», no «imposible».",
    "**Por qué la aleatoria con m/n=4,0 es más difícil que la planted con 4,26:** la planted está sesgada hacia una solución oculta; la aleatoria no tiene esa «pista» (sección 2.3). La dificultad no depende solo de n ni de m/n.",
    "**No se sabe con certeza que las aleatorias de la escala sean satisfacibles.** Con m/n = 4,0, por debajo del umbral, probablemente lo son, pero no está probado en el TP: por eso un UNKNOWN allí es aún más ambiguo.",
    "**Tamaño de la muestra:** 3 por tipo. Un «1 de 3» no permite estimar una probabilidad.",
    "El gráfico 2 muestra el tiempo medio vs. n en escala log-log; en n=10.000 se aplana porque el tiempo está acotado por el tope de pasos, no porque WalkSAT «se estabilice».",
  ],
  decir: [
    "«Probamos WalkSAT con 100, 500, 1.000, 5.000 y 10.000 variables. Hasta 1.000 resolvió las seis, con un tiempo medio de 337 ms en 1.000.»",
    "«Con 5.000 resolvió 4 de 6: las tres planted y una de las tres aleatorias; tiempo medio 16 s.»",
    "«Con 10.000 no llegó a la solución en ninguna de las seis corridas. Pero atención: se detuvo por el tope de 5 millones de pasos que fijamos, y dejó hasta 0,06 % de cláusulas sin cumplir. No dice que no haya solución: es el límite de nuestro experimento.»",
    "La idea de que la dificultad depende de la estructura: las planted fueron más fáciles.",
  ],
  memoria:
    "100: 6/6, 2,5 ms · 500: 6/6, 205 ms · 1.000: 6/6, 337 ms · 5.000: 4/6 (planted 3/3, aleatoria 1/3), 16,0 s · 10.000: 0/6, 26,1 s · tope 5.000.000 de pasos (≈ 27,4 s) · hasta 0,06 % sin cumplir.",
  error:
    "decir «WalkSAT no puede con 10.000 variables» o «las fórmulas de 10.000 son insatisfacibles». Ninguna de las dos cosas está probada. Lo correcto: «no llegó dentro del tope que fijamos».",
  preg: [
    [
      "¿Qué pasó con 10.000 variables?",
      "En ninguna de las 6 corridas encontró solución dentro de los 5.000.000 de pasos que fijamos (≈ 27 s por corrida). La mejor asignación dejó muy pocas cláusulas sin cumplir (como mucho 0,06 %). No significa que las fórmulas sean insatisfacibles ni que WalkSAT no pueda: con más pasos, otro p o reinicios quizá la encontraría. Es el límite del experimento, y así lo reportamos.",
    ],
    [
      "¿Por qué las aleatorias fueron más difíciles si tienen menos cláusulas por variable?",
      "Porque las planted están construidas alrededor de una solución oculta, lo que sesga las cláusulas hacia ella y facilita la búsqueda; las aleatorias no tienen esa pista. La dificultad depende de la estructura, no solo de n o de m/n. Con solo 3 fórmulas por punto lo tomamos como tendencia, no como ley.",
    ],
    [
      "¿Cómo saben que las aleatorias de la escala eran satisfacibles?",
      "No lo sabemos con certeza. m/n = 4,0 está por debajo de la zona del umbral (≈ 4,26 según la literatura), así que es probable que lo sean, pero no lo probamos. Por eso un UNKNOWN en esas fórmulas no puede interpretarse como UNSAT.",
    ],
  ],
});

DIAP({
  n: 10,
  titulo: "Aplicación real: horarios en Argentina",
  tiempo: "11:10 a 12:30",
  dice: "UTDT: Ferreiro y Marenco (53 JAIIO/SIIIO) usan programación lineal entera (ILP) para armar horarios, 35 módulos; impacto solo cualitativo; no es SAT puro, pero se codifica en CNF. UNR: Escalante, Chiesa y Fekete (SUMA 2019): colorear un grafo con k colores se reduce a SAT, con x(v,c) = el vértice v lleva el color c; adyacentes, colores distintos. Puente: horarios → coloreo de grafos → SAT.",
  detras: [
    "**ILP (programación lineal entera):** optimizar una función lineal con variables enteras y restricciones lineales. También es NP-completo en su versión de decisión; **no es SAT**, pero hay traducciones entre ellos.",
    "**UTDT:** semana de 5 días × 7 bandas = 35 módulos. Dos fases: reparto de estudiantes en comisiones y esquema de horario por comisión. Impacto reportado: cualitativo (soluciones de buena calidad en tiempos aceptables). Solver, tamaños y tiempos: [VERIFICAR] (el informe los marca como no verificados).",
    "**UNR:** asignación de aulas con dos sedes como coloreo de grafos por listas, resuelto con ILP sobre datos de 2018. **Ellos no usan SAT.** El puente a SAT es la reducción estándar de k-coloreo a CNF (teoría).",
    "**Horarios = coloreo:** comisiones = vértices; choques (mismo docente o alumnos) = aristas; franjas = colores. Dos comisiones en conflicto no pueden tener el mismo color.",
  ],
  decir: [
    "Presentar primero el caso argentino mejor documentado (UTDT) y **aclarar de entrada que usa ILP, no SAT**. Después, por qué lo traemos: es el mismo problema de fondo y las restricciones son lógicas.",
    "El puente: coloreo ↔ SAT con las variables x(v,c) y las tres familias de cláusulas.",
    "Cerrar la idea: «armar horarios es, en esencia, colorear: comisiones, choques y franjas».",
  ],
  memoria:
    "UTDT: Ferreiro y Marenco, 53 JAIIO (SIIIO), 35 módulos (5 × 7), ILP, impacto cualitativo · UNR: Escalante, Chiesa y Fekete, SUMA 2019, coloreo (también ILP).",
  error:
    "decir que la UTDT o la UNR usan SAT (usan ILP), inventar tamaños de instancia, tiempos o ahorros, o atribuirles el puente a SAT: el puente es nuestro razonamiento.",
  preg: [
    [
      "¿Por qué incluyeron un caso de ILP si el TP es de SAT?",
      "Porque es el caso argentino contemporáneo de horarios mejor documentado, y lo presentamos como tal: no es SAT puro. Lo traemos porque el problema de fondo es el mismo (NP-completo), las restricciones son lógicas y se codifican en CNF, y el coloreo de la UNR es el puente concreto. Lo usamos para mostrar que el problema existe y es importante, no para afirmar que ese caso use SAT.",
    ],
    [
      "¿Cómo se codifica un coloreo en CNF?",
      [
        "Variables x(v,c) por vértice v y color c. Tres familias de cláusulas: (1) al menos un color por vértice: (x(v,1) ∨ … ∨ x(v,k)); (2) a lo sumo un color por vértice: (¬x(v,c) ∨ ¬x(v,d)) para c<d; (3) los vecinos no comparten color: (¬x(u,c) ∨ ¬x(v,c)) para cada arista (u,v) y cada color c. Es satisfacible si y solo si hay un k-coloreo.",
        "Ejemplo: un triángulo con k=3 tiene 9 variables y 3 + 9 + 9 = 21 cláusulas, y es satisfacible; con k=2 es insatisfacible (un triángulo no es bipartito).",
      ],
    ],
    [
      "¿Qué diferencia hay entre SAT e ILP?",
      "En ILP las variables son enteras y las restricciones lineales (se optimiza una función); en SAT las variables son booleanas y las restricciones son cláusulas. Ambos son NP-completos en decisión y se pueden traducir uno al otro. SAT es muy bueno para restricciones lógicas; ILP, para objetivos numéricos como minimizar huecos.",
    ],
  ],
});

DIAP({
  n: 11,
  titulo: "Usos industriales y herramientas",
  tiempo: "12:30 a 13:40",
  dice: "Usos documentados: bounded model checking, dependencias de paquetes (Eclipse con Sat4j, libsolv), timetabling con SAT/MaxSAT (Asín Achá y Nieuwenhuis, 2014). Herramientas: MiniSat, Glucose, Kissat (ganó la SAT Competition 2024), Z3, OR-Tools CP-SAT.",
  detras: [
    "**Bounded model checking (BMC):** se verifica un sistema (hardware o software) buscando un error en los primeros k pasos de su ejecución; se codifica «hay un error en ≤ k pasos» como una fórmula SAT. Si el solver devuelve SAT, la asignación es un **contraejemplo** (el error); UNSAT significa «no hay error hasta k pasos». Acá se ve la importancia de un algoritmo **completo**: un UNKNOWN no sirve.",
    "**Dependencias de paquetes:** elegir qué versiones instalar, respetando «A requiere B ≥ 2», «C choca con D»: restricciones lógicas → CNF. Eclipse usó Sat4j (biblioteca de Java); libsolv (openSUSE) usa un solver SAT.",
    "**MiniSat, Glucose, Kissat:** solvers **CDCL** (completos). **Z3:** solver **SMT** (SAT + teorías: aritmética, vectores, etc.). **OR-Tools CP-SAT:** solver de programación por restricciones de Google, basado en técnicas SAT.",
    "**CDCL** = Conflict-Driven Clause Learning: backtracking que, ante un conflicto, analiza su causa, aprende una cláusula nueva y vuelve atrás varios niveles (sección 6, pregunta 11).",
  ],
  decir: [
    "Listar los tres usos con una frase cada uno, enfatizando por qué SAT sirve ahí (restricciones lógicas, necesidad de saber si no hay solución).",
    "Herramientas: «son libres y están pensadas para instancias mucho más grandes que nuestro prototipo en Python». No prometer rendimiento concreto.",
  ],
  memoria:
    "Usos: BMC; Sat4j/Eclipse y libsolv; horarios con SAT/MaxSAT (Asín Achá y Nieuwenhuis 2014) · Herramientas: MiniSat, Glucose, Kissat (SAT Competition 2024), Z3 (SMT), OR-Tools CP-SAT.",
  error:
    "afirmar que empresas específicas «usan SAT internamente» (Intel, Microsoft, Debian, Cargo, etc.): no se verificó. El informe dice: «No se afirma que empresas específicas usen SAT internamente». Tampoco digas qué tan rápido resolvería MiniSat/Kissat nuestras fórmulas: no lo medimos.",
  preg: [
    [
      "¿Por qué no usaron un solver CDCL como MiniSat para comparar?",
      "Porque el trabajo compara un exacto y una heurística implementados para el TP; un CDCL es otro nivel y habría opacado la comparación. Lo declaramos como limitación y como primer paso a futuro: correr MiniSat o Kissat sobre los mismos archivos DIMACS. No tenemos mediciones propias de un CDCL, así que no afirmamos nada sobre cuánto tardarían.",
    ],
    [
      "¿Qué es CDCL?",
      "Es el algoritmo de los solvers modernos. Parte de DPLL (elegir una variable, propagar las consecuencias, y volver atrás si hay conflicto) y agrega: análisis de conflictos que produce una cláusula nueva «aprendida» que evita repetir el error, retroceso no cronológico (volver varios niveles), reinicios y heurísticas para elegir variables (como VSIDS), y estructuras eficientes (watched literals). Es completo: responde SAT o UNSAT.",
    ],
    [
      "¿Qué diferencia hay entre SAT, SMT y CP-SAT?",
      "SAT: solo booleanas. SMT (Z3): SAT más teorías (enteros, vectores de bits, arreglos). CP-SAT: programación por restricciones con enteros, apoyada en técnicas de SAT. Todos se usan para modelar problemas de restricciones; cambia el lenguaje de modelado.",
    ],
  ],
});

DIAP({
  n: 12,
  titulo: "Modelo FRCU e impacto económico",
  tiempo: "13:40 a 14:50",
  dice: "Modelo HIPOTÉTICO (FRCU): x(c,f) = la comisión c va en la franja f. Al menos una franja, a lo sumo una, sin choques de docente ni de alumnos. Ejemplo: 9 variables y 19 cláusulas, resuelto por ambos algoritmos. «No hay evidencia pública de que la FRCU use SAT.» Impacto: sin cifra pública atribuible a SAT; analogía (no es SAT): horario ferroviario neerlandés, 50 M€ por año (Kroon, Huisman et al., Interfaces 2009).",
  detras: [
    "El ejemplo de juguete: 3 comisiones × 3 franjas = 9 variables. Cláusulas: 3 de «al menos una» + 9 de «a lo sumo una» (3 pares por comisión × 3 comisiones) + 6 de conflicto (2 pares en conflicto × 3 franjas) + 1 unitaria (el docente de C3 no puede en la franja 1) = **19**.",
    "Conflictos: C1 y C2 comparten docente; C2 y C3 comparten alumnos. Numeración x(c,f) = 3(c−1)+f.",
    "Resultado según el informe: el exacto devolvió SAT tras evaluar **139 asignaciones** (C1 en franja 2, C2 en 1, C3 en 2); WalkSAT (semilla 1) devolvió SAT en **4 pasos** (C1 en 1, C2 en 2, C3 en 3). **Yo lo reproduje** con mi propia codificación y obtuve exactamente esos números. Ojo: no hay un script en el repo para reproducirlo (ver Anexo).",
    "La fórmula **no es 3-SAT puro** (tiene cláusulas de 1, 2 y 3 literales). Se convierte a 3-SAT con el procedimiento de la sección 1.7.",
    "**Honestidad económica:** no hay cifra de ahorro. La analogía de 50 M€ es investigación operativa, no SAT, y está marcada como no verificada en el texto original [VERIFICAR].",
  ],
  decir: [
    "Insistir en la palabra **hipotético** al principio y al final.",
    "Explicar las variables y las tres o cuatro familias de cláusulas con un ejemplo.",
    "El ejemplo de 9 variables y 19 cláusulas: ambos algoritmos encontraron un horario válido.",
    "«No existe una cifra pública atribuible a SAT y no vamos a inventarla.» La analogía de los ferrocarriles, rotulada como analogía.",
  ],
  memoria:
    "9 variables, 19 cláusulas (3 + 9 + 6 + 1) · exacto: 139 asignaciones · WalkSAT: 4 pasos · analogía: 50 M€/año (Kroon, Huisman et al., 2009) [VERIFICAR].",
  error:
    "presentar el modelo como «lo que usa la FRCU», prometer un ahorro («reduciría X horas»), o atribuir los 50 M€ a SAT. Es una analogía de otro método.",
  preg: [
    [
      "¿Por qué no hay datos de ahorro económico?",
      "Porque no existe una cifra pública atribuible a SAT para horarios, y inventarla sería deshonesto. El impacto del caso de la UTDT es cualitativo. La única cifra del trabajo (50 millones de euros por año del horario ferroviario neerlandés) es una analogía de investigación operativa, rotulada así. Lo que sí proponemos son indicadores a medir antes y después de una herramienta: conflictos evitados, huecos de docentes, ocupación de aulas, casos sin solución.",
    ],
    [
      "¿La UTN FRCU usa SAT para los horarios?",
      "No encontramos evidencia pública de que lo haga. Documentamos que publican horarios con comisiones y franjas, y proponemos un modelo hipotético de cómo se podría codificar.",
    ],
    [
      "¿Cómo modelarían «a lo sumo una franja» para muchas franjas?",
      "Con la codificación por pares: una cláusula (¬x(c,f) ∨ ¬x(c,g)) por cada par de franjas. Para F franjas son F(F−1)/2 cláusulas por comisión, que crece cuadráticamente. Si F fuera grande se usan codificaciones más compactas con variables auxiliares (por ejemplo, el contador secuencial de Sinz [LITERATURA, VERIFICAR]).",
    ],
  ],
});

DIAP({
  n: 13,
  titulo: "Viabilidad de la aplicación",
  tiempo: "14:50 a 16:00",
  dice: "Cuatro planos. Técnica: la fuerza bruta no escala (n=32: más de 1 h); para producción, solvers como MiniSat o Kissat. Datos: hoy solo hay PDFs de horarios; faltan docentes, aulas y capacidades en formato usable. Económica: costo bajo (software libre), ahorro sin cifra, solo cualitativo. Límite del modelo: preferencias (huecos, carga) piden MaxSAT u ILP, no SAT puro.",
  detras: [
    "**Técnica:** un problema de horarios real tendría del orden de (comisiones × franjas) variables. [HIPOTÉTICO, ilustración mía, no un dato de la FRCU]: con 100 comisiones y 35 franjas serían 3.500 variables y decenas de miles de cláusulas, muy lejos de lo que la fuerza bruta soporta (n=32) pero chico para un solver CDCL.",
    "**Datos:** la barrera real. Un modelo necesita disponibilidades, aulas y capacidades en formato estructurado; hoy hay PDFs.",
    "**Económica:** software libre ⇒ costo directo bajo; el costo real es el tiempo de modelar y de obtener datos. El beneficio es cualitativo.",
    "**Límite del modelo:** SAT es de satisfacción (cumplir todo o nada). Los objetivos blandos («evitar huecos») son optimización: MaxSAT (cláusulas blandas con peso) o ILP.",
  ],
  decir: [
    "Cuatro planos, uno por oración, y la conclusión: viable como modelo, con un solver real y datos estructurados; no con nuestro prototipo.",
    "Recalcar el tema de los datos como el factor limitante más realista.",
  ],
  memoria:
    "Cuatro planos: técnico, datos, económico, límite del modelo · n=32 ya supera 1 h con fuerza bruta · solvers: MiniSat, Kissat.",
  error:
    "decir que «es viable» sin condiciones, o dar números de tamaño real de la FRCU: no los tienen.",
  preg: [
    [
      "¿Sería viable aplicar esto en la FRCU?",
      "Como modelo, sí: el problema de horarios es codificable en CNF y existen solvers que manejan instancias mucho mayores que las de una facultad. Pero faltan datos estructurados (hoy son PDFs), habría que cuantificar las preferencias (que no son SAT puro) y habría que usar un solver real, no nuestro prototipo. No lo probamos con datos reales.",
    ],
    [
      "¿Qué hacen con las preferencias blandas (huecos, equilibrio de carga)?",
      "SAT solo expresa restricciones duras (todo o nada). Para preferencias se usa MaxSAT (cláusulas blandas con peso, se maximiza la suma de las cumplidas) o ILP (función objetivo). Por eso en la diapositiva decimos que el modelo tiene este límite.",
    ],
    [
      "¿Y si el solver dice UNSAT para un horario?",
      "Sería información valiosa: significa que las restricciones duras son contradictorias y hay que relajar alguna. Solo un algoritmo completo lo puede afirmar; WalkSAT diría UNKNOWN. Es el argumento a favor de un CDCL en aplicaciones reales.",
    ],
  ],
});

DIAP({
  n: 14,
  titulo: "Reflexión crítica",
  tiempo: "16:00 a 17:10",
  dice: "Límites de nuestro trabajo: una fórmula y una corrida del exacto por tamaño; planted irregular (n=30: 9,83 s más rápido que n=28: 42,5 s); una sola máquina, en Python; p = 0,5 sin ajustar; 3 fórmulas por punto en escala; aplicación hipotética. Qué haríamos: repetir y promediar; comparar con MiniSat; usar benchmarks estándar; probar otros valores de p.",
  detras: [
    "Cada límite tiene una consecuencia sobre qué se puede concluir: una corrida ⇒ sin dispersión ni intervalos de confianza; planted fáciles ⇒ no generalizar la rapidez de WalkSAT; Python ⇒ valores absolutos no comparables con solvers optimizados; 3 fórmulas ⇒ «1 de 3» no estima probabilidades.",
    "**Validez interna:** ambos algoritmos resuelven las mismas fórmulas; WalkSAT no dio falsos SAT en 180. **Validez externa:** limitada a este tipo de instancias, tamaño de muestra y hardware.",
    "**Reflexión de fondo:** SAT es intratable en el peor caso (NP-completo) pero en la práctica se resuelve a gran escala si hay estructura. Los datos propios lo ilustran: con 32 variables WalkSAT halló la solución en ms y el exacto no terminó en una hora cuando no la había.",
    "**Riesgo de heurísticas incompletas:** en aplicaciones críticas (verificación, seguridad) un UNKNOWN no es «no hay problema».",
  ],
  decir: [
    "Los límites sin disculparse y con su consecuencia: «esto nos limita a tendencias, no a leyes».",
    "Las cuatro mejoras concretas.",
    "La reflexión de fondo en dos oraciones.",
  ],
  memoria:
    "Planted: n=30 9,83 s < n=28 42,5 s · 1 fórmula y 1 corrida por tamaño del exacto · 3 fórmulas por punto en escala · p=0,5 sin ajustar · 0 falsos SAT en 180.",
  error:
    "defender los resultados como definitivos, o al revés, descalificar el trabajo entero. La postura correcta es la precisión: qué se puede concluir y qué no.",
  preg: [
    [
      "¿Qué se puede y qué no se puede concluir de sus datos?",
      "Se puede concluir: que la fuerza bruta crece ≈ ×4 cada dos variables en las insatisfacibles construidas, que WalkSAT resolvió las fórmulas planted en milisegundos y las de 5.000 variables en segundos, que no dio falsos SAT, y que UNKNOWN no es UNSAT. No se puede concluir: que WalkSAT sea siempre más rápido, que resuelva instancias estructuradas reales, que 10.000 variables sea imposible, ni nada sobre solvers CDCL o sobre la FRCU.",
    ],
    [
      "¿Por qué una sola corrida del exacto no es un problema?",
      "Es una limitación y la reconocemos. Para n grandes cada corrida dura minutos u horas. Lo que sostiene la conclusión es que la tendencia (×4 por cada +2) se repite en 6 pares consecutivos y que el conteo de asignaciones es exacto, no ruidoso. Los tiempos absolutos individuales no deberían citarse con más precisión que la tendencia.",
    ],
    [
      "¿Qué harían distinto con más tiempo?",
      "Repetir y promediar el exacto (y fijar un criterio para el tope de pasos), agregar un solver CDCL sobre los mismos DIMACS, usar benchmarks estándar con estructura, barrer p y reinicios, y medir n=10.000 con topes mayores para distinguir «difícil» de «imposible con este tope».",
    ],
  ],
});

DIAP({
  n: 15,
  titulo: "Conclusiones",
  tiempo: "17:10 a 18:40",
  dice: "Tres cifras: «> 1 h» (exacto n=32 insatisfacible no terminó; n=30 tardó 29,5 min); «1,4 ms» (WalkSAT en la fórmula con solución de n=32; exacto: 631 s); «UNKNOWN ≠ UNSAT» (la heurística gana velocidad y pierde la garantía). Cierre: exacto cuando n es chico o hay que demostrar UNSAT; WalkSAT cuando n crece y basta una solución. En la práctica se usan solvers industriales.",
  detras: [
    "Es la síntesis de todo: la tensión entre **garantía** (exacto) y **velocidad** (heurística).",
    "La regla de decisión: ¿necesito demostrar que no hay solución? ⇒ completo. ¿Basta con encontrar una? ⇒ heurística razonable. ¿Instancias reales? ⇒ solver CDCL.",
    "Tienen que poder justificar cada cifra: «> 1 h» (TIMEOUT a 3.600 s), «1,4 ms» (promedio de 10 corridas de WalkSAT en planted n=32), «UNKNOWN ≠ UNSAT».",
  ],
  decir: [
    "Las tres ideas con las tres cifras, más la regla de uso.",
    "Cerrar con los límites en una oración (una instancia por tamaño en la comparación, planted fáciles, Python, aplicación hipotética).",
  ],
  memoria:
    "> 1 h (n=32 insat.; n=30: 29,5 min) · 1,4 ms (WalkSAT n=32 planted; exacto 631 s) · UNKNOWN ≠ UNSAT · 5.000 variables en segundos (≈ 16 s de media en las 6; 9,5 s en las planted).",
  error:
    "agregar conclusiones que no están respaldadas por datos (por ejemplo, «WalkSAT es mejor que CDCL» o «ahorra X»). Las conclusiones tienen que salir de la Parte 5.",
  preg: [
    [
      "¿Cuál es la conclusión más importante?",
      "Que hay una compensación entre garantía y velocidad: el exacto es correcto y completo pero exponencial (n=32 insatisfacible superó la hora), y la heurística es rapidísima pero incompleta (UNKNOWN no es UNSAT). La elección depende de si necesito demostrar la insatisfacibilidad o me alcanza con encontrar una solución.",
    ],
    [
      "¿Cuándo usarían cada uno?",
      "El exacto, cuando n es chico o cuando hay que demostrar que no hay solución; WalkSAT, cuando n es grande y basta una solución, sabiendo que un fallo no dice nada. Para casos reales, un solver CDCL, que combina completitud con buen desempeño en instancias con estructura.",
    ],
    [
      "¿Qué les sorprendió de los resultados?",
      "Dos cosas honestas para decir: que la irregularidad de las planted fue grande (n=30 más rápido que n=28) y que la razón ≈ 4 se sostuvo tan limpia; y que 10.000 variables no se resolviera dentro del tope. Ambas muestran cuánto dependen los resultados del azar y de los parámetros.",
    ],
  ],
});

DIAP({
  n: 16,
  titulo: "Referencias y cierre",
  tiempo: "18:40 a 20:00",
  dice: "Siete referencias: Cook (1971); Selman, Kautz y Cohen (1994); Ferreiro y Marenco (UTDT, 53 JAIIO/SIIIO); Escalante, Chiesa y Fekete (SUMA 2019, UNR); Asín Achá y Nieuwenhuis (2014); Kroon, Huisman et al. (2009); SAT Competition 2024.",
  detras: [
    "Lo que respalda cada referencia: Cook ⇒ NP-completitud; Selman-Kautz-Cohen ⇒ WalkSAT; Ferreiro-Marenco ⇒ caso UTDT (ILP); Escalante et al. ⇒ coloreo UNR; Asín Achá-Nieuwenhuis ⇒ horarios con SAT/MaxSAT; Kroon-Huisman ⇒ la analogía de 50 M€; SAT Competition 2024 ⇒ Kissat.",
    "La diapositiva no lista Levin, Karp ni Biere (BMC, etc.): están en la lista completa del informe. Si preguntan por Levin, la referencia es Levin (1973), «Universal sequential search problems».",
    "**Antes de exponer, abran las fuentes de verificación pendiente** (Parte 7).",
  ],
  decir: [
    "«Las fuentes están en pantalla. No citamos nada que no hayamos podido verificar por lo menos en su resumen.» (Esto solo si de verdad lo hicieron.)",
    "Cierre breve: «Eso es todo de nuestra parte. Muchas gracias. Quedamos a disposición para las preguntas.»",
    "Dejar la diapositiva 16 en pantalla durante las preguntas.",
  ],
  memoria:
    "7 referencias en pantalla; la lista completa está en el informe (21 entradas).",
  error:
    "decir «todas las referencias están verificadas» si no abrieron las de la lista de la Parte 7. El informe reconoce que algunas se verificaron solo por resumen o metadatos.",
  preg: [
    [
      "¿Leyeron los trabajos de la UTDT y de la UNR?",
      "Respuesta honesta según lo que hayan hecho. El informe dice que se verificaron los metadatos y resúmenes («verificado por búsqueda: metadatos y resumen»). Si abrieron los PDF antes de exponer, pueden decir qué dice cada uno; si no, digan exactamente eso: «verificamos el resumen y los metadatos, no el texto completo».",
    ],
    [
      "¿De dónde sale la cifra de 50 millones de euros?",
      "Del artículo de Kroon, Huisman et al. (Interfaces, 2009) sobre el nuevo horario de los ferrocarriles neerlandeses, tomada de su resumen. Es una analogía de investigación operativa, no SAT. [Verificar el número y el período antes de decirlo.]",
    ],
    [
      "¿Por qué citan a Selman, Kautz y Cohen?",
      "Porque propusieron WalkSAT en «Noise strategies for improving local search» (AAAI-94). Es la fuente del algoritmo que implementamos.",
    ],
  ],
});

H2("4.4 Lo que Camilo tiene que saber de la mitad de Facundo");
P(
  "Si le preguntan a Camilo sobre las diapositivas 9 a 16, tiene que poder contestar a nivel básico:",
);
BS(
  "Escalabilidad: WalkSAT resuelve todo hasta 1.000 variables; con 5.000 resolvió 4 de 6; con 10.000, 0 de 6, por el tope de 5.000.000 de pasos (no prueba UNSAT).",
  "Casos reales: UTDT usa ILP (35 módulos, no SAT puro); UNR trabaja coloreo (ILP); el puente a SAT es coloreo → CNF con x(v,c).",
  "Herramientas: MiniSat, Glucose, Kissat (CDCL, completos), Z3 (SMT), OR-Tools CP-SAT; usos: BMC, dependencias, horarios.",
  "Modelo FRCU: hipotético; x(c,f); 9 variables y 19 cláusulas; sin cifra de ahorro; 50 M€ es una analogía no SAT.",
  "Viabilidad: la fuerza bruta no escala; hacen falta solvers y datos estructurados; preferencias ⇒ MaxSAT/ILP.",
  "Reflexión: una corrida por tamaño, planted fáciles, Python, p sin ajustar; mejoras: repetir, comparar con MiniSat, benchmarks, barrer p.",
);
// =====================================================================
// PARTE 5
// =====================================================================
H1("Parte 5 · Resultados explicados");
P(
  "Acá aprenden a leer cada gráfico y la tabla, y a explicar por qué los números son como son. Es la parte que más conviene tener fresca ante las preguntas: casi todas terminan en «¿y eso por qué da así?».",
);

H2("5.1 La tabla del algoritmo exacto, completa");
P(
  "Fuente: §resultados_exacto.csv§ (informe sec. 8.1). Una sola corrida por celda. «Asignaciones» = cuántas probó la fuerza bruta antes de terminar.",
);
TABLA(
  [
    "n",
    "m",
    "Insat.: estado",
    "Insat.: tiempo",
    "Insat.: asignaciones",
    "Planted: tiempo",
    "Planted: asignaciones",
    "Planted: % del espacio",
  ],
  unsat.map((r) => {
    const q = porN(planted, r.n);
    return [
      String(r.n),
      String(r.m),
      r.estado_exacto,
      r.estado_exacto === "TIMEOUT" ? "> 3.600 s (cortado)" : ft(r.t_exacto_s),
      fn(r.asignaciones),
      ft(q.t_exacto_s),
      fn(q.asignaciones),
      aprox(q.asignaciones, r.n),
    ];
  }),
  [500, 550, 1100, 1450, 1800, 1100, 1526, 1000],
  { size: 16 },
);
P(
  "**Cómo leerla.** Las insatisfacibles evalúan exactamente 2ⁿ asignaciones (columna 5; para n=32 el TIMEOUT cortó en 1.805.975.552 de 4.294.967.296). En las planted, la última columna dice qué fracción del espacio hubo que recorrer hasta dar con la primera solución: es el motivo de la irregularidad (sección 5.4).",
);

H2(
  "5.2 Por qué una recta en escala logarítmica significa crecimiento exponencial",
);
P(
  "Supongamos que el tiempo es t(n) = k · 2ⁿ. Tomando logaritmo en base 2: log₂ t(n) = log₂ k + n. Es una **recta en n** con pendiente 1 (en log₂ por cada variable). Y por cada **+2 variables** el logaritmo sube 2, es decir log₂(4) = 2: el tiempo se multiplica por 2² = 4.",
);
P(
  "En el gráfico el eje vertical usa escala logarítmica (cada marca multiplica por 10, no suma). En ese eje, una cantidad que se multiplica por un factor constante cada vez que n sube 2 avanza **la misma distancia vertical** cada vez, o sea dibuja una recta. Una función polinomial (n², n³) en ese mismo gráfico se curvaría hacia abajo; una exponencial, en cambio, queda recta. **Recta en log ⇒ exponencial.**",
);
CAJA(
  "MEMORIA",
  `La pendiente de la recta en el gráfico 1 corresponde a ×4 por cada +2 en n. Las insatisfacibles dan ×${rg} de media geométrica (entre ${rmin} y ${rmax}). Ajuste con n=30: t ≈ k·2ⁿ con k ≈ ${fn((n30.t_exacto_s / 2 ** 30) * 1e6, 2)} × 10⁻⁶ s (informe sec. 8.1).`,
);

H2("5.3 La razón t(n+2)/t(n) con los números reales");
TABLA(
  ["n → n+2", ...D.razon.todas.map((x) => `${x.n}→${x.n + 2}`)],
  [["t(n+2)/t(n)", ...D.razon.todas.map((x) => fn(x.r, 2))]],
  [1100, ...D.razon.todas.map(() => 790)],
  { size: 16 },
);
BS(
  `Para **n ≥ 18** (donde el tiempo ya es mucho mayor que el ruido del reloj): entre ${rmin} y ${rmax}; **media geométrica ${rg}**. Son 6 pares (18→20 a 28→30).`,
  "Para **n chico** (10 a 16) los tiempos son de milésimas de segundo y dominan costos fijos (arranque de la función, resolución del reloj, tiempo de Python), por eso las razones se desvían (3,32; 4,22; 4,46; 3,58). Por eso el informe usa n ≥ 18.",
  "**Por qué media geométrica:** las razones son multiplicativas; el promedio natural de factores es el geométrico (exp del promedio de logaritmos).",
  `**No hay razón para n=30→32** porque n=32 no terminó. Si la hubiera (proyección): ×${rg} ⇒ ≈ ${fn(n30.t_exacto_s * razon.gm, 0)} s ≈ ${proy32} h.`,
);
CAJA(
  "OJO",
  "Una razón de 4 también aparecería si el crecimiento fuera 2ⁿ con constantes; no distingue 2ⁿ de, digamos, 4^(n/2) (que es lo mismo). Lo que sí descarta es un polinomio: en un n^c la razón t(n+2)/t(n) = ((n+2)/n)^c tiende a 1 al crecer n; con n=30 y c=5 sería (32/30)^5 ≈ 1,38, no 4. [DERIVADO]",
);

H2("5.4 Por qué los tiempos del exacto en las planted no son monótonos");
P(
  `En una fórmula satisfacible la fuerza bruta se detiene en la **primera solución que encuentra al recorrer los enteros 0, 1, 2, …**. Si la solución «cae» temprano en ese orden, el tiempo es chico; si cae tarde, es grande. Dónde cae depende de la fórmula puntual (de su solución oculta), no del tamaño. Con los datos:`,
);
TABLA(
  [
    "n",
    "Asignaciones hasta la 1.ª solución",
    "Fracción del espacio 2ⁿ",
    "Tiempo",
  ],
  [22, 24, 26, 28, 30, 32].map((n) => {
    const q = porN(planted, n);
    return [
      String(n),
      fn(q.asignaciones),
      aprox(q.asignaciones, n),
      ft(q.t_exacto_s),
    ];
  }),
  [800, 3000, 2600, 2626],
);
BS(
  `**n=28:** ${fn(n28p.asignaciones)} asignaciones (${aprox(n28p.asignaciones, 28)} del espacio), ${ft(n28p.t_exacto_s)}. **n=30:** solo ${fn(n30p.asignaciones)} (${aprox(n30p.asignaciones, 30)}), ${ft(n30p.t_exacto_s)}. La de n=30 tuvo una solución «temprana» (pese a que su espacio es el cuádruple) y por eso tardó menos.`,
  "En el orden de enumeración x₁ es el bit menos significativo y x_n el más significativo. Las soluciones con x_n = F caen en la **primera mitad** del orden; las que tienen x_n = V, en la segunda. Si la solución oculta empieza (en los bits altos) con muchas variables en F, la primera solución está muy cerca del principio.",
  "No es un error del programa ni del cronómetro: es un efecto del orden de enumeración. Por eso **para medir el crecimiento del exacto se usan las insatisfacibles**, donde siempre se recorre todo y el tiempo es una curva limpia.",
  "El número a retener: «n=30 (9,83 s) más rápido que n=28 (42,5 s)». Está también en la diapositiva 14.",
);

H2("5.5 Cómo leer cada gráfico");
H3(
  "Gráfico 1 · Tiempo de la fuerza bruta según n (diapositiva 5; informe sec. 8.1)",
);
BS(
  "**Ejes:** x = n (10 a 32); y = tiempo en segundos, **escala logarítmica**.",
  "**Series:** insatisfacible (círculos azules, línea continua) y satisfacible/planted (cuadrados naranja, línea de trazos). La **x roja** en n=32 es el TIMEOUT: se dibuja en y = 3.600 pero **no es un tiempo medido**.",
  "**Qué ver:** la serie insatisfacible es casi una recta (exponencial, ×4 cada +2). La planted está por debajo y es irregular (sección 5.4).",
  "**Qué decir:** «cada dos variables el tiempo se multiplica por cuatro; con 32 no terminó».",
);
H3("Gráfico 2 · Tiempo medio de WalkSAT según n (diapositiva 9; informe sec. 8.4)");
BS(
  "**Ejes:** x = n (100 a 10.000, escala log); y = tiempo medio (log), promedio de las 3 fórmulas por punto.",
  `**Series:** planted (m/n=4,26) y aleatoria (m/n=4,0). Valores medios: planted 2 ms, 165 ms, 280 ms, 9,46 s, 26,8 s; aleatoria 3 ms, 244 ms, 394 ms, 22,6 s, 25,3 s.`,
  "**Qué ver:** crece con n; en n=10.000 se aplana en ≈ 26 s porque el tiempo está acotado por el tope de 5.000.000 de pasos (≈ 5 µs por paso): **no es una propiedad de WalkSAT sino del límite**. No confundan esa meseta con «WalkSAT se estabiliza».",
);
H3(
  "Gráfico 3 · Exacto vs. WalkSAT sobre las mismas fórmulas (diapositiva 7; informe sec. 8.2)",
);
BS(
  "**Tres series:** fuerza bruta satisfacible, fuerza bruta insatisfacible y WalkSAT satisfacible; más la x roja del TIMEOUT.",
  "**Qué ver:** WalkSAT está pegado al piso (≤ 1,4 ms) mientras el exacto sube varios órdenes de magnitud. La brecha es de ≈ 4,4 × 10⁵ en n=32 planted.",
  "**Qué NO hay:** una serie de WalkSAT en insatisfacibles, porque su tiempo (≈ 392 ms) es solo el costo de agotar los 100.000 pasos.",
);
H3("Gráfico 4 · Éxito de WalkSAT según n (informe sec. 8.4)");
BS(
  "**Eje y:** % de corridas (de 3 por punto) que encontraron solución. Con 3 muestras los únicos valores posibles son 0, 33, 67 y 100 %.",
  "**Qué ver:** planted: 100 % hasta n=5.000 y 0 % en 10.000. Aleatoria: 100 % hasta 1.000, 33 % en 5.000 y 0 % en 10.000.",
  "**Cuidado:** con 3 muestras, «33 %» es «1 de 3»; no es una estimación fina.",
);
H3("Gráfico 5 · Asignaciones evaluadas por la fuerza bruta (informe sec. 8.1)");
BS(
  "**Insatisfacible:** exactamente 2ⁿ (de 1.024 a 1.073.741.824), una recta perfecta en escala log. La x roja en n=32 está en 1.805.975.552: lo evaluado hasta cortar.",
  "**Satisfacible:** menos y más irregular; es el efecto de la 1.ª solución (sección 5.4).",
  "**Por qué importa:** este gráfico es la prueba más limpia del 2ⁿ: no depende del cronómetro.",
);
H3("Gráfico 6 · Cláusulas insatisfechas por iteración (informe sec. 8.4)");
BS(
  `**Qué es:** una corrida de WalkSAT en una fórmula aleatoria con n=500 y m=2.000 (${fn(hist.pasos)} pasos). Eje x = paso; eje y = cantidad de cláusulas insatisfechas.`,
  `**Números:** empieza con ${hist.inicial} insatisfechas (≈ m/8 = 250: la asignación inicial al azar rompe 1/8 de las cláusulas), baja a 96 en el paso 100, a 23 en el paso 1.000, y primero baja de 10 en el paso 1.593. Después **oscila cerca de cero** y recién en el paso ${fn(hist.pasos)} llega a ${hist.final}.`,
  "**Lección:** la mayor parte del trabajo es rápido (descenso en los primeros cientos de pasos); lo difícil es el tramo final, donde queda un puñado de cláusulas y hay que encontrar la combinación justa. Por eso el ruido importa.",
);

H2("5.6 Por qué WalkSAT 0/10 en UNSAT no es un fallo");
BS(
  "En una fórmula insatisfacible **no existe** ninguna asignación que cumpla todo. Un algoritmo no puede «encontrar» lo que no existe.",
  "Lo que hace WalkSAT es lo máximo posible: devuelve la mejor asignación que vio (1 cláusula insatisfecha = óptimo de Max-SAT) y el veredicto UNKNOWN.",
  "UNKNOWN significa «no encontré solución en mis pasos». UNSAT significa «demostré que no hay». WalkSAT **no puede demostrar** porque solo prueba asignaciones sueltas; demostrar la inexistencia exige cubrir todo el espacio (o una prueba por resolución).",
  `Medimos 0/10 en las 12 insatisfacibles (120 corridas, todas UNKNOWN). Todas gastaron los 100.000 pasos (entre ${fn(Math.min(...D.walkUnsatT) * 1000, 0)} y ${fn(Math.max(...D.walkUnsatT) * 1000, 0)} ms; media ≈ ${fn(walkUnsatMed * 1000, 0)} ms).`,
);
CAJA(
  "MEMORIA",
  "Frase para la oral: «Que WalkSAT devuelva UNKNOWN en una insatisfacible no es un error: es el comportamiento correcto de un algoritmo incompleto. Lo que sería un error es que devolviera SAT, y eso no pasó nunca.»",
);

H2("5.7 Por qué falla en n = 10.000");
BS(
  "Las 6 corridas (3 planted + 3 aleatorias) terminaron en UNKNOWN **por el tope de 5.000.000 de pasos**. No por el límite de 120 s (llegaron a 24,5 a 27,4 s).",
  "La mejor asignación de cada corrida dejó entre 1 y 24 cláusulas sin cumplir de 40.000 a 42.600 (≥ 99,94 %). Está «casi» resuelta.",
  "Por qué cuesta más: la cantidad de pasos necesarios crece con n (en n=5.000 ya hacen falta 1,6 a 3,2 millones de pasos para las que se resuelven; en n=10.000 el tope queda corto). Además la caminata final entre unas pocas cláusulas insatisfechas es ruidosa.",
  "Las aleatorias de la escala (m/n=4,0) están cerca de la zona difícil, donde las instancias aleatorias grandes son duras para la búsqueda local [LITERATURA].",
  "**Qué NO se puede decir:** que las fórmulas son insatisfacibles; que WalkSAT «no puede» con 10.000 variables; que «el tope de 5.000.000 es suficiente». Se puede decir: «no alcanzó dentro del tope que fijamos, y no sabemos qué pasaría con más pasos, otro p o reinicios».",
);

H2("5.8 Qué se puede y qué no se puede concluir");
TABLA(
  ["Se puede concluir (con respaldo)", "No se puede concluir (no respaldado)"],
  [
    [
      "La fuerza bruta es exponencial: el trabajo crece ×4 cada +2 variables (insatisfacibles construidas, n=18 a 30).",
      "Que 2ⁿ sea el mejor posible para SAT (hay algoritmos exactos mucho mejores, y P vs NP está abierto).",
    ],
    [
      "El exacto con n=32 insatisfacible superó 1 hora (en esta máquina, en Python).",
      "Cuánto habría tardado en terminar (solo hay proyección).",
    ],
    [
      "WalkSAT resolvió 10/10 las planted de n=10 a 32 en ≤ 1,4 ms y las de n=5.000 en 6,7 a 12,4 s.",
      "Que WalkSAT sea «siempre» más rápido o que resuelva instancias reales estructuradas.",
    ],
    [
      "WalkSAT nunca dijo SAT sobre una insatisfacible (0 de 180 + 12 construidas).",
      "Que no pueda ocurrir (depende de que el código sea correcto; se verificó empíricamente).",
    ],
    [
      "En n=10.000 no encontró solución dentro de 5.000.000 de pasos.",
      "Que las fórmulas sean insatisfacibles ni que sea imposible para WalkSAT.",
    ],
    [
      "En n=16, la fracción de fórmulas satisfacibles cae de 100 % a 3 % al pasar de m/n=2 a 6.",
      "Que el umbral esté en 4,26 (eso es la literatura, no este experimento).",
    ],
    [
      "El caso UTDT usa ILP y el modelo FRCU es una hipótesis con ejemplo de 9 variables.",
      "Que la FRCU use SAT, que ahorre algo, o cuánto.",
    ],
  ],
  [4513, 4513],
);
P(
  "**Limitaciones que condicionan todo** (informe sec. 11): una sola máquina y un solo lenguaje (Python), una corrida del exacto por tamaño, una fórmula de cada tipo por n, 3 fórmulas por punto en la escala, WalkSAT con un solo p y un solo reinicio, sin solver CDCL ni benchmarks estándar.",
);
// =====================================================================
// PARTE 6
// =====================================================================
H1("Parte 6 · Banco de preguntas difíciles de la cátedra");
P(
  "32 preguntas con respuesta modelo. Están agrupadas por tema. Para estudiar: tapen la respuesta, contesten en voz alta y comparen. Las respuestas están escritas para decirse, no para leerse: úsenlas como esqueleto. Entre corchetes dice quién es el responsable natural según la diapositiva.",
);

H2("6.1 Teoría");
QA("1. ¿Por qué SAT es NP-completo? [Camilo]", [
  "Hay que probar dos cosas. Primero, que está en NP: dada una asignación, verifico en tiempo polinomial que cada cláusula tiene un literal verdadero. Segundo, que es NP-difícil: es el teorema de Cook-Levin. Dado cualquier lenguaje L en NP, existe una máquina de Turing no determinística que lo decide en tiempo polinomial; la ejecución de esa máquina sobre una entrada w se describe con una tabla de tamaño polinomial, y se escriben cláusulas que dicen que la tabla es una computación válida y aceptadora. Esa fórmula es satisfacible si y solo si la máquina acepta w. La construcción es polinomial, así que L se reduce a SAT.",
]);
QA("2. ¿Qué implicaría P = NP? [Camilo]", [
  "Que todos los problemas de NP se podrían resolver en tiempo polinomial, porque basta un algoritmo polinomial para SAT (todo NP se reduce a él). Cambiaría la optimización, la planificación y la verificación, y rompería la criptografía de clave pública tal como hoy se la concibe. Es un problema abierto; se cree que P ≠ NP, pero no está demostrado. Nuestros datos (la fuerza bruta es exponencial) no dicen nada sobre esto: miden un algoritmo, no todos los posibles.",
]);
QA("3. ¿Qué es una reducción? Den un ejemplo. [Camilo/Facundo]", [
  "Una función calculable en tiempo polinomial que transforma instancias de un problema A en instancias de B, de modo que la respuesta sea la misma. Si sé resolver B, resuelvo A. Ejemplos: SAT a 3-SAT (partir cláusulas largas con variables nuevas) y k-coloreo a SAT (variables x(v,c)). Se usan para demostrar dificultad (si A es NP-difícil, B también) o para aprovechar un solver existente.",
]);
QA("4. ¿Cuál es la diferencia entre NP-difícil y NP-completo? [Camilo]", [
  "NP-completo = NP-difícil y además pertenece a NP. NP-difícil es al menos tan difícil como todo NP pero puede no estar en NP: por ejemplo, Max-SAT (optimización) o el problema de la parada son NP-difíciles y no son NP-completos.",
]);
QA("5. ¿Por qué 3-SAT es NP-completo y 2-SAT no? [Camilo]", [
  "2-SAT se resuelve en tiempo polinomial: cada cláusula (a ∨ b) equivale a dos implicaciones (¬a→b y ¬b→a); se arma un grafo y la fórmula es insatisfacible si y solo si alguna variable está en la misma componente fuertemente conexa que su negación. Con tres literales ese truco deja de funcionar, y 3-SAT es NP-completo porque SAT se reduce a él.",
]);
QA("6. ¿Existe algún algoritmo para 3-SAT mejor que 2ⁿ? [Camilo]", [
  "Sí: hay algoritmos de peor caso exponencial pero con base menor, del orden de 1,3ⁿ (como el de Schöning, que es en esencia una caminata aleatoria, o PPSZ) [LITERATURA, VERIFICAR]. Siguen siendo exponenciales. Y en la práctica los solvers CDCL resuelven instancias con estructura de millones de variables aunque el peor caso sea exponencial. La Exponential Time Hypothesis conjetura que no hay algoritmo subexponencial para 3-SAT.",
]);
QA("7. ¿Qué es el umbral m/n ≈ 4,26? ¿Lo verificaron? [Camilo]", [
  "Es un resultado conocido de la literatura: en 3-SAT aleatorio, con m = α·n cláusulas, para α menor que ≈ 4,26 casi todas las fórmulas son satisfacibles y para α mayor casi todas son insatisfacibles; cerca del umbral, las instancias son más difíciles. **No lo verificamos.** Lo usamos para elegir m = round(4,26·n) y lo que sí medimos es cuántas de 30 fórmulas de n=16 eran satisfacibles para m/n de 2 a 6 (de 30 a 1). Con n=16 la transición es suave; el umbral nítido necesita n grande.",
]);
QA("8. ¿Qué es Max-SAT y para qué lo usaron? [Camilo]", [
  "Max-SAT busca la asignación que cumple más cláusulas. Lo usamos solo para medir qué tan cerca queda WalkSAT cuando no encuentra solución: error_relativo = cláusulas insatisfechas de la mejor asignación dividido m. No lo resolvimos como problema de optimización.",
]);

H2("6.2 Algoritmos y código");
QA("9. ¿Cómo saben que fuerza_bruta es correcta? [Camilo]", [
  "Por varias evidencias independientes, ya que no hay una suite de tests: (1) en las fórmulas construidas, devuelve SAT en las 12 planted y UNSAT en las 11 insatisfacibles que terminaron, coincidiendo con lo que construimos; (2) en cada UNSAT, la cantidad de asignaciones evaluadas es exactamente 2ⁿ; (3) concuerda con WalkSAT en las 180 fórmulas de densidad (en las 111 satisfacibles WalkSAT encontró solución, y nunca dijo SAT sobre una insatisfacible); (4) verificamos con un chequeo independiente que los modelos devueltos cumplen todas las cláusulas. Es evidencia empírica; no una demostración formal.",
]);
QA("10. ¿Puede WalkSAT dar un falso SAT? [Camilo]", [
  "No, no puede, salvo un error de programación. Declara SAT solo cuando no quedan cláusulas insatisfechas, es decir, cuando la asignación realmente las cumple a todas. Por eso es correcto aunque sea incompleto. Lo comprobamos: 0 falsos SAT en 180 fórmulas de densidad, 0 en las 12 insatisfacibles construidas (120 corridas), y los 120 modelos devueltos en las planted pasan un verificador independiente. Un matiz honesto: el código mantiene la lista de insatisfechas de manera incremental y no vuelve a validar con un recorrido completo.",
]);
QA("11. ¿Qué es DPLL / CDCL? [Facundo/Camilo]", [
  "DPLL (Davis-Putnam-Logemann-Loveland) es una búsqueda con retroceso: se elige una variable y se le da un valor; se propagan las consecuencias (si una cláusula queda con un solo literal sin falsificar, ese literal queda forzado: propagación de unidades); si hay conflicto, se vuelve atrás y se prueba el otro valor. CDCL (Conflict-Driven Clause Learning) lo mejora: ante cada conflicto analiza su causa (grafo de implicaciones), aprende una cláusula nueva que evita repetir ese error, vuelve atrás varios niveles de una vez, reinicia periódicamente y usa heurísticas de decisión (VSIDS) y estructuras eficientes (watched literals). Es completo: responde SAT o UNSAT. MiniSat, Glucose y Kissat son CDCL.",
]);
QA("12. ¿Por qué no usaron un solver CDCL? [Facundo]", [
  "Porque el trabajo compara un exacto y una heurística implementados para el TP; un CDCL cambia de liga y la comparación pasa a otra cosa. Lo reconocemos como limitación y como siguiente paso: correr MiniSat o Kissat sobre los mismos archivos DIMACS ya generados. No tenemos medición propia de un CDCL, por eso no afirmamos nada sobre su desempeño en nuestras fórmulas. [Mi razonamiento, no medido: es probable que lo resuelva en tiempo despreciable, porque el núcleo insatisfacible es evidente; pero es una hipótesis, no un resultado.]",
]);
QA("13. ¿Cómo funciona WalkSAT? Explíquenlo con un ejemplo. [Camilo]", [
  "Parte de una asignación al azar. Mientras haya cláusulas insatisfechas, elige una al azar, calcula para cada variable de la cláusula cuántas cláusulas hoy satisfechas se romperían al darla vuelta (break count). Si alguna rompe 0, la da vuelta. Si no, con probabilidad p da vuelta una cualquiera y con 1−p la que menos rompe. Repite hasta no tener cláusulas insatisfechas o agotar los pasos. (Ejemplo: Parte 2, secciones 2.8.)",
]);
QA(
  "14. ¿Por qué WalkSAT no puede probar que una fórmula es insatisfacible? [Camilo]",
  [
    "Porque explora asignaciones sueltas guiado por el azar; que no encuentre solución en N pasos no implica que no exista. Probar UNSAT exige cubrir todo el espacio de búsqueda o producir una prueba (por ejemplo por resolución), que es lo que hacen los algoritmos completos. Por eso su respuesta posible es SAT o UNKNOWN, y UNKNOWN ≠ UNSAT.",
  ],
);
QA(
  "15. ¿Por qué el código usa un entero para la asignación en la fuerza bruta y una lista en WalkSAT? [Camilo]",
  [
    "En la fuerza bruta hay que recorrer las 2ⁿ asignaciones: contar de 0 a 2ⁿ−1 las recorre todas sin repetir y evita armar listas en cada paso. En WalkSAT hay que dar vuelta una variable por vez y consultar su valor muchas veces: una lista indexada por variable es lo natural.",
  ],
);
QA(
  "16. ¿Por qué el tiempo de la fuerza bruta no depende tanto de m como dice O(2ⁿ·m·k)? [Camilo]",
  [
    "Porque O(2ⁿ·m·k) es el peor caso. En la práctica una asignación al azar rompe una cláusula 3-SAT con probabilidad 1/8, así que la fuerza bruta descarta la asignación al revisar en promedio unas 8 cláusulas, no m. En nuestros datos: ≈ 7,9 cláusulas por asignación en n=30 insatisfacible. Sigue siendo exponencial; cambia la constante.",
  ],
);

H2("6.3 Experimentos y resultados");
QA("17. ¿Qué pasa si cambia m/n? [Camilo]", [
  "Cambia la proporción de fórmulas satisfacibles y la dificultad. Con pocas cláusulas por variable casi todas son satisfacibles y fáciles; con muchas, casi todas insatisfacibles. Cerca de 4,26 (literatura) hay mezcla y las instancias son más duras. En n=16 y 30 fórmulas por densidad medimos 30/30 satisfacibles en m/n=2 y 3; 24 en 4,0; 18 en 4,26; 8 en 5,0 y 1 en 6,0. WalkSAT encontró solución en todas las satisfacibles y no dio falsos SAT.",
]);
QA("18. ¿Llegaron a una hora de ejecución? [Camilo]", [
  "Sí: el exacto con n=32 insatisfacible no terminó en una hora; lo cortamos a 3.600 s con el 42 % de las asignaciones evaluadas. El último que terminó fue n=30 (1.768 s, 29,5 min). No hay un tiempo medido para n=32; la proyección con la razón ×4 da ≈ 2,0 h, y con la velocidad observada, ≈ 2,4 h. Son proyecciones.",
]);
QA(
  "19. ¿Por qué una sola corrida por punto? ¿Es estadísticamente válido? [Camilo/Facundo]",
  [
    "Es una limitación que declaramos. Para n grande cada corrida dura minutos u horas. Los tiempos individuales son ilustrativos; lo que sostiene la conclusión es la tendencia, que se repite en 6 pares consecutivos de tamaños (razón entre 3,8 y 4,3), y el conteo de asignaciones evaluadas, que es exacto y no depende del cronómetro. Con más tiempo repetiríamos y promediaríamos.",
  ],
);
QA("20. ¿Los tiempos serían iguales en otra computadora? [Camilo]", [
  "Los absolutos no: dependen de la CPU, el sistema, la versión de Python y la carga. La forma de la curva (×4 por cada +2 variables) sí. Y los conteos (éxitos, asignaciones, cláusulas evaluadas) son idénticos, porque las fórmulas y las semillas están fijas. Por eso registramos el equipo en hardware.txt.",
]);
QA(
  "21. ¿Por qué con n=30 la fórmula planted tardó menos que con n=28? [Camilo]",
  [
    "Porque en una satisfacible la fuerza bruta se detiene en la primera solución del orden 0, 1, 2…, y esa posición depende de la fórmula puntual. En n=30 la primera solución estuvo en el 0,6 % del espacio (6.300.215 asignaciones); en n=28, en el 9,5 % (25.495.706). No es un error: es la razón por la que medimos el crecimiento con las insatisfacibles.",
  ],
);
QA(
  "22. ¿Por qué falla WalkSAT en 10.000 variables? ¿Mejoraría con más pasos? [Facundo]",
  [
    "Todas las corridas se detuvieron por el tope de 5.000.000 de pasos, dejando como mucho 0,06 % de las cláusulas sin cumplir. No sabemos si con más pasos, otro p o reinicios lo resolvería: no lo probamos, y lo señalamos como trabajo futuro. Lo único que podemos afirmar es que no llegó dentro de nuestro tope. No es una prueba de insatisfacibilidad.",
  ],
);
QA(
  "23. ¿Por qué definen el error relativo así, y qué significa n.a.? [Camilo]",
  [
    "En optimización el error relativo es (valor − óptimo)/óptimo. En SAT, que es de decisión, el óptimo de una satisfacible es 0 insatisfechas y esa división no sirve. Usamos Max-SAT como vara: error_relativo = insatisfechas de la mejor asignación / m. Vale 0 cuando resolvió. «n.a.» significa que el óptimo no se conoce con los datos (insatisfacibles, donde el exacto solo prueba que no hay solución perfecta, o TIMEOUT). Por construcción sabemos que en nuestras insatisfacibles el óptimo es 1 insatisfecha, y WalkSAT lo alcanzó; pero el programa no lo verifica, así que no lo afirmamos en la columna.",
  ],
);
QA(
  "24. ¿Qué es una fórmula planted y por qué es fácil? ¿Y la insatisfacible controlada? [Camilo]",
  [
    "Planted: se elige una solución oculta y cada cláusula se corrige para que la cumpla; garantiza satisfacibilidad, pero sesga las cláusulas hacia esa solución, y por eso WalkSAT la halla en pocos pasos. Insatisfacible controlada: 8 cláusulas con todas las combinaciones de signos sobre 3 variables (cada asignación de esas 3 variables viola exactamente una) más cláusulas aleatorias; es UNSAT sin importar el resto. Tiene un núcleo obvio, por eso no es una instancia «difícil» típica, aunque a la fuerza bruta igual le cuesta 2ⁿ.",
  ],
);

H2("6.4 Aplicación");
QA("25. ¿Por qué un caso ILP en un TP de SAT? [Facundo]", [
  "Porque es el caso argentino de horarios mejor documentado y lo presentamos explícitamente como no-SAT. Sirve para mostrar el problema de fondo (NP-completo), que sus restricciones son lógicas y codificables en CNF, y que el coloreo de la UNR es el puente concreto. No afirmamos que ese caso use SAT.",
]);
QA("26. ¿Cómo se codifica un coloreo de grafos en CNF? [Facundo]", [
  "Variables x(v,c): el vértice v lleva el color c. Cláusulas: (1) cada vértice, al menos un color: (x(v,1) ∨ … ∨ x(v,k)); (2) cada vértice, a lo sumo uno: (¬x(v,c) ∨ ¬x(v,d)) para c < d; (3) para cada arista (u,v) y cada color c: (¬x(u,c) ∨ ¬x(v,c)). Total: n·k variables y n + n·k(k−1)/2 + e·k cláusulas (n vértices, e aristas). Satisfacible ⇔ k-coloreable. (La familia (2) puede omitirse: si un vértice quedara con dos colores, se elige uno y sigue siendo válido.) Ejemplo: un triángulo con 3 colores tiene 9 variables y 21 cláusulas.",
]);
QA("27. ¿Por qué no hay ahorro económico? [Facundo]", [
  "Porque no existe una cifra pública atribuible a SAT para horarios, y no la vamos a inventar. La UTDT reporta solo un impacto cualitativo. La cifra de 50 millones de euros por año es del horario ferroviario neerlandés (Kroon, Huisman et al., 2009), obtenida con investigación operativa, y la presentamos únicamente como analogía. Proponemos indicadores para medir un beneficio real: conflictos evitados, huecos de docentes, ocupación de aulas y casos sin solución.",
]);
QA("28. ¿La FRCU usa SAT? ¿Por qué un modelo hipotético? [Facundo]", [
  "No encontramos evidencia pública de que lo haga. Lo que sabemos es que publica horarios en PDF con comisiones y franjas. Modelamos cómo se podría codificar (x(c,f), al menos una franja, a lo sumo una, conflictos) con un ejemplo de 9 variables y 19 cláusulas, resuelto por ambos algoritmos, y lo rotulamos hipotético porque no usamos datos reales de la institución.",
]);

H2("6.5 Formato, herramientas y contexto");
QA("29. ¿Qué es DIMACS y para qué lo usaron? [Camilo]", [
  "Es el formato de texto estándar para fórmulas CNF: líneas «c» de comentario, una línea «p cnf n m» y una cláusula por línea de enteros terminada en 0. Lo usamos para guardar las fórmulas exactas que resolvieron ambos algoritmos (carpeta instancias/), de modo que se pueda repetir todo y probarlas con solvers externos.",
]);
QA("30. ¿Quién escribió el código? ¿Cómo trabajaron? [Ambos]", [
  "Respondan con la verdad y con seguridad. Lo importante es que puedan explicar y defender cada parte. Si usaron herramientas de IA para generar código o texto, y la cátedra lo admite, dígalo con naturalidad: «usamos herramientas de IA para generar el código y los documentos, y nuestro trabajo fue estudiarlos, ejecutarlos, verificar los resultados y poder explicarlos». Si no están seguros de que el reglamento lo permita, consúltenlo con el docente **antes** de la exposición. Lo que no deben hacer es decir que lo programaron línea por línea si no es cierto: el informe (sección 7.4) y el README ya declaran el uso de IA (Claude); mantengan esa misma versión en la oral.",
]);
QA("31. ¿Por qué Python y no un lenguaje compilado? [Camilo]", [
  "Por claridad y reproducibilidad; el objetivo era estudiar el comportamiento, no batir un récord. Python es lento (≈ 600.000 asignaciones por segundo), y por eso el límite de una hora se alcanza en n=32. Un lenguaje compilado llegaría a más variables con la misma curva exponencial: solo se desplazaría el punto donde se corta.",
]);
QA("32. ¿Qué harían distinto? [Facundo]", [
  "Repetir y promediar cada medición del exacto con varias fórmulas por tamaño; agregar un solver CDCL (MiniSat o Kissat) sobre los mismos DIMACS; probar instancias con estructura (codificación de coloreo o de horarios) y benchmarks estándar; barrer p y los reinicios; y medir n=10.000 con topes mayores para distinguir «difícil» de «imposible con este tope».",
]);
// =====================================================================
// PARTE 7
// =====================================================================
H1("Parte 7 · Glosario, autoevaluación, checklist y qué NO decir");

H2("7.1 Glosario");
TABLA(
  ["Término", "Significado"],
  [
    [
      "**Asignación**",
      "Un valor V/F para cada variable. Con n variables hay 2ⁿ.",
    ],
    [
      "**Break count**",
      "Cuántas cláusulas hoy satisfechas quedarían insatisfechas si se da vuelta una variable. Lo usa WalkSAT para decidir.",
    ],
    [
      "**Bounded model checking (BMC)**",
      "Técnica de verificación: se codifica «existe un error en los primeros k pasos» como fórmula SAT; si es SAT, la asignación es el contraejemplo.",
    ],
    [
      "**CDCL**",
      "Conflict-Driven Clause Learning: algoritmo completo de los solvers modernos; aprende cláusulas de los conflictos.",
    ],
    [
      "**Certificado**",
      "Evidencia corta que permite verificar una respuesta SÍ en tiempo polinomial (en SAT, una asignación).",
    ],
    [
      "**Cláusula**",
      "OR de literales; se cumple si al menos uno es verdadero.",
    ],
    ["**CNF**", "Forma normal conjuntiva: AND de cláusulas."],
    [
      "**Completo / incompleto**",
      "Un algoritmo es completo si siempre puede responder (incluso UNSAT) dado tiempo suficiente. WalkSAT es incompleto.",
    ],
    [
      "**coNP**",
      "Clase de problemas cuyas respuestas NO tienen certificado corto verificable. UNSAT está en coNP.",
    ],
    ["**Cook-Levin**", "Teorema: SAT es NP-completo."],
    ["**Densidad (α = m/n)**", "Cantidad de cláusulas por variable."],
    ["**DIMACS**", "Formato de texto estándar para fórmulas CNF."],
    [
      "**DPLL**",
      "Búsqueda con retroceso y propagación de unidades (Davis-Putnam-Logemann-Loveland).",
    ],
    [
      "**Error relativo (en el TP)**",
      "Cláusulas insatisfechas de la mejor asignación de WalkSAT dividido m.",
    ],
    [
      "**Exacto**",
      "Algoritmo que siempre da la respuesta correcta (y que es completo).",
    ],
    [
      "**Flip**",
      "Dar vuelta el valor de una variable. Un «paso» de WalkSAT es un flip.",
    ],
    [
      "**Free move**",
      "Variable con break count 0: darla vuelta no rompe nada. WalkSAT la elige siempre.",
    ],
    ["**Fuerza bruta**", "Probar todas las asignaciones: O(2ⁿ·m·k)."],
    ["**Heurística**", "Método práctico sin garantía de éxito."],
    [
      "**ILP**",
      "Programación lineal entera: variables enteras, restricciones y objetivo lineales.",
    ],
    [
      "**Insatisfacible (UNSAT)**",
      "No existe ninguna asignación que cumpla la fórmula.",
    ],
    ["**Literal**", "Una variable o su negación."],
    [
      "**Max-SAT**",
      "Cumplir la mayor cantidad posible de cláusulas (optimización).",
    ],
    [
      "**Mínimo local**",
      "Estado donde cualquier cambio chico empeora, pero que no es la solución.",
    ],
    ["**NP**", "Problemas verificables en tiempo polinomial."],
    ["**NP-completo**", "En NP y NP-difícil."],
    ["**NP-difícil**", "Al menos tan difícil como todo NP."],
    [
      "**Núcleo insatisfacible**",
      "En el TP, las 8 cláusulas sobre 3 variables que hacen UNSAT a la fórmula.",
    ],
    ["**P**", "Problemas resolubles en tiempo polinomial."],
    [
      "**Planted**",
      "Fórmula construida alrededor de una solución oculta; satisfacible por construcción.",
    ],
    [
      "**Propagación de unidades**",
      "Si una cláusula queda con un solo literal posible, ese literal queda forzado.",
    ],
    [
      "**Reducción polinomial**",
      "Traducción en tiempo polinomial que preserva la respuesta.",
    ],
    ["**Reinicio**", "Volver a empezar una búsqueda con una asignación nueva."],
    [
      "**Satisfacible (SAT)**",
      "Existe al menos una asignación que cumple la fórmula.",
    ],
    [
      "**Semilla**",
      "Número que fija el generador de azar para que los resultados sean reproducibles.",
    ],
    [
      "**SMT**",
      "SAT Modulo Theories: SAT con teorías (aritmética, vectores…), como Z3.",
    ],
    [
      "**Tableau**",
      "Tabla que representa la ejecución de una máquina de Turing en la prueba de Cook-Levin.",
    ],
    [
      "**TIMEOUT**",
      "El algoritmo exacto no terminó dentro del límite de tiempo; no hay tiempo medido.",
    ],
    [
      "**Tseitin**",
      "Transformación a CNF con variables auxiliares, de tamaño lineal.",
    ],
    [
      "**Umbral de fase**",
      "Densidad donde las fórmulas aleatorias pasan de casi siempre SAT a casi siempre UNSAT (≈ 4,26 en 3-SAT).",
    ],
    [
      "**UNKNOWN**",
      "Veredicto de WalkSAT cuando no encontró solución: no significa UNSAT.",
    ],
  ],
  [2500, 6526],
);

H2("7.2 Autoevaluación");
P(
  "Diez preguntas por persona. Contesten **sin mirar**, en voz alta. Las respuestas están al final de esta sección (no las miren antes).",
);
H3("Camilo: 10 preguntas");
BS(
  "C1. Dada F = (x1 ∨ ¬x2 ∨ x3) ∧ (¬x1 ∨ x2 ∨ x4) ∧ (x2 ∨ ¬x3 ∨ ¬x4), ¿la asignación x1=V, x2=F, x3=F, x4=F la satisface? ¿Por qué?",
  "C2. ¿Cuántas asignaciones hay con n=20 y con n=30? ¿Y cómo se escribe la complejidad de la fuerza bruta, letra por letra?",
  "C3. ¿Por qué se mide el algoritmo exacto con fórmulas insatisfacibles?",
  "C4. ¿Cuánto tardó n=30 insatisfacible y qué número sale de dividir los tiempos de dos tamaños consecutivos (n y n+2)?",
  "C5. ¿Qué significa la x roja de n=32 y qué cosas NO podés decir de ese punto?",
  "C6. Describí un paso de WalkSAT: qué elige primero y las tres posibilidades para la variable.",
  "C7. ¿Qué diferencia hay entre UNKNOWN y UNSAT? ¿Quién puede decir cada una?",
  "C8. ¿Qué es el error relativo en el TP y por qué en las insatisfacibles vale 1/m?",
  "C9. ¿Por qué las 8 cláusulas sobre 3 variables hacen insatisfacible la fórmula?",
  "C10. ¿Por qué n=30 planted (9,83 s) es más rápida que n=28 planted (42,5 s)?",
);
H3("Facundo: 10 preguntas");
BS(
  "F1. ¿Qué pasó en n=10.000, por qué se detuvo y qué NO se puede decir?",
  "F2. En n=5.000, ¿cuántos éxitos hubo por tipo de fórmula y qué conclusión es válida y cuál no?",
  "F3. ¿Por qué el tiempo medio se aplana en ≈ 26 s para n=10.000?",
  "F4. ¿Qué método usa el caso UTDT, cuántos módulos tiene el horario y por qué se presenta en un TP de SAT?",
  "F5. Codificá en CNF «un triángulo se puede colorear con 3 colores»: cantidad de variables y de cláusulas.",
  "F6. Desglosá las 19 cláusulas del ejemplo FRCU (3 comisiones × 3 franjas).",
  "F7. Explicá CDCL en tres frases, y por qué es completo.",
  "F8. ¿Por qué las preferencias de horarios (evitar huecos) llevan a MaxSAT o ILP y no a SAT puro?",
  "F9. ¿Qué cifra económica aparece en el TP, de dónde viene y cómo se rotula?",
  "F10. Nombrá dos limitaciones del TP y dos mejoras, y decí cuál de las limitaciones afecta más a las conclusiones sobre WalkSAT.",
);
H3("Respuestas de autoevaluación");
P("**Camilo**", { after: 40 });
BS(
  "C1. **No.** La cláusula 2, (¬x1 ∨ x2 ∨ x4), tiene ¬x1 = F (x1 es V), x2 = F y x4 = F: ningún literal verdadero. (La cláusula 1 se cumple por x1; la 3 por ¬x3.)",
  "C2. n=20: 2²⁰ = 1.048.576. n=30: 2³⁰ = 1.073.741.824. O(2ⁿ·m·k): 2ⁿ asignaciones, m cláusulas por asignación en el peor caso, k literales por cláusula.",
  "C3. Porque no tienen solución: el exacto debe recorrer las 2ⁿ asignaciones antes de poder contestar UNSAT, así el tiempo mide el trabajo máximo. En una satisfacible se detiene en la primera solución y depende de la suerte.",
  "C4. 1.768 s (29,5 min) con 1.073.741.824 asignaciones. El cociente t(n+2)/t(n) ≈ 4,0 (de 3,8 a 4,3 entre n=18 y n=30), que es 2².",
  "C5. Que el exacto se cortó a los 3.600 s sin terminar, con el 42 % de las asignaciones evaluadas. NO se puede decir cuánto tardó, ni que «tardó una hora», ni que tardaría exactamente 2 horas (eso es proyección).",
  "C6. Elige al azar una cláusula insatisfecha. Para cada una de sus variables calcula el break count. (a) Si alguna tiene break 0: la da vuelta (free move). (b) Si no, con probabilidad p=0,5 da vuelta una cualquiera. (c) Con 1−p da vuelta la de menor break.",
  "C7. UNKNOWN: «no encontré solución en mis pasos»; lo dice WalkSAT. UNSAT: «demostré que no hay»; solo lo puede decir un algoritmo completo (la fuerza bruta).",
  "C8. Es (cláusulas insatisfechas de la mejor asignación) / m. En las insatisfacibles construidas el mínimo posible es 1 insatisfecha (por el núcleo de 8 cláusulas), y WalkSAT lo alcanzó: 1/m, por ejemplo 1/43 = 2,3 % en n=10.",
  "C9. Cada asignación de las 3 variables viola exactamente una de las 8 cláusulas (la de signos opuestos). Entonces toda asignación viola al menos una, sin importar el resto.",
  "C10. La fuerza bruta se detiene en la primera solución del orden 0, 1, 2…, y en n=30 esa solución cayó en el 0,6 % del espacio (6.300.215 asignaciones) mientras que en n=28 cayó en el 9,5 % (25.495.706).",
);
P("**Facundo**", { after: 40, before: 80 });
BS(
  "F1. Las 6 corridas (3 planted + 3 aleatorias) terminaron en UNKNOWN por el tope de 5.000.000 de pasos (≈ 27 s). Dejaron hasta 0,06 % de las cláusulas sin cumplir. NO se puede decir que sean insatisfacibles ni que WalkSAT no pueda.",
  "F2. Planted 3/3, aleatoria 1/3 (4/6 en total). Válido: a n=5.000 las planted fueron más fáciles. No válido: estimar una probabilidad de éxito con 3 muestras, o generalizar.",
  "F3. Porque el tiempo está acotado por el tope de pasos (5.000.000 × ≈ 5 µs ≈ 26 s), no porque WalkSAT se estabilice.",
  "F4. Programación lineal entera (ILP); 35 módulos (5 días × 7 bandas). Se presenta como el caso argentino de horarios mejor documentado, aclarando que no es SAT puro; sirve para mostrar el problema y su pariente NP-completo; el puente a SAT es el coloreo.",
  "F5. Variables x(v,c): 3 vértices × 3 colores = 9. Cláusulas: 3 (al menos un color) + 3·C(3,2) = 9 (a lo sumo uno) + 3 aristas × 3 colores = 9 (vecinos distintos) = 21.",
  "F6. 3 («al menos una franja» por comisión) + 9 (3 pares de franjas por comisión × 3 comisiones, «a lo sumo una») + 6 (2 pares de comisiones en conflicto × 3 franjas) + 1 (cláusula unitaria: el docente de C3 no puede en la franja 1) = 19.",
  "F7. DPLL con aprendizaje: ante un conflicto, se analiza su causa y se aprende una cláusula nueva que evita repetirlo; se vuelve atrás varios niveles. Es completo porque sigue siendo una búsqueda sistemática del espacio, que termina probando SAT o UNSAT.",
  "F8. SAT solo expresa restricciones duras (todo o nada). «Evitar huecos» es un objetivo a optimizar: MaxSAT (cláusulas blandas con peso) o ILP (función objetivo).",
  "F9. 50 M€ por año del horario ferroviario neerlandés (Kroon, Huisman et al., Interfaces, 2009). Se rotula como analogía de investigación operativa, no SAT, y [VERIFICAR] contra el artículo.",
  "F10. Limitaciones: una corrida del exacto por tamaño; planted fáciles; Python; p=0,5 sin ajustar; 3 fórmulas por punto en escala. Mejoras: repetir y promediar, comparar con MiniSat, benchmarks estándar, barrer p. Sobre WalkSAT afectan más el p sin ajustar y las 3 fórmulas por punto (más el tope de pasos).",
);

H2("7.3 Checklist del día de la exposición");
H3("Una semana antes");
BS(
  "Completar la fecha en la portada del informe y de la presentación («[completar fecha de exposición]») y regenerar con §node generadores/informe.js§ y §node generadores/presentacion.js§.",
  "Abrir y leer las referencias de la lista 7.5. Tachar las que verificaron; las que no, no se dicen.",
  "Resolver el Anexo (inconsistencias): decidir qué corregir y qué solo hay que saber.",
  "Decidir cómo van a contestar la pregunta sobre autoría y uso de IA (pregunta 30 del banco) y, si hace falta, consultar a la cátedra.",
  "Ensayo completo con cronómetro: 10 minutos cada uno. Grabarse y escucharse.",
  "Pase de preguntas cruzadas: cada uno le hace 10 preguntas al otro del banco.",
);
H3("La noche anterior");
BS(
  "Repasar los números de memoria de su mitad (los recuadros azules «Números para saber de memoria»).",
  "Repasar el ejemplo a mano de WalkSAT (2.8) y el de la fuerza bruta (2.6).",
  "Dormir. En serio.",
);
H3("El día, antes de entrar");
BS(
  "Notebook cargada y cargador; presentación abierta en el notebook **y** copia en pendrive o nube (PPTX y PDF).",
  "Control remoto o forma de pasar las diapositivas; verificar proyector y resolución.",
  "Tener abiertos a un costado: el informe, esta guía, la carpeta graficos/ y una consola con §python sat.py§ por si piden mostrar el ejemplo.",
  "Acordar quién contesta qué (según la diapositiva) y que el otro complementa solo si hace falta.",
);
H3("Durante");
BS(
  "Camilo termina en el minuto 10:00 aproximadamente. Si va largo, acorta la diapositiva 8 (no la 5).",
  "Hablar de las diapositivas, no leerlas. Mirar al tribunal.",
  "Cuando no sepan algo: «Eso no lo medimos en nuestro trabajo. Lo que sí podemos decir es…». **No inventen datos.**",
  "Para ganar tiempo: «Es una buena pregunta. Lo que sabemos con certeza es esto…».",
  "Dejar la diapositiva 16 en pantalla durante las preguntas.",
);

H2("7.4 Qué NO decir (afirmaciones no verificadas)");
CAJA(
  "OJO",
  "Estas frases suenan bien pero **no están respaldadas** por los datos ni por fuentes verificadas del TP. Decirlas es el camino más corto para que una pregunta de seguimiento los deje sin respuesta.",
);
TABLA(
  ["No digan", "Por qué", "Qué decir en cambio"],
  [
    [
      "«Intel usa SAT para verificar sus chips» (o cualquier empresa concreta).",
      "No se verificó. El informe dice explícitamente que no afirma que empresas específicas usen SAT.",
      "«El bounded model checking, que usa solvers SAT, es una técnica documentada de verificación de hardware y software (Biere et al., 1999).»",
    ],
    [
      "«Microsoft usa SLAM/Z3 para verificar drivers con SAT.»",
      "No está en el informe ni se verificó.",
      "Hablar de Z3 solo como «solver SMT» (de Moura y Bjørner, 2008).",
    ],
    [
      "«Debian / Cargo (PubGrub) resuelven dependencias con CDCL.»",
      "No verificado. PubGrub es un resolvedor propio inspirado en ideas de CDCL, y no figura en el informe.",
      "Limitarse a lo documentado: Eclipse con Sat4j y libsolv (openSUSE) [VERIFICAR].",
    ],
    [
      "«La FRCU/UTN ahorra / ahorraría X horas o X pesos con SAT.»",
      "No hay cifra, ni datos, ni evidencia de uso.",
      "«No hay cifra pública atribuible a SAT; proponemos indicadores para medirlo.»",
    ],
    [
      "«Los 50 M€ son el ahorro de usar SAT.»",
      "Son del horario ferroviario neerlandés, con investigación operativa, tomada de un resumen.",
      "«Como analogía, que no es SAT…»",
    ],
    [
      "«La UTDT / la UNR usan SAT.»",
      "Usan ILP.",
      "«Usan ILP; el puente a SAT es nuestro razonamiento (coloreo → CNF).»",
    ],
    [
      "«Verificamos que el umbral es 4,26.»",
      "Es un resultado de la literatura; el TP midió n=16.",
      "«Según la bibliografía el umbral está cerca de 4,26; nosotros medimos…»",
    ],
    [
      "«n=32 tarda 2 horas.»",
      "No terminó; hay una proyección.",
      "«No terminó en una hora; proyectamos unas 2 horas, pero no lo medimos.»",
    ],
    [
      "«WalkSAT es 400.000 veces más rápido.»",
      "Es una sola fórmula planted fácil, de n=32.",
      "«En esa fórmula, 1,4 ms contra 631 s.»",
    ],
    [
      "«WalkSAT no puede con 10.000 variables» / «es insatisfacible».",
      "Se detuvo por el tope de pasos; no está probado.",
      "«No llegó dentro de nuestro tope de 5.000.000 de pasos.»",
    ],
    [
      "«MiniSat/Kissat lo resolvería en milisegundos.»",
      "No se midió ningún CDCL.",
      "«Esperamos que mucho más rápido, pero no lo medimos.»",
    ],
    [
      "«P ≠ NP está demostrado» o «SAT necesita tiempo exponencial».",
      "Es un problema abierto.",
      "«Nadie conoce un algoritmo polinomial para SAT; saber si existe es P vs NP.»",
    ],
    [
      "«Karp demostró que SAT es NP-completo.»",
      "Fue Cook (1971) y Levin (1973). Karp (1972) usó SAT para demostrar otros 21.",
      "«Cook-Levin.»",
    ],
    [
      "«Kissat ganó todas las categorías de la SAT Competition 2024» a ciegas.",
      "El informe lo dice, pero se tomó de una fuente a confirmar.",
      "Verificar en la página de la competencia, o decir «Kissat figura entre los ganadores [si lo verificaron]».",
    ],
    [
      "«Programamos todo nosotros línea por línea.»",
      "No es lo que pasó. Se generó con una IA.",
      "Ver pregunta 30 del banco. Decir la verdad y poder defender cada parte.",
    ],
  ],
  [3000, 3000, 3026],
  { size: 17 },
);

H2("7.5 Referencias a abrir y verificar antes de exponer");
P(
  "Marquen cada una cuando la hayan abierto. Si una no se puede abrir o dice otra cosa, **no se dice** en la exposición.",
);
TABLA(
  ["", "Referencia", "Qué comprobar"],
  [
    [
      "☐",
      "Ferreiro y Marenco. Un enfoque de programación entera para la programación de horarios de clases en la UTDT. 53 JAIIO (SIIIO), ISSN 2451-7496, pp. 375-382. https://revistas.unlp.edu.ar/JAIIO/article/download/18019/17687/81231",
      "Autores y año (el informe dice 2024 «según metadatos del PDF»); que sea ILP; la grilla de 35 módulos (5 × 7); las dos fases; que el impacto sea cualitativo; si menciona solver, tamaños o tiempos (el informe dice que no los verificó).",
    ],
    [
      "☐",
      "Escalante, Chiesa y Fekete. Problema de asignación de aulas en institución con dos sedes. SUMA 2019, FCEIA-UNR. https://www.union-matematica.org.ar/suma2019/files/10-8-025.pdf",
      "Coloreo de grafos por listas con ILP; dos sedes; datos de 2018; que **no** usan SAT.",
    ],
    [
      "☐",
      "Kroon, Huisman et al. (2009). The new Dutch timetable: the OR revolution. Interfaces 39(1):6-17. https://ideas.repec.org/a/inm/orinte/v39y2009i1p6-17.html",
      "La cifra de 50 M€: valor, año de referencia y de qué es (beneficio anual). Si no coincide, corregir informe, diapositiva 12 y guion.",
    ],
    [
      "☐",
      "Asín Achá y Nieuwenhuis (2014). Curriculum-based course timetabling with SAT and MaxSAT. Ann. Oper. Res. 218(1):71-91. DOI 10.1007/s10479-012-1081-x.",
      "Que sea horarios universitarios con SAT/MaxSAT y el año de la edición impresa (el DOI sugiere publicación en línea en 2012).",
    ],
    [
      "☐",
      "SAT Competition 2024, medallas. https://cca.informatik.uni-freiburg.de/sat24medals",
      "Que Kissat (o una variante) haya ganado, y en cuáles categorías; el informe dice «las tres categorías de la pista principal».",
    ],
    [
      "☐",
      "Biere, Cimatti, Clarke y Zhu (1999). Symbolic model checking without BDDs. TACAS 1999; ETAPS Test of Time Award 2017.",
      "Que sea el origen del BMC con SAT y el premio de 2017.",
    ],
    [
      "☐",
      "Le Berre y Parrain (2010). The Sat4j library, release 2.2. JSAT 7; libsolv (openSUSE).",
      "Que Eclipse use Sat4j para resolver dependencias y que libsolv use un solver SAT.",
    ],
    [
      "☐",
      "Cook (1971) STOC 151-158; Levin (1973); Karp (1972); Selman, Kautz y Cohen (1994) AAAI-94 337-343; Marques-Silva y Sakallah (1999) IEEE TC 48(5):506-521; Eén y Sörensson (2003).",
      "Años, congresos y páginas. Son clásicas, pero conviene confirmar cada dato citado.",
    ],
    [
      "☐",
      "Mitchell, Selman y Levesque (1992); Crawford y Auton (1993); Kirkpatrick y Selman (1994); Friedgut (1999) (umbral de fase).",
      "**No están en el informe.** Solo si quieren citar el umbral con fuente: buscar y verificar antes de nombrarlos.",
    ],
    [
      "☐",
      "Håstad (2001), Schöning (1999), PPSZ, Sinz (2005) (contador secuencial).",
      "**No están en el informe.** Se mencionan en esta guía como contexto [LITERATURA]; no los citen sin verificar.",
    ],
  ],
  [450, 4800, 3776],
  { size: 17 },
);

// =====================================================================
// ANEXO
// =====================================================================
H1("Anexo · Inconsistencias a revisar");
P(
  "Lista de cosas en las que el código contradice, o matiza, lo que dicen los documentos, o que están pendientes. Ordenadas por importancia para la defensa. Ninguna cambia los números de resultados; sí cambian lo que conviene decir o corregir.",
);
TABLA(
  ["#", "Dónde", "Qué pasa", "Qué hacer"],
  [
    [
      "1",
      "Informe y presentación: portada",
      "La fecha dice «[completar fecha de exposición]». (Integrantes y docente ya están completos.)",
      "Completar en §generadores/informe.js§ y §presentacion.js§ y regenerar.",
    ],
    [
      "2",
      "Informe sec. 5, sec. 8.3, sec. 11.2 vs. §sat.py§ → §walksat§",
      "El informe dice que WalkSAT «verifica la asignación antes de responder SAT». El código declara SAT cuando la lista de insatisfechas queda vacía (se mantiene de forma incremental); no vuelve a recorrer las cláusulas ni llama a §contar_insatisfechas§. Es correcto si los contadores son correctos; lo verifiqué sobre 120 modelos.",
      "**Corregido.** Informe y guion ahora dicen que WalkSAT declara SAT cuando no quedan cláusulas insatisfechas según sus contadores internos; el control con §contar_insatisfechas§ sobre los 120 modelos se reproduce con §verificar_modelos.py§ (120/120 en 0).",
    ],
    [
      "3",
      "§sat.py§: docstring de §walksat§ vs. código",
      "El docstring dice: con prob. p una variable cualquiera; si no, la que rompe menos (si alguna rompe 0, esa). El código mira **primero** si hay free move (break 0) y recién después aplica p (versión estándar SKC). La diapositiva 6 también simplifica (omite el free move). El informe sec. 5 lo describe bien.",
      "Saber la versión del código. En la oral describir el algoritmo con el free move primero. Opcional: corregir docstring.",
    ],
    [
      "4",
      "§experimentos.py§ → §guardar§ / §resultados_exacto.csv§",
      "La primera versión escribía el CSV sin §encoding§: en Windows quedó en cp1252 y «sí» se guardó como un byte no UTF-8 (se veía «s�»).",
      "**Corregido.** §experimentos.py§ y §graficos.py§ abren y escriben con §encoding='utf-8'§ (§newline=''§ en CSV); §graficos.py§, §datos.js§ e §informe.js§ leen utf-8 y caen a cp1252/latin1 si falla. §resultados_exacto.csv§ se convirtió de cp1252 a UTF-8 sin tocar ningún valor (ahora «sí» se lee bien en cualquier editor).",
    ],
    [
      "5",
      "Informe sec. 10.3 y diapositiva 12",
      "El ejemplo de 9 variables y 19 cláusulas (139 asignaciones del exacto; 4 pasos de WalkSAT con semilla 1) se reproduce con §ejemplo_horarios.py§ (codificación x(c,f) = 3(c−1)+f) y da exactamente 19 cláusulas, 139 asignaciones y 4 pasos, con las soluciones C1-2, C2-1, C3-2 (exacto) y C1-1, C2-2, C3-3 (WalkSAT).",
      "Corregido. Se agregó §ejemplo_horarios.py§ (python ejemplo_horarios.py) y el informe lo cita.",
    ],
    [
      "6",
      "Guion Q&A («escritos por nosotros en Python»), informe sec. 1 («Se programó…»)",
      "Atribuyen la autoría del código a ustedes. Si el código lo generó una IA, la frase no es literal.",
      "**Corregido.** Informe, guion, guía y README declaran que código y documentos se desarrollaron con asistencia de IA (Claude) y fueron revisados y ejecutados por los autores; informe sec. 7.4. Falta decidir cómo responder la pregunta 30 del banco.",
    ],
    [
      "7",
      "Informe sec. 7.1 vs. §hardware.txt§",
      "El informe dice que el exacto **y** la escalabilidad se corrieron en el mismo equipo, «registrados en hardware.txt». El archivo lo escribe solo el comando §exacto§ (y se pisa en cada tanda). La escala y la densidad no registran hardware. Los tiempos de los archivos son consistentes con una misma sesión (exacto 19:45-21:37, escala 21:42), pero no hay registro.",
      "Decir «la escala se corrió después, en el mismo equipo» solo si están seguros; o agregar §escribir_hardware()§ a §escala()§.",
    ],
    [
      "8",
      "§experimentos.py§: §LIMITE§",
      "Por defecto vale 300 s, no 3.600 s. El README lo aclara con §LIMITE=3600§, pero si alguien corre sin la variable, n ≥ 28 darán TIMEOUT.",
      "Saberlo. Opcional: cambiar el valor por defecto.",
    ],
    [
      "9",
      "§sat.py§: mejor asignación en §walksat§",
      "La mejor asignación se anota al principio de cada iteración, antes de dar vuelta: el último estado (tras el último paso) no se compara. Solo podría afectar el «mejor» en un UNKNOWN. En las insatisfacibles de la comparación el mejor es siempre 1 (el óptimo), así que no cambia nada.",
      "Saberlo; corregir si se toca el código.",
    ],
    [
      "10",
      "Diapositivas 4 y 6 vs. experimento de densidad",
      "Dicen «100.000 pasos (n ≤ 32)». El experimento de densidad (n=16) usó 20.000 pasos (el informe sec. 7.3 lo dice, las diapositivas no). El historial usó 200.000.",
      "Si preguntan, aclarar que los 100.000 son de la comparación exacto vs. WalkSAT.",
    ],
    [
      "11",
      "§resultados_densidad.csv§",
      "Se dijo que no se regeneró. Lo volví a correr con el código actual (≈ 20 s): da exactamente el mismo CSV. Es reproducible.",
      "Nada. Puede decirse con tranquilidad.",
    ],
    [
      "12",
      "§experimentos.py§: comentario de §RAZON§",
      "Dice «la zona más difícil del 3-SAT aleatorio», pero se aplica a fórmulas planted e insatisfacibles construidas, no aleatorias puras.",
      "No decir que las fórmulas del exacto están «en la zona más difícil» sin matizar.",
    ],
    [
      "13",
      "§contar_insatisfechas§",
      "Existe pero ningún experimento lo usa.",
      "Saberlo; es el verificador independiente que se puede mostrar.",
    ],
    [
      "14",
      "Informe sec. 10.1: «SAT, ILP y coloreo… son NP-completos»",
      "El coloreo con k ≥ 3 colores es NP-completo; con k = 2 es polinomial (bipartito).",
      "Precisión menor; decir «k-coloreo con k ≥ 3».",
    ],
    [
      "15",
      "Diapositivas 2 y 16",
      "La 2 nombra a Levin y la 16 lista solo a Cook (Levin y Karp están en el informe).",
      "Si preguntan por Levin: Levin (1973), «Universal sequential search problems».",
    ],
  ],
  [450, 2000, 3900, 2676],
  { size: 16 },
);

// =====================================================================
// ARMADO
// =====================================================================
const doc = new Document({
  creator: "TP SAT",
  title: "Guía de estudio · SAT (Problema de Satisfacción Booleana): exacto frente a la heurística WalkSAT",
  description: "Material de estudio para la defensa oral del TP de SAT",
  styles: { default: { document: { run: { font: "Arial", size: 22 } } } },
  numbering: {
    config: [
      {
        reference: "vi",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 500, hanging: 260 } } },
          },
          {
            level: 1,
            format: LevelFormat.BULLET,
            text: "–",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 900, hanging: 260 } } },
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
          margin: { top: 1300, bottom: 1200, left: 1440, right: 1440 },
        },
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: "Guía de estudio · TP SAT · Camilo Escar y Facundo Iglesias · pág. ",
                  font: "Arial",
                  size: 16,
                  color: GRIS,
                }),
                new TextRun({
                  children: [PageNumber.CURRENT],
                  font: "Arial",
                  size: 16,
                  color: GRIS,
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
  const out = path.join(D.RAIZ, "GUIA_DE_ESTUDIO.docx");
  fs.writeFileSync(out, buf);
  console.log("OK", out, Math.round(buf.length / 1024) + " KB");
});
