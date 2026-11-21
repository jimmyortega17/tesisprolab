const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/root/.claude/skills/synced/8eb71c26-ef7b-4c54-a359-b2c4555876f2_b18771c6-8bcb-42a3-98aa-2dff77f105b5/pptx/scripts/apply_theme.js");

const THEME = {
  name: "Cosecha Viva",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1E1A16", lt1: "FFFFFF", dk2: "2F4A2B", lt2: "F1F5EC",
    accent1: "3F7D3A", accent2: "C8762B", accent3: "6B4226", accent4: "B8332B",
    accent5: "2D6A8F", accent6: "E3B23C", hlink: "2D6A8F", folHlink: "6B4226",
  },
};
const HEX = { dk1: "1E1A16", dk2: "2F4A2B", lt2: "F1F5EC", g: "3F7D3A", am: "C8762B", br: "6B4226", red: "B8332B", bl: "2D6A8F", gold: "E3B23C", gray: "5C5750" };

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Cosecha Viva: problema, solución y rentabilidad";
pres.author = "Laura Pampa, Quinto Huatuco, Aguilar Lazo, Estrada Cruz";
const C = pres.SchemeColor;
const S = pres.shapes;
const IMG = "img/";

// ---------- iconos ----------
async function icon(name, hex, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + hex, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
const iconCache = {};
async function ic(name, hex) { const k = name + hex; if (!iconCache[k]) iconCache[k] = await icon(name, hex); return iconCache[k]; }

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "CONTENT", background: { color: "FFFFFF" },
  objects: [{ text: { text: "Cosecha Viva | PUCP | Escuela de Posgrado", options: { x: 0.5, y: 5.28, w: 5, h: 0.25, fontSize: 10, color: "5C5750", margin: 0, isTextBox: true } } }],
  slideNumber: { x: 9.0, y: 5.28, w: 0.5, h: 0.25, fontSize: 10, color: "5C5750", align: "right" },
});
pres.defineSlideMaster({ title: "DARK", background: { color: "1E1A16" } });
pres.defineSlideMaster({
  title: "SOFT", background: { color: "F1F5EC" },
  slideNumber: { x: 9.0, y: 5.28, w: 0.5, h: 0.25, fontSize: 10, color: "5C5750", align: "right" },
});

let notesOn = true;
function content(title, opts = {}) {
  const s = pres.addSlide({ masterName: opts.soft ? "SOFT" : "CONTENT" });
  s.addText(title, { x: 0.5, y: 0.3, w: 9.0, h: 0.7, fontSize: 30, bold: true, fontFace: "Cambria", color: C.text2, margin: 0, valign: "middle", isTextBox: true, objectName: "Titulo" });
  return s;
}
function source(s, text) { s.addText(text, { x: 0.5, y: 4.98, w: 9.0, h: 0.28, fontSize: 10, italic: true, color: C.text1, transparency: 25, margin: 0, valign: "middle", isTextBox: true, objectName: "Fuente" }); }

function card(s, x, y, w, h, fill, line) {
  s.addShape(S.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: line ? { color: line, width: 1 } : { color: fill, width: 0 }, shadow: undefined });
}
async function badge(s, x, y, d, fill, iconName, iconHex = "FFFFFF") {
  s.addShape(S.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 } });
  const pad = d * 0.25;
  s.addImage({ data: await ic(iconName, iconHex), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad });
}
function T(s, text, o) { s.addText(text, Object.assign({ margin: 0, isTextBox: true, fontFace: "Calibri", color: C.text1 }, o)); }

(async () => {
  async function imageSlide(file, notes) {
    const sl = pres.addSlide({ masterName: "DARK" });
    sl.background = { color: "FFFFFF" };
    sl.addImage({ path: IMG + file, x: 0, y: 0, w: 10, h: 5.625 });
    sl.addNotes(notes);
  }
  // =============== 1. PORTADA ===============
  {
    const s = pres.addSlide({ masterName: "DARK" });
    s.addImage({ path: IMG + "ef033cf01018c12822652c4c4db8dfa64114d72c.jpg", x: 5.6, y: 0, w: 4.4, h: 5.625, sizing: { type: "cover", w: 4.4, h: 5.625 } });
    T(s, "PONTIFICIA UNIVERSIDAD CATÓLICA DEL PERÚ | ESCUELA DE POSGRADO", { x: 0.5, y: 0.45, w: 4.9, h: 0.5, fontSize: 11, color: "E3B23C", bold: true });
    T(s, "Cosecha Viva", { x: 0.5, y: 1.15, w: 4.9, h: 0.9, fontSize: 44, bold: true, fontFace: "Cambria", color: "FFFFFF" });
    T(s, "Preventa de café tecnificado con trazabilidad para elevar los ingresos de los pequeños productores de Chanchamayo", { x: 0.5, y: 2.1, w: 4.9, h: 1.3, fontSize: 18, color: "F1F5EC" });
    T(s, "Problema socialmente relevante, solución y rentabilidad", { x: 0.5, y: 3.55, w: 4.9, h: 0.4, fontSize: 14, italic: true, color: "E3B23C" });
    T(s, "Laura Pampa, D. W. · Quinto Huatuco, H. H.\nAguilar Lazo, R. H. · Estrada Cruz, F.\nMaestría en Administración Estratégica de Empresas · Lima, noviembre de 2026", { x: 0.5, y: 4.25, w: 4.9, h: 0.95, fontSize: 11, color: "F1F5EC" });
    s.addNotes("Presentación de la tesis Cosecha Viva. La foto corresponde al trabajo de campo del equipo en la Selva Central. Guía: problema, lienzos de investigación del usuario, solución en forma de historia, flujo de caja y ODS.");
  }

  // =============== 2. IMPACTO DEL PSR ===============
  {
    const s = content("Un récord que no llega a la chacra");
    const stats = [
      ["US$ 1,796 M", "en exportaciones de café en 2025, un récord histórico", "(El Peruano, 2026)", "globe"],
      ["15 qq/ha", "de rendimiento promedio; con nutrición balanceada se superan los 50", "(Agraria.pe, 2019b)", "seed"],
      ["S/ 15 a 20", "por kg de pergamino al productor; el kg tostado de especialidad se ofrece en Lima a ~S/ 130", "(Gestión, 2025; Mercado Libre Perú, 2026)", "money"],
      ["340,000", "familias rurales sostenidas por el café y el cacao", "(Agraria.pe, 2026)", "users"],
    ];
    const icons = { globe: "FaGlobeAmericas", seed: "FaSeedling", money: "FaMoneyBillWave", users: "FaUsers" };
    const colors = [HEX.bl, HEX.g, HEX.am, HEX.br];
    for (let i = 0; i < 4; i++) {
      const x = 0.5 + i * 2.3;
      card(s, x, 1.25, 2.1, 3.0, HEX.lt2);
      await badge(s, x + 0.15, 1.4, 0.55, colors[i], icons[stats[i][3]]);
      T(s, stats[i][0], { x: x + 0.15, y: 2.05, w: 1.85, h: 0.6, fontSize: 22, bold: true, fontFace: "Cambria", color: colors[i], valign: "middle" });
      T(s, stats[i][1], { x: x + 0.15, y: 2.7, w: 1.85, h: 1.1, fontSize: 13, valign: "top" });
      T(s, stats[i][2], { x: x + 0.15, y: 3.8, w: 1.85, h: 0.4, fontSize: 10, italic: true, color: C.text1, valign: "bottom" });
    }
    s.addShape(S.ROUNDED_RECTANGLE, { x: 0.5, y: 4.4, w: 9.0, h: 0.5, rectRadius: 0.08, fill: { color: HEX.dk2 }, line: { color: HEX.dk2, width: 0 } });
    T(s, "El valor se queda en el acopio, la exportación y el tostado, no en la chacra (El Peruano, 2026).", { x: 0.7, y: 4.4, w: 8.6, h: 0.5, fontSize: 13, bold: true, color: "FFFFFF", valign: "middle" });
    source(s, "Fuentes: ver referencias al final de la presentación.");
    s.addNotes("El café peruano vive un récord de exportaciones, pero el pequeño productor de Chanchamayo, con parcelas de unas dos hectáreas, vende pergamino sin diferenciar al intermediario. La brecha de rendimiento (15 frente a más de 50 quintales por hectárea) y la brecha de precio (S/ 15 a 20 frente a unos S/ 130 por kilogramo tostado) muestran dónde se pierde el valor. Aclaración: el precio del productor es por kilogramo de pergamino y el de Lima por kilogramo de café tostado; no son el mismo producto.");
  }

  // =============== 3. EL PSR: CICLO ===============
  {
    const s = content("El problema: ingresos bajos e inestables");
    // ciclo
    const nodes = [
      { x: 1.9, y: 1.2, t: "Baja productividad", d: "~15 qq/ha, sin plan de nutrición ni fertirriego", ic: "FaSeedling", c: HEX.g },
      { x: 3.65, y: 2.5, t: "Venta sin diferenciar", d: "El intermediario fija un precio bajo y volátil", ic: "FaStore", c: HEX.am },
      { x: 1.9, y: 3.8, t: "Sin capital ni comprador", d: "No hay inversión en tecnificación", ic: "FaMoneyBillWave", c: HEX.br },
      { x: 0.15, y: 2.5, t: "Calidad irregular", d: "Mezcla de calidades; menor precio", ic: "FaCoffee", c: HEX.red },
    ];
    for (const n of nodes) {
      card(s, n.x, n.y, 1.75, 1.05, "FFFFFF", n.c);
      await badge(s, n.x + 0.08, n.y + 0.1, 0.4, n.c, n.ic);
      T(s, n.t, { x: n.x + 0.55, y: n.y + 0.08, w: 1.15, h: 0.45, fontSize: 11, bold: true, valign: "middle" });
      T(s, n.d, { x: n.x + 0.1, y: n.y + 0.55, w: 1.6, h: 0.48, fontSize: 10, valign: "top" });
    }
    const arrows = [[3.67, 1.45, 4.35, 2.45], [4.4, 3.55, 3.7, 4.2], [1.9, 4.3, 1.2, 3.55], [0.95, 2.45, 1.8, 1.7]];
    // flechas simples como conectores
    s.addShape(S.LINE, { x: 3.65, y: 1.55, w: 0.75, h: 0.95, line: { color: HEX.gray, width: 2, endArrowType: "triangle" } });
    s.addShape(S.LINE, { x: 3.7, y: 3.55, w: 0.7, h: 1.05, line: { color: HEX.gray, width: 2, endArrowType: "triangle" }, flipH: true });
    s.addShape(S.LINE, { x: 1.0, y: 3.55, w: 0.9, h: 1.0, line: { color: HEX.gray, width: 2, endArrowType: "triangle" }, flipH: true, flipV: true });
    s.addShape(S.LINE, { x: 1.0, y: 1.7, w: 0.9, h: 0.8, line: { color: HEX.gray, width: 2, endArrowType: "triangle" }, flipV: true });
    s.addShape(S.OVAL, { x: 2.35, y: 2.6, w: 0.85, h: 0.85, fill: { color: HEX.dk2 }, line: { color: HEX.dk2, width: 0 } });
    T(s, "PSR", { x: 2.35, y: 2.6, w: 0.85, h: 0.85, fontSize: 16, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    // enunciado
    card(s, 5.85, 1.2, 3.65, 3.65, HEX.lt2);
    T(s, "Problema socialmente relevante", { x: 6.05, y: 1.3, w: 3.3, h: 0.4, fontSize: 16, bold: true, fontFace: "Cambria", color: C.text2 });
    T(s, "Los pequeños productores de café de Chanchamayo obtienen ingresos bajos e inestables: producen por debajo de su potencial, venden sin diferenciar a intermediarios y no llegan al consumidor urbano que paga por calidad y origen.", { x: 6.05, y: 1.75, w: 3.3, h: 1.9, fontSize: 13 });
    await badge(s, 6.05, 3.85, 0.45, HEX.red, "FaFemale");
    T(s, "Consecuencia social: migración de jóvenes y abandono de fincas", { x: 6.6, y: 3.8, w: 2.8, h: 0.6, fontSize: 12, bold: true, valign: "middle" });
    T(s, "Actuar solo sobre el precio o solo sobre la productividad no basta: hay que romper el ciclo en ambos extremos.", { x: 6.05, y: 4.4, w: 3.3, h: 0.4, fontSize: 10, italic: true });
    source(s, "Adaptado de Agraria.pe (2019b, 2026), Gestión (2025) y del trabajo de campo del equipo.");
    s.addNotes("El problema es un círculo: sin un comprador que asegure un mejor precio, el productor no tiene incentivos ni recursos para tecnificar; sin tecnificación no produce la calidad que ese comprador exigiría. Por eso la solución debe actuar sobre los dos extremos a la vez y con recursos de mercado, no con programas públicos.");
  }

  // =============== 4. LIENZOS BLANCO DE RELEVANCIA Y QUÉ ES / QUÉ NO ES ===============
  await imageSlide("lz2-2.png", "Lienzo ¿Qué es? y ¿Qué no es? del problema: es baja productividad, venta sin diferenciar, falta de capital y de comprador seguro, y migración de jóvenes; no es solo un problema de precio ni de productividad, ni lo resuelve el turismo cafetalero ni depende del Estado.");

  // =============== 5. LIENZO META-USUARIO ===============
  await imageSlide("lz4-4.png", "Matriz de meta-usuario de Rosa, una productora de café de 42 años de Chanchamayo con dos hectáreas. Sus dolores son el bajo rendimiento, el precio que fija el intermediario y la falta de capital; sus ganancias esperadas son más producción, mejor precio y un comprador seguro.");

  // =============== 6 y 7. MAPAS DE EMPATÍA ===============
  async function empathy(title, persona, sub, boxes, srcTxt, notes, color, iconName) {
    const s = content(title);
    card(s, 0.5, 1.15, 9.0, 0.65, HEX.lt2);
    await badge(s, 0.62, 1.2, 0.55, color, iconName);
    T(s, persona, { x: 1.3, y: 1.15, w: 3.2, h: 0.65, fontSize: 16, bold: true, fontFace: "Cambria", color: C.text2, valign: "middle" });
    T(s, sub, { x: 4.4, y: 1.15, w: 5.0, h: 0.65, fontSize: 12, italic: true, valign: "middle" });
    const bw = 2.9, bh = 1.45;
    for (let i = 0; i < 6; i++) {
      const col = i % 3, row = Math.floor(i / 3);
      const x = 0.5 + col * (bw + 0.15), y = 1.95 + row * (bh + 0.1);
      const kind = boxes[i][2];
      const fill = kind === "pain" ? "FBE9E7" : kind === "gain" ? "E5F1E2" : "FFFFFF";
      const line = kind === "pain" ? HEX.red : kind === "gain" ? HEX.g : "C9D4C0";
      s.addShape(S.ROUNDED_RECTANGLE, { x, y, w: bw, h: bh, rectRadius: 0.08, fill: { color: fill }, line: { color: line, width: 1.25 } });
      T(s, boxes[i][0], { x: x + 0.15, y: y + 0.08, w: bw - 0.3, h: 0.3, fontSize: 13, bold: true, color: kind === "pain" ? HEX.red : kind === "gain" ? HEX.g : C.text3 });
      T(s, boxes[i][1], { x: x + 0.15, y: y + 0.4, w: bw - 0.3, h: bh - 0.45, fontSize: 11.5, valign: "top" });
    }
    source(s, srcTxt);
    s.addNotes(notes);
  }
  await empathy("Mapa de empatía: el productor", "Rosa, 42 años", "Dos hectáreas de café en Chanchamayo; cosecha ~30 quintales por campaña",
    [["Ve", "Su parcela con poca nutrición, plagas, café mezclado al cosechar y precios que cambian sin explicación.", ""],
     ["Escucha", "Del acopiador, el precio del día; de otros productores, quién paga más; de los técnicos, prácticas que no puede costear.", ""],
     ["Piensa y siente", "Preocupación por pagar las deudas de la campaña; desconfianza hacia quien promete mejores precios; orgullo por su café.", ""],
     ["Dice y hace", "“El intermediario pone el precio”; “no tengo plata para abonar”. Vende al acopiador más cercano; acepta visitas si traen ingreso (85%).", ""],
     ["Dolores", "Falta de capital para abonar; roya y broca; mezcla de calidades; poca capacidad de negociar el precio.", "pain"],
     ["Ganancias", "Producir más y mejor, vender a un comprador seguro con prima y recibir apoyo técnico sin endeudarse.", "gain"]],
    "Adaptado del trabajo de campo del equipo (n = 20 productores). Persona ilustrativa de la narrativa de la tesis.",
    "Rosa es la persona que usamos para contar la historia. Cada recuadro sale del trabajo de campo con 20 productores: 85% acepta visitantes y 80% espera un mayor valor para su café.", HEX.g, "FaFemale");
  await empathy("Mapa de empatía: el consumidor", "Carla, 34 años", "Profesional de Lima que compra café en grano cada mes",
    [["Ve", "Marcas con historias de origen en redes, cafeterías de especialidad y bolsas con datos de finca y variedad.", ""],
     ["Escucha", "Recomendaciones de baristas, amigos y comunidades de café; críticas al café viejo o mezclado.", ""],
     ["Piensa y siente", "Quiere un café fresco y auténtico; duda de que el origen sea real; le gusta ayudar si ve el resultado.", ""],
     ["Dice y hace", "“No sé de dónde viene mi café”; “me gustaría conocer la finca”. Compra cada mes, compara precios y regala.", ""],
     ["Dolores", "Desconfianza en el origen, café no siempre fresco, poca conexión con el productor.", "pain"],
     ["Ganancias", "Café fresco y trazable, impacto visible en el productor y una experiencia que puede compartir.", "gain"]],
    "Hipótesis del equipo a validar con la encuesta a 150 compradores (Apéndice B). 92% de los turistas encuestados valora que su pago beneficie al productor.",
    "Este mapa es una hipótesis: la encuesta de campo se hizo a turistas potenciales y no a compradores recurrentes de café de especialidad. Se valida en el piloto con la preventa real.", HEX.br, "FaUserTie");

  // =============== 8. MAPA EXPERIENCIA PRODUCTOR ===============
  await imageSlide("lz3-3.png", "Storyboard del productor en la situación actual, con diez momentos de la campaña, sus pensamientos y su curva emocional. Los puntos más bajos son la cosecha mezclada, el secado con lluvia y el momento posterior a la venta, cuando no queda dinero para abonar.");

  // =============== 9. MAPA EXPERIENCIA CONSUMIDOR ===============
  await imageSlide("lz5-5.png", "Storyboard del consumidor con Cosecha Viva: la curva sube a medida que Carla ve el origen del café, con una duda al decidir la compra por adelantado. Los puntos de dolor son hipótesis que se confirman en el piloto.");

  // =============== 10. EVIDENCIA DE CAMPO ===============
  {
    const s = content("Evidencia de campo: interés, no compra");
    const labels = ["Considera que el pago beneficie al productor (turistas)", "Interesado en una experiencia cafetalera (turistas)", "Pagaría S/ 100 o más (turistas)", "Mejora ingresos de la familia (productores)", "Acepta recibir visitantes (productores)", "Espera mayor valor para su café (productores)"];
    const vals = [92, 88, 76, 90, 85, 80];
    s.addChart(pres.charts.BAR, [{ name: "% de respuestas", labels, values: vals }], {
      x: 0.4, y: 1.15, w: 5.6, h: 3.7, barDir: "bar", chartColors: [HEX.g], barGapWidthPct: 45,
      showValue: true, dataLabelFormatCode: "0\"%\"", dataLabelFontSize: 11, dataLabelColor: "1E1A16", dataLabelPosition: "outEnd",
      catAxisLabelFontSize: 10.5, catAxisLabelColor: "1E1A16", catAxisOrientation: "maxMin", valAxisHidden: true, valAxisMaxVal: 110, valAxisMinVal: 0,
      valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, catAxisLabelFontFace: "Calibri", dataLabelFontFace: "Calibri",
    });
    s.addImage({ path: IMG + "ef033cf01018c12822652c4c4db8dfa64114d72c.jpg", x: 6.2, y: 1.2, w: 3.3, h: 1.86, sizing: { type: "cover", w: 3.3, h: 1.86 } });
    T(s, "Trabajo de campo del equipo en la Selva Central", { x: 6.2, y: 3.08, w: 3.3, h: 0.25, fontSize: 10, italic: true });
    card(s, 6.2, 3.45, 3.3, 1.4, "FFF3E0", HEX.am);
    T(s, "Límite de la evidencia", { x: 6.35, y: 3.5, w: 3.0, h: 0.3, fontSize: 12, bold: true, color: HEX.am });
    T(s, "Muestra por conveniencia (20 productores y 25 turistas potenciales). Valida interés, pero aún no demuestra la compra anticipada del producto: eso lo prueba el piloto.", { x: 6.35, y: 3.82, w: 3.05, h: 1.0, fontSize: 11 });
    source(s, "Elaboración propia a partir del trabajo de campo del equipo (Apéndice A), abril y mayo de 2026.");
    s.addNotes("La evidencia de campo respalda tres hipótesis: H1, el consumidor valora que su pago beneficie al productor (92%); H2, paga S/ 100 o más por una experiencia (76%); y H3, el productor acepta visitas (85%). Son respuestas declaradas, de una muestra pequeña y por conveniencia, de modo que no sustituyen la prueba de compra real.");
  }

  await imageSlide("lz1-1.png", "Lienzo blanco de relevancia con la plantilla de la consultora. Se construye después de probar el prototipo con los usuarios (PMV 1: 20 productores y 25 turistas potenciales): cosas interesantes (88% se interesa, 92% valora que el pago llegue al productor, 85% de productores acepta visitas), críticas constructivas (el turismo ya existe y no sube el rendimiento), nuevas preguntas (las hipótesis H5 a H8) y nuevas ideas (vender el café, preventa, tecnificar, turismo como puerta de entrada).");
  await imageSlide("lz6-6.png", "Lienzo 6x6: seis preguntas generadoras, una por necesidad, con seis ideas de solución cada una y la idea elegida en la fila final.");
  await imageSlide("lz7-7.png", "Matriz costo-impacto de las cuatro alternativas. La alternativa D (tecnificación, compra asegurada, preventa con trazabilidad y tour) es la elegida porque actúa sobre las dos causas del núcleo crítico y se financia con el dinero del cliente.");

  // =============== 11. LA SOLUCIÓN: 4 PALANCAS ===============
  {
    const s = content("La solución: Cosecha Viva");
    T(s, "Una empresa privada que tecnifica al productor, le compra su café selecto con prima y lo vende por adelantado al consumidor, mostrándole el origen.", { x: 0.5, y: 1.1, w: 9.0, h: 0.6, fontSize: 15, color: C.text1 });
    const p = [
      ["Tecnifica", "Kit de fertirriego y secador solar en comodato, más asistencia técnica", "FaSeedling", HEX.g],
      ["Compra con prima", "S/ 20 por kg de pergamino: 33% sobre el precio del intermediario", "FaHandshake", HEX.am],
      ["Preventa", "Paquete de cosecha de 1 kg (4 bolsas de 250 g) a S/ 140", "FaBoxOpen", HEX.br],
      ["Muestra el origen", "Videos por etapa, código QR por lote y visita opcional a la finca", "FaQrcode", HEX.bl],
    ];
    for (let i = 0; i < 4; i++) {
      const x = 0.5 + i * 2.3;
      card(s, x, 1.85, 2.1, 1.95, HEX.lt2);
      await badge(s, x + 0.15, 1.97, 0.6, p[i][3], p[i][2]);
      T(s, p[i][0], { x: x + 0.15, y: 2.65, w: 1.85, h: 0.35, fontSize: 14, bold: true, fontFace: "Cambria", color: C.text2 });
      T(s, p[i][1], { x: x + 0.15, y: 3.0, w: 1.85, h: 0.78, fontSize: 11.5, valign: "top" });
    }
    // flujo
    const fx = [0.5, 3.6, 6.7];
    const ft = [["Productor", "más producción, mejor precio, comprador seguro", HEX.g], ["Cosecha Viva", "acopia, tuesta con aliados y entrega", HEX.dk2], ["Consumidor", "café fresco y trazable; impacto visible", HEX.br]];
    for (let i = 0; i < 3; i++) {
      s.addShape(S.ROUNDED_RECTANGLE, { x: fx[i], y: 4.05, w: 2.8, h: 0.8, rectRadius: 0.08, fill: { color: ft[i][2] }, line: { color: ft[i][2], width: 0 } });
      T(s, ft[i][0], { x: fx[i] + 0.1, y: 4.07, w: 2.6, h: 0.35, fontSize: 14, bold: true, color: "FFFFFF", align: "center", fontFace: "Cambria" });
      T(s, ft[i][1], { x: fx[i] + 0.1, y: 4.4, w: 2.6, h: 0.42, fontSize: 11, color: "FFFFFF", align: "center", valign: "top" });
      if (i < 2) {
        s.addShape(S.RIGHT_ARROW, { x: fx[i] + 2.85, y: 4.28, w: 0.4, h: 0.34, fill: { color: HEX.am }, line: { color: HEX.am, width: 0 } });
      }
    }
    source(s, "Fuente: tesis Cosecha Viva, capítulo IV. Precio del paquete S/ 140 con IGV.");
    s.addNotes("Cuatro palancas actúan a la vez sobre las causas del problema: la tecnificación ataca la baja productividad; la compra con prima ataca el precio bajo; la preventa resuelve el capital y la certeza de compra; y la trazabilidad con videos y QR cierra la brecha de información con el consumidor. El turismo es el segundo ingreso, no el centro del modelo.");
  }

  await imageSlide("lz8-8.png", "Lienzo de la propuesta de valor: el mapa de valor de Cosecha Viva (productos y servicios, generadores de alegrías y mitigadores de dolor) frente al perfil del usuario (alegrías, temas del cliente y dolores).");

  // =============== 12. EL CICLO ===============
  {
    const s = content("Cómo funciona: el ciclo del café", { soft: true });
    const steps = [
      ["Técnico en la finca", "Plan de nutrición y asistencia técnica", "FaUserCheck", HEX.g],
      ["Kit en comodato", "Fertirriego y secador solar", "FaTint", HEX.bl],
      ["Acuerdo de compra", "Precio con prima y piso garantizado", "FaFileSignature", HEX.br],
      ["Preventa en Lima", "El cliente paga por adelantado S/ 140", "FaShoppingCart", HEX.am],
      ["Videos por etapa", "Cosecha, secado y tueste por WhatsApp", "FaVideo", HEX.red],
      ["Cosecha selectiva", "Solo granos maduros; secado solar", "FaSun", HEX.gold],
      ["Tueste y entrega", "Fecha de tueste y QR a la ficha del lote", "FaBoxOpen", HEX.dk2],
      ["Visita y recompra", "Tour a la finca y renovación del paquete", "FaRedo", HEX.g],
    ];
    for (let i = 0; i < 8; i++) {
      const row = Math.floor(i / 4), col = i % 4;
      const x = 0.5 + col * 2.3, y = 1.2 + row * 1.85;
      card(s, x, y, 2.1, 1.65, "FFFFFF", "C9D4C0");
      await badge(s, x + 0.12, y + 0.12, 0.55, steps[i][3], steps[i][2], steps[i][3] === HEX.gold ? "1E1A16" : "FFFFFF");
      T(s, String(i + 1), { x: x + 1.65, y: y + 0.1, w: 0.35, h: 0.4, fontSize: 24, bold: true, fontFace: "Cambria", color: "C9D4C0", align: "right" });
      T(s, steps[i][0], { x: x + 0.12, y: y + 0.75, w: 1.9, h: 0.3, fontSize: 13, bold: true, color: C.text2 });
      T(s, steps[i][1], { x: x + 0.12, y: y + 1.05, w: 1.9, h: 0.55, fontSize: 11.5, valign: "top" });
      if (col < 3) s.addShape(S.RIGHT_ARROW, { x: x + 2.1, y: y + 0.7, w: 0.2, h: 0.25, fill: { color: HEX.am }, line: { color: HEX.am, width: 0 } });
    }
    source(s, "Fuente: tesis Cosecha Viva, capítulos IV y V.");
    s.addNotes("El ciclo tiene ocho pasos. Los cuatro primeros preparan la oferta y la venta; los cuatro últimos cierran la relación con el cliente. El pago anticipado del paso 4 financia la tecnificación y reduce el capital de trabajo.");
  }

  // =============== 13-16. HISTORIA ===============
  function scenePanel(s, x, y, w, h, bg) { s.addShape(S.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.1, fill: { color: bg }, line: { color: bg, width: 0 } }); }
  async function sceneFarm(s, x, y, w, h) {
    scenePanel(s, x, y, w, h, "CFE6F2");
    s.addShape(S.SUN, { x: x + w - 0.95, y: y + 0.15, w: 0.75, h: 0.75, fill: { color: HEX.gold }, line: { color: HEX.gold, width: 0 } });
    s.addShape(S.CLOUD, { x: x + 1.9, y: y + 0.2, w: 0.9, h: 0.5, fill: { color: "FFFFFF" }, line: { color: "FFFFFF", width: 0 } });
    s.addShape(S.ISOSCELES_TRIANGLE, { x: x + 0.1, y: y + 0.95, w: 2.2, h: 1.0, fill: { color: "6FA36A" }, line: { color: "6FA36A", width: 0 } });
    s.addShape(S.ISOSCELES_TRIANGLE, { x: x + 1.6, y: y + 0.75, w: 2.6, h: 1.2, fill: { color: "4E8A4B" }, line: { color: "4E8A4B", width: 0 } });
    s.addShape(S.OVAL, { x: x + 0.1, y: y + h - 1.0, w: w - 0.2, h: 0.85, fill: { color: "8DB87F" }, line: { color: "8DB87F", width: 0 } });
    s.addShape(S.OVAL, { x: x + 0.5, y: y + h - 0.8, w: w - 1.0, h: 0.62, fill: { color: HEX.g }, line: { color: HEX.g, width: 0 } });
    for (let i = 0; i < 6; i++) {
      s.addImage({ data: await ic("FaSeedling", "1F4D1B"), x: x + 0.55 + i * 0.55, y: y + h - 0.95, w: 0.46, h: 0.46 });
      s.addShape(S.OVAL, { x: x + 0.78 + i * 0.55, y: y + h - 0.6, w: 0.1, h: 0.1, fill: { color: HEX.red }, line: { color: HEX.red, width: 0 } });
    }
    const it = [["FaUserCheck", HEX.am, "Técnico"], ["FaTint", HEX.bl, "Fertirriego"], ["FaSun", HEX.gold, "Secador solar"]];
    for (let i = 0; i < 3; i++) {
      const bx = x + 0.2 + i * 1.15;
      await badge(s, bx + 0.12, y + 0.18, 0.6, it[i][1], it[i][0], it[i][1] === HEX.gold ? "1E1A16" : "FFFFFF");
      s.addShape(S.ROUNDED_RECTANGLE, { x: bx - 0.02, y: y + 0.8, w: 1.0, h: 0.26, rectRadius: 0.05, fill: { color: "FFFFFF" }, line: { color: "FFFFFF", width: 0 } });
      T(s, it[i][2], { x: bx - 0.02, y: y + 0.8, w: 1.0, h: 0.26, fontSize: 10, bold: true, align: "center", valign: "middle" });
    }
  }
  async function sceneLima(s, x, y, w, h) {
    scenePanel(s, x, y, w, h, "26324A");
    s.addShape(S.MOON, { x: x + w - 0.7, y: y + 0.15, w: 0.4, h: 0.6, fill: { color: "F6E7CE" }, line: { color: "F6E7CE", width: 0 } });
    const bx = [0.15, 0.7, 1.25, 1.85, 2.45, 3.05, 3.6];
    const bh = [1.0, 1.6, 1.25, 1.85, 1.35, 1.0, 1.5];
    for (let i = 0; i < bx.length; i++) {
      s.addShape(S.RECTANGLE, { x: x + bx[i], y: y + h - bh[i] - 0.1, w: 0.5, h: bh[i], fill: { color: "3A4A6B" }, line: { color: "3A4A6B", width: 0 } });
      for (let k = 0; k < 3; k++) s.addShape(S.RECTANGLE, { x: x + bx[i] + 0.1, y: y + h - bh[i] + 0.05 + k * 0.28, w: 0.12, h: 0.12, fill: { color: i % 2 ? HEX.gold : "7F8FB0" }, line: { color: "3A4A6B", width: 0 } });
    }
    const it = [["FaInstagram", "C13584", "Ve el video"], ["FaCreditCard", HEX.am, "Paga S/ 140"], ["FaWhatsapp", "25D366", "Recibe videos"]];
    for (let i = 0; i < 3; i++) {
      const bx2 = x + 0.25 + i * 1.3;
      await badge(s, bx2 + 0.15, y + 0.2, 0.7, it[i][1], it[i][0]);
      s.addShape(S.ROUNDED_RECTANGLE, { x: bx2, y: y + 0.97, w: 1.0, h: 0.26, rectRadius: 0.05, fill: { color: "FFFFFF" }, line: { color: "FFFFFF", width: 0 } });
      T(s, it[i][2], { x: bx2, y: y + 0.97, w: 1.0, h: 0.26, fontSize: 10, bold: true, align: "center", valign: "middle" });
      if (i < 2) s.addShape(S.RIGHT_ARROW, { x: bx2 + 1.02, y: y + 0.42, w: 0.2, h: 0.24, fill: { color: "F6E7CE" }, line: { color: "F6E7CE", width: 0 } });
    }
  }
  async function sceneHarvest(s, x, y, w, h) {
    scenePanel(s, x, y, w, h, "F6E7CE");
    s.addShape(S.SUN, { x: x + w - 0.8, y: y + 0.12, w: 0.6, h: 0.6, fill: { color: HEX.gold }, line: { color: HEX.gold, width: 0 } });
    s.addShape(S.OVAL, { x: x + 0.1, y: y + h - 0.75, w: w - 0.2, h: 0.6, fill: { color: "C9A878" }, line: { color: "C9A878", width: 0 } });
    const it = [["FaMapMarkedAlt", HEX.g, "Visita"], ["FaCoffee", HEX.am, "Tueste"], ["FaBoxOpen", HEX.br, "4 bolsas"], ["FaQrcode", HEX.dk2, "QR del lote"]];
    for (let i = 0; i < 4; i++) {
      const cx = x + 0.2 + i * ((w - 0.4) / 4);
      await badge(s, cx + 0.12, y + 0.4, 0.75, it[i][1], it[i][0]);
      s.addShape(S.ROUNDED_RECTANGLE, { x: cx + 0.02, y: y + 1.25, w: 0.95, h: 0.26, rectRadius: 0.05, fill: { color: "FFFFFF" }, line: { color: "FFFFFF", width: 0 } });
      T(s, it[i][2], { x: cx + 0.02, y: y + 1.25, w: 0.95, h: 0.26, fontSize: 10, bold: true, align: "center", valign: "middle" });
      if (i < 3) s.addShape(S.RIGHT_ARROW, { x: cx + 0.9, y: y + 0.68, w: 0.18, h: 0.2, fill: { color: HEX.gray }, line: { color: HEX.gray, width: 0 } });
    }
    await badge(s, x + 0.2, y + h - 0.65, 0.5, HEX.red, "FaHeart");
    T(s, "Tour guiado a la finca", { x: x + 0.8, y: y + h - 0.65, w: 2.5, h: 0.5, fontSize: 11, bold: true, valign: "middle" });
  }
  const storyMeta = (s, n, total, head) => {
    s.addShape(S.OVAL, { x: 0.5, y: 0.35, w: 0.6, h: 0.6, fill: { color: HEX.am }, line: { color: HEX.am, width: 0 } });
    T(s, String(n), { x: 0.5, y: 0.35, w: 0.6, h: 0.6, fontSize: 22, bold: true, fontFace: "Cambria", color: "FFFFFF", align: "center", valign: "middle" });
  };
  async function story(n, title, quote, text, panel, srcTxt, notes) {
    const s = pres.addSlide({ masterName: "CONTENT" });
    storyMeta(s, n);
    T(s, title, { x: 1.25, y: 0.3, w: 8.25, h: 0.7, fontSize: 28, bold: true, fontFace: "Cambria", color: C.text2, valign: "middle" });
    await panel(s, 5.2, 1.2, 4.3, 2.55);
    // cita
    card(s, 0.5, 1.2, 4.45, 0.95, HEX.lt2);
    T(s, quote, { x: 0.65, y: 1.2, w: 4.15, h: 0.95, fontSize: 14, italic: true, color: C.text2, valign: "middle", fontFace: "Cambria" });
    T(s, text, { x: 0.5, y: 2.3, w: 4.45, h: 2.55, fontSize: 14, valign: "top", paraSpaceAfter: 6 });
    return s;
  }
  {
    const s = await story(1, "Escena 1: Rosa tecnifica su finca", "“Vendía al precio que fijaba el acopiador. Hoy tengo un comprador y un técnico.”",
      "Rosa tiene 42 años y dos hectáreas de café en Chanchamayo. Cosecha unos 30 quintales y los vende al acopiador. Un técnico de Cosecha Viva le enseña a abonar con un plan de nutrición, instala fertirriego y un secador solar, y le propone un acuerdo: si cosecha solo los granos maduros, le compra ese café a un precio mayor.", sceneFarm);
    card(s, 5.2, 3.9, 4.3, 0.95, "FFFFFF", "C9D4C0");
    T(s, [{ text: "Qué cambia para Rosa: ", options: { bold: true, color: HEX.g } }, { text: "recibe el kit sin pagarlo por adelantado y un precio de S/ 20 por kg de pergamino (S/ 15 del intermediario)." }], { x: 5.35, y: 3.9, w: 4.0, h: 0.95, fontSize: 12, valign: "middle" });
    source(s, "Fuente: narrativa de la tesis (4.2). Kit en comodato de S/ 3,500 por productor nuevo. Cita ilustrativa.");
    s.addNotes("Escena 1. Rosa es la usuaria-beneficiaria. El kit se entrega en comodato, así que no necesita capital para empezar. La compra con prima del café selecto (S/ 20 por kg de pergamino frente a S/ 15) le da un incentivo para cosechar solo granos maduros. La cita es ilustrativa y no corresponde a una entrevista textual.");
  }
  {
    const s = await story(2, "Escena 2: Carla compra por adelantado", "“Quiero saber de dónde viene mi café y que mi compra le sirva a alguien.”",
      "En Lima, Carla, de 34 años, ve en Instagram un video de Rosa explicando cómo cambia de color el grano. Compra por adelantado un paquete de cosecha de la finca de Rosa por S/ 140. Durante los meses siguientes recibe por WhatsApp videos de la cosecha, del secado y del tueste.", sceneLima);
    card(s, 5.2, 3.9, 4.3, 0.95, "FFFFFF", "C9D4C0");
    T(s, [{ text: "Qué ve Carla: ", options: { bold: true, color: HEX.br } }, { text: "el café que financió, etapa por etapa. Su pago adelantado financia la tecnificación de Rosa." }], { x: 5.35, y: 3.9, w: 4.0, h: 0.95, fontSize: 12, valign: "middle" });
    source(s, "Fuente: narrativa de la tesis (4.2). El precio de S/ 140 se valida con la preventa del piloto (H5).");
    s.addNotes("Escena 2. El pago por adelantado es el corazón financiero del modelo: financia la tecnificación y reduce el capital de trabajo. El precio de S/ 140 es una decisión del equipo que debe confirmarse con compras reales en el piloto.");
  }
  {
    const s = await story(3, "Escena 3: la visita y la entrega", "“Probé el café recién cosechado con Rosa y semanas después llegó a mi casa.”",
      "En junio, Carla viaja a Chanchamayo con su pareja, recorre la finca con Rosa y prueba el café recién cosechado. Semanas después recibe en su casa cuatro bolsas de 250 g con la fecha de tueste y un código QR que lleva a la ficha de Rosa.", sceneHarvest);
    card(s, 5.2, 3.9, 4.3, 0.95, "FFFFFF", "C9D4C0");
    T(s, [{ text: "Segundo ingreso: ", options: { bold: true, color: HEX.am } }, { text: "el tour guiado (S/ 180 por persona) paga además a Rosa por cada visitante." }], { x: 5.35, y: 3.9, w: 4.0, h: 0.95, fontSize: 12, valign: "middle" });
    source(s, "Fuente: narrativa de la tesis (4.2) y supuestos del modelo financiero (tour de S/ 180 por persona con IGV).");
    s.addNotes("Escena 3. La visita es opcional y es el segundo ingreso del modelo: se supone que 5% de los clientes visita la finca con 2.5 personas por visita. El código QR por lote cierra la trazabilidad.");
  }
  {
    const s = pres.addSlide({ masterName: "CONTENT" });
    storyMeta(s, 4);
    T(s, "Escena 4: el resultado para Rosa y Carla", { x: 1.25, y: 0.3, w: 8.25, h: 0.7, fontSize: 28, bold: true, fontFace: "Cambria", color: C.text2, valign: "middle" });
    s.addChart(pres.charts.BAR, [{ name: "Ingreso neto anual (S/)", labels: ["Sin el modelo", "Primera campaña", "Tecnificación madura"], values: [20700, 20960, 34100] }], {
      x: 0.4, y: 1.15, w: 4.9, h: 3.0, barDir: "col", chartColors: [HEX.g], barGapWidthPct: 60, showValue: true, dataLabelFormatCode: "#,##0", dataLabelFontSize: 12, dataLabelColor: "1E1A16", dataLabelPosition: "outEnd",
      catAxisLabelFontSize: 11, catAxisLabelColor: "1E1A16", valAxisHidden: true, valAxisMaxVal: 40000, valAxisMinVal: 0, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
      showTitle: true, title: "Ingreso neto anual de Rosa (S/)", titleFontSize: 13, titleColor: "2F4A2B", catAxisLabelFontFace: "Calibri", dataLabelFontFace: "Calibri", titleFontFace: "Calibri",
    });
    T(s, "+64.7%", { x: 0.7, y: 4.2, w: 1.9, h: 0.6, fontSize: 32, bold: true, fontFace: "Cambria", color: HEX.g, valign: "middle" });
    T(s, "de ingreso neto con la tecnificación madura (+S/ 13,400)", { x: 2.6, y: 4.2, w: 2.7, h: 0.6, fontSize: 12, valign: "middle" });
    T(s, "En su primera campaña, Rosa cosecha 30 quintales y vende 12 como café selecto a mejor precio. En la segunda, con la tecnificación madura, cosecha 50 y vende 20 como selecto, y recibe un pago por cada visitante.", { x: 5.6, y: 1.2, w: 3.9, h: 1.9, fontSize: 13, valign: "top" });
    card(s, 5.6, 3.2, 3.9, 1.6, HEX.lt2);
    await badge(s, 5.75, 3.35, 0.5, HEX.br, "FaGift");
    T(s, "Carla renueva su paquete y regala otro a su madre: la recompra de 60% anual es el supuesto que el piloto debe confirmar (H8).", { x: 6.4, y: 3.3, w: 3.0, h: 1.4, fontSize: 12.5, valign: "middle" });
    source(s, "Fuente: hoja Productor del modelo financiero v5. Estimación; no es un aumento comprobado en campo.");
    s.addNotes("Escena 4. En la primera campaña el aumento es de apenas S/ 260, porque Rosa ya invierte en insumos pero todavía no cosecha más. Con la tecnificación madura, el ingreso neto anual pasa de S/ 20,700 a S/ 34,100, un aumento de 64.7%. Es una estimación del modelo con supuestos regionales, no un resultado medido en campo.");
  }

  // =============== 17. FLUJO DE CAJA LIBRE ===============
  const FCL = [-271700, 48543, 460485, 947356, 1394761, 2044317];
  const ACUM = [-271700, -231560, 83288, 618890, 1270926, 2061174];
  {
    const s = content("Flujo de caja libre y rentabilidad");
    const labels = ["Año 0", "Año 1", "Año 2", "Año 3", "Año 4", "Año 5"];
    s.addChart([
      { type: pres.charts.BAR, data: [{ name: "FCL = FCA (sin préstamo)", labels, values: FCL }], options: { chartColors: [HEX.g], barGapWidthPct: 50 } },
      { type: pres.charts.LINE, data: [{ name: "Flujo descontado acumulado", labels, values: ACUM }], options: { chartColors: [HEX.am], lineSize: 3, lineDataSymbolSize: 7 } },
    ], {
      x: 0.4, y: 1.1, w: 5.9, h: 3.8, showLegend: true, legendPos: "b", legendFontSize: 11, legendColor: "1E1A16", legendFontFace: "Calibri",
      catAxisLabelFontSize: 11, catAxisLabelColor: "1E1A16", valAxisLabelFontSize: 10, valAxisLabelColor: "5C5750", valAxisLabelFormatCode: "#,##0",
      valGridLine: { color: "E3E8DD", size: 0.75 }, catGridLine: { style: "none" }, catAxisLabelFontFace: "Calibri", valAxisLabelFontFace: "Calibri",
      showTitle: true, title: "Soles sin IGV", titleFontSize: 12, titleColor: "2F4A2B", titleFontFace: "Calibri",
    });
    const k = [["VAN económico (FCL al WACC)", "S/ 2,061,174", HEX.g], ["VAN financiero (FCA al COK)", "S/ 2,061,174", HEX.br], ["TIR del FCL y del FCA", "128.8%", HEX.bl], ["TIR modificada del FCA", "85.9%", HEX.am]];
    for (let i = 0; i < 4; i++) {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 6.5 + col * 1.54, y = 1.15 + row * 1.35;
      card(s, x, y, 1.45, 1.22, HEX.lt2);
      T(s, k[i][1], { x: x + 0.06, y: y + 0.12, w: 1.33, h: 0.55, fontSize: i < 2 ? 16 : 22, bold: true, fontFace: "Cambria", color: k[i][2], align: "center", valign: "middle" });
      T(s, k[i][0], { x: x + 0.06, y: y + 0.7, w: 1.33, h: 0.45, fontSize: 11.5, align: "center", valign: "top", bold: true });
    }
    card(s, 6.5, 3.9, 3.04, 1.0, "FFFFFF", "C9D4C0");
    T(s, [{ text: "Sin préstamo: FCA = FCL y WACC = COK = 20.94%. ", options: { bold: true, color: HEX.g } }, { text: "El año 1 ya es positivo (+S/ 48,543); recupero descontado de 1.74 años." }], { x: 6.62, y: 3.92, w: 2.82, h: 0.96, fontSize: 11.5, valign: "middle" });
    source(s, "Fuente: modelo financiero v5 (hojas FCL e Indicadores_Extra). Proyección con precio de S/ 140 aún no probado con compra real.");
    s.addNotes("El flujo de caja libre es negativo solo en el año 0 (S/ -271,700: inversión de S/ 156,200 más kits de S/ 115,500). Desde el año 1 es positivo. Con un COK de 20.94% el VAN financiero es de S/ 2,061,174, la TIR de 128.8% y la TIR modificada de 85.9%, que supone reinvertir al COK y es la medida más prudente. Advertencia: el precio de S/ 140 es una decisión del equipo que aún no se prueba con compra real; con el precio inicial de S/ 120 el VAN era de S/ 922,276 y el año 1 era negativo.");
  }

  // =============== 18. COK Y VAN ===============
  {
    const s = content("Del COK y el WACC al VAN");
    card(s, 0.5, 1.15, 3.8, 3.75, HEX.lt2);
    T(s, "Costo de oportunidad (CAPM, beta total)", { x: 0.65, y: 1.22, w: 3.6, h: 0.35, fontSize: 13.5, bold: true, fontFace: "Cambria", color: C.text2 });
    const cok = [["Tasa libre de riesgo (bono EE. UU. 10 años)", "4.88%"], ["Beta total, sector agrícola", "3.56"], ["Prima de riesgo de mercado", "4.14%"], ["Riesgo país (EMBIG Perú)", "1.18%"], ["COK en dólares = 4.88% + 3.56 × 4.14% + 1.18%", "20.80%"], ["Inflación esperada: Perú 2.25% y EE. UU. 2.13%", "ajuste"], ["COK en soles", "20.94%"]];
    cok.forEach((r, i) => {
      const y = 1.65 + i * 0.4; const last = i === cok.length - 1;
      T(s, r[0], { x: 0.65, y, w: 2.65, h: 0.38, fontSize: 10.5, bold: last || i === 4, valign: "middle" });
      T(s, r[1], { x: 3.3, y, w: 0.9, h: 0.38, fontSize: last ? 16 : 12, bold: true, color: last ? HEX.g : C.text1, align: "right", valign: "middle" });
    });
    T(s, "VAN = flujo del año 0 + suma de cada flujo / (1 + tasa) elevado al año. FCL con WACC; FCA con COK.", { x: 4.5, y: 1.15, w: 5.0, h: 0.45, fontSize: 11.5, bold: true, color: C.text2, valign: "middle" });
    const hd = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: HEX.dk2 }, align: "center", fontSize: 11 } });
    const cl = (t, b) => ({ text: t, options: { align: "right", fontSize: 10, bold: !!b } });
    const lb = (t) => ({ text: t, options: { bold: true, fontSize: 10.5, fill: { color: HEX.lt2 } } });
    const rows = [[hd("S/"), hd("Año 0"), hd("Año 1"), hd("Año 2"), hd("Año 3"), hd("Año 4"), hd("Año 5")],
      [lb("FCL = FCA"), ...["-271,700", "48,543", "460,485", "947,356", "1,394,761", "2,044,317"].map((v) => cl(v))],
      [lb("FCL descontado"), ...["-271,700", "40,140", "314,848", "535,602", "652,037", "790,248"].map((v) => cl(v))],
      [lb("Acumulado"), ...["-271,700", "-231,560", "83,288", "618,890", "1,270,926", "2,061,174"].map((v, i) => cl(v, i === 5))]];
    s.addTable(rows, { x: 4.5, y: 1.65, w: 5.0, colW: [0.85, 0.66, 0.62, 0.66, 0.7, 0.75, 0.76], rowH: 0.4, border: { type: "solid", color: "DDE5D5", pt: 1 }, fontFace: "Calibri", margin: [0.02, 0.04, 0.02, 0.04] });
    const kp = [["VAN económico (FCL al WACC)", "S/ 2,061,174"], ["VAN financiero (FCA al COK)", "S/ 2,061,174"], ["WACC = COK (sin deuda)", "20.94%"]];
    kp.forEach((r, i) => {
      const y = 3.4 + i * 0.5;
      s.addShape(S.ROUNDED_RECTANGLE, { x: 4.5, y, w: 5.0, h: 0.42, rectRadius: 0.06, fill: { color: i === 0 ? HEX.g : "FFFFFF" }, line: { color: HEX.g, width: 1 } });
      T(s, r[0], { x: 4.65, y, w: 2.8, h: 0.42, fontSize: 12.5, bold: true, color: i === 0 ? "FFFFFF" : C.text1, valign: "middle" });
      T(s, r[1], { x: 7.5, y, w: 1.9, h: 0.42, fontSize: 15, bold: true, fontFace: "Cambria", color: i === 0 ? "FFFFFF" : HEX.g, valign: "middle", align: "right" });
    });
    source(s, "Fuentes: Damodaran (2026a, 2026b); RoboForex (2026); Andina (2026); BCRP (2026); Federal Reserve (2026). Modelo financiero v5.");
    s.addNotes("El proyecto no usa préstamo, así que el FCA es igual al FCL y el WACC es igual al COK; si hubiera deuda, el FCA sería distinto del FCL y el WACC menor que el COK. El FCL se descuenta con el WACC y el FCA con el COK. El COK se calcula con CAPM usando la beta total del sector agrícola, porque el riesgo dominante es el abastecimiento. El COK en dólares es 20.80% y, con las expectativas de inflación de Perú y de Estados Unidos, el COK en soles es 20.94%. El VAN financiero descuenta el flujo de caja libre a esa tasa; la tasa social de 8% se usa solo para el VAN social.");
  }

  // =============== 19. ODS ===============
  {
    const s = content("Cada ODS genera un beneficio social");
    const od = [
      ["2", "Hambre cero", "Meta 2.3", "Duplicar la productividad y los ingresos de pequeños productores", "Beneficio 1: valor de la producción adicional", "De 15 a 25 qq/ha: +20 qq por productor maduro × S/ 690", "DDA63A"],
      ["1", "Fin de la pobreza", "Meta 1.2", "Eleva el ingreso neto del productor", "Beneficio 2: sobreprecio del café selecto", "20 qq selectos × S/ 230 por qq sobre el intermediario", "E5243B"],
      ["8", "Trabajo decente y crecimiento económico", "Metas 8.3 y 8.9", "Empleo calificado en el campo e ingreso local por turismo", "Beneficios 3 y 4: empleo de técnicos y gasto local de los tours", "Remuneración de técnicos (S/ 37,500 al año) y gasto local de los tours", "A21942"],
    ];
    T(s, "ODS · meta", { x: 1.45, y: 1.05, w: 2.4, h: 0.25, fontSize: 10.5, bold: true, color: HEX.gray });
    T(s, "Contribución de la solución", { x: 3.95, y: 1.05, w: 2.6, h: 0.25, fontSize: 10.5, bold: true, color: HEX.gray });
    T(s, "Beneficio social en el VAN social", { x: 6.65, y: 1.05, w: 2.8, h: 0.25, fontSize: 10.5, bold: true, color: HEX.gray });
    for (let i = 0; i < 3; i++) {
      const y = 1.35 + i * 1.12;
      s.addShape(S.ROUNDED_RECTANGLE, { x: 0.5, y, w: 9.0, h: 1.0, rectRadius: 0.06, fill: { color: i % 2 ? "FFFFFF" : HEX.lt2 }, line: { color: "DDE5D5", width: 0.75 } });
      s.addShape(S.RECTANGLE, { x: 0.58, y: y + 0.1, w: 0.7, h: 0.8, fill: { color: od[i][6] }, line: { color: od[i][6], width: 0 } });
      T(s, od[i][0], { x: 0.58, y: y + 0.1, w: 0.7, h: 0.8, fontSize: 28, bold: true, fontFace: "Cambria", color: "FFFFFF", align: "center", valign: "middle" });
      T(s, [{ text: od[i][1] + "\n", options: { bold: true, fontSize: 13 } }, { text: od[i][2], options: { fontSize: 11, color: HEX.gray } }], { x: 1.4, y, w: 2.45, h: 1.0, valign: "middle" });
      T(s, od[i][3], { x: 3.95, y, w: 2.55, h: 1.0, fontSize: 12, valign: "middle" });
      T(s, [{ text: od[i][4] + "\n", options: { bold: true, fontSize: 12, color: C.text2 } }, { text: od[i][5], options: { fontSize: 11 } }], { x: 6.65, y, w: 2.75, h: 1.0, valign: "middle" });
    }
    source(s, "Adaptado de Naciones Unidas (2015), Agenda 2030 (A/RES/70/1), y de la Tabla 29 de la tesis (capítulo VII). Hojas Beneficios Sociales y VAN Social del modelo v5.");
    s.addNotes("La tesis declara tres ODS y cada uno genera un beneficio social que entra al VAN social. ODS 2, meta 2.3: el productor pasa de 15 a 25 quintales por hectárea, lo que da 20 quintales adicionales por productor maduro, valorados al precio del intermediario de S/ 690. ODS 1, meta 1.2: recibe S/ 230 más por quintal de café selecto frente al intermediario. ODS 8, metas 8.3 y 8.9: la remuneración del empleo calificado de los técnicos de campo y el gasto local de los tours. En el quinto año el modelo trabaja con 117 familias productoras y unos 11,000 clientes activos. Los colores de cada número son los oficiales del ODS; no se reproducen los logotipos.");
  }

  // =============== 20. RENTABILIDAD SOCIAL ===============
  {
    const s = content("VAN social: beneficios menos costos");
    const yl = ["Año 1", "Año 2", "Año 3", "Año 4", "Año 5"];
    s.addChart(pres.charts.BAR, [
      { name: "ODS 2: producción adicional", labels: yl, values: [0, 455400, 676200, 1062600, 1324800] },
      { name: "ODS 1: sobreprecio selecto", labels: yl, values: [91080, 195960, 302680, 406640, 499560] },
      { name: "ODS 8: empleo y turismo local", labels: yl, values: [97881, 124607, 188741, 252966, 275847] },
    ], {
      x: 0.4, y: 1.1, w: 5.2, h: 3.8, barDir: "col", barGrouping: "stacked", chartColors: ["DDA63A", "E5243B", "A21942"], barGapWidthPct: 45,
      showLegend: true, legendPos: "b", legendFontSize: 10.5, legendColor: "1E1A16", legendFontFace: "Calibri",
      catAxisLabelFontSize: 11, catAxisLabelColor: "1E1A16", valAxisLabelFontSize: 10, valAxisLabelColor: "5C5750", valAxisLabelFormatCode: "#,##0",
      valGridLine: { color: "E3E8DD", size: 0.75 }, catGridLine: { style: "none" }, catAxisLabelFontFace: "Calibri", valAxisLabelFontFace: "Calibri",
      showTitle: true, title: "Beneficios sociales por ODS (S/ por año)", titleFontSize: 12, titleColor: "2F4A2B", titleFontFace: "Calibri",
    });
    T(s, "S/ 3,279,853", { x: 5.85, y: 1.2, w: 3.65, h: 0.8, fontSize: 36, bold: true, fontFace: "Cambria", color: HEX.g, valign: "middle" });
    T(s, "VAN social al 8% (tasa social del MEF)", { x: 5.85, y: 2.0, w: 3.65, h: 0.35, fontSize: 13 });
    const lines = [["S/ 4,462,364", "VP de los beneficios sociales de los ODS 1, 2 y 8", HEX.g], ["S/ 1,182,511", "VP de los costos sociales: insumos y cosecha selectiva adicionales", HEX.am], ["3.77", "Relación beneficio-costo social", HEX.br]];
    lines.forEach((r, i) => {
      const y = 2.5 + i * 0.7;
      T(s, r[0], { x: 5.85, y, w: 1.45, h: 0.6, fontSize: 15, bold: true, fontFace: "Cambria", color: r[2], valign: "middle" });
      T(s, r[1], { x: 7.35, y, w: 2.15, h: 0.6, fontSize: 10.5, valign: "middle" });
    });
    card(s, 5.85, 4.55, 3.65, 0.35, HEX.lt2);
    T(s, "Sin inversión ni flujo de caja libre", { x: 5.95, y: 4.55, w: 3.45, h: 0.35, fontSize: 11.5, bold: true, color: C.text2, valign: "middle" });
    source(s, "Fuente: modelo financiero v5 (hojas Beneficios Sociales, Costos Sociales y VAN Social). Estimación con supuestos regionales del productor.");
    s.addNotes("El VAN social es beneficio social menos costo social, descontado al 8%. No usa el flujo de caja libre ni la inversión de la empresa. Hay un beneficio por cada ODS declarado: ODS 2, valor de la producción adicional; ODS 1, sobreprecio del café selecto; ODS 8, empleo calificado de los técnicos y gasto local de los tours. Los costos son los insumos y la cosecha selectiva adicionales que paga el productor. El VAN social es S/ 3,279,853 y la relación beneficio-costo social es 3.77. Es una estimación con supuestos regionales; se recalcula con las fincas piloto.");
  }

  // =============== 21. CIERRE ===============
  {
    const s = pres.addSlide({ masterName: "DARK" });
    T(s, "La decisión: probar antes de escalar", { x: 0.5, y: 0.4, w: 9.0, h: 0.7, fontSize: 32, bold: true, fontFace: "Cambria", color: "FFFFFF" });
    const it = [["FaSeedling", "10 fincas piloto", "H4: rendimiento de 22 qq/ha o más en la segunda campaña"], ["FaShoppingCart", "300 paquetes", "H5: 3% de conversión con pago real; probar S/ 140, S/ 130 y S/ 120"], ["FaHandshake", "15 de 30", "H6: productores que firman el acuerdo de compra"], ["FaRedo", "60% de recompra", "H8: recompra y 5% de reservas de tour"]];
    for (let i = 0; i < 4; i++) {
      const x = 0.5 + i * 2.3;
      s.addShape(S.ROUNDED_RECTANGLE, { x, y: 1.4, w: 2.1, h: 2.25, rectRadius: 0.08, fill: { color: "2F4A2B" }, line: { color: "2F4A2B", width: 0 } });
      await badge(s, x + 0.15, 1.55, 0.55, HEX.gold, it[i][0], "1E1A16");
      T(s, it[i][1], { x: x + 0.15, y: 2.2, w: 1.85, h: 0.4, fontSize: 15, bold: true, color: "FFFFFF", fontFace: "Cambria" });
      T(s, it[i][2], { x: x + 0.15, y: 2.6, w: 1.85, h: 1.0, fontSize: 11.5, color: "F1F5EC", valign: "top" });
    }
    T(s, "Piloto de seis meses con presupuesto de S/ 93,550. Se escala solo si H5 y H6 se cumplen y el VAN recalculado es positivo.", { x: 0.5, y: 3.95, w: 9.0, h: 0.6, fontSize: 16, color: "E3B23C", bold: true });
    T(s, "Los resultados son proyecciones condicionadas: el precio de S/ 140 aún no se ha probado con compra real.", { x: 0.5, y: 4.55, w: 9.0, h: 0.4, fontSize: 13, italic: true, color: "F1F5EC" });
    s.addNotes("Cierre. La evidencia de campo valida interés, pero la compra anticipada del producto definitivo se prueba en un piloto de seis meses con criterios de decisión definidos antes de conocer los resultados.");
  }

  // =============== 22-23. REFERENCIAS ===============
  const refs1 = [
    "Agraria.pe. (2019b, 11 de julio). Yara lanzará “NuestroCafé” para incrementar rendimiento productivo y calidad de taza. https://agraria.pe/noticias/yara-lanzara-nuestrocafe-para-incrementar-rendimiento-produc-19392",
    "Agraria.pe. (2025, 26 de diciembre). Consumo de café en Perú alcanza los 1.7 kilos por persona al año. https://agraria.pe/noticias/consumo-de-cafe-en-peru-alcanza-los-1-7-kilos-por-persona-al-41376",
    "Agraria.pe. (2026, 26 de mayo). Perú desaprovecha coyuntura de precios históricos en café y cacao por falta de tecnología. https://agraria.pe/noticias/peru-desaprovecha-coyuntura-de-precios-historicos-en-cafe-y--42338",
    "Andina. (2026, 16 de noviembre). Peru: Country risk down to 118, below region’s average. https://andina.pe/agencia/noticia-peru-country-risk-down-to-118-below-regions-average-1100303.aspx",
    "Banco Central de Reserva del Perú. (2026, 15 de abril). Expectativas de inflación a 12 meses siguen en el rango meta [Nota informativa]. https://www.bcrp.gob.pe/docs/Transparencia/Notas-Informativas/2026/nota-informativa-2026-04-15.pdf",
    "Damodaran, A. (2026a). Total betas by sector (for computing private company costs of equity): US [Base de datos]. Stern School of Business, New York University. https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/totalbeta.html",
    "Damodaran, A. (2026b). Damodaran Online [Implied ERP on September 1, 2026]. Stern School of Business, New York University. https://pages.stern.nyu.edu/~adamodar/New_Home_Page/home.htm",
    "El Peruano. (2026, 17 de julio). Café peruano llegó a más de 50 países. https://elperuano.pe/noticia/288443-cafe-peruano-llego-a-mas-de-50-paises",
    "El Peruano. (2026, 28 de agosto). Aroma que conquista el mundo. https://elperuano.pe/noticia/303531-aroma-que-conquista-el-mundo",
  ];
  const refs2 = [
    "Federal Reserve. (2026, 16 de septiembre). Summary of economic projections. https://www.federalreserve.gov/monetarypolicy/files/fomcprojtabl20260916.pdf",
    "Gestión. (2025). Precio del café cae a mínimos de siete meses: ¿cómo impacta en el Perú? https://gestion.pe/economia/precio-del-cafe-cae-a-minimos-de-siete-meses-como-impacta-en-el-peru-noticia/",
    "Mercado Libre Perú. (2026). Café Chanchamayo [Listado de productos]. https://listado.mercadolibre.com.pe/cafe-chanchamayo",
    "Naciones Unidas. (2015). Transformar nuestro mundo: la Agenda 2030 para el Desarrollo Sostenible (A/RES/70/1).",
    "RoboForex. (2026, 10 de noviembre). ¿Por qué sube el rendimiento del bono americano a 10 años? https://roboforex.com/es/blog/2026/11/10/why-are-treasury-yields-rising/",
    "Trabajo de campo del equipo. (2026). Encuestas a 20 pequeños productores de café y 25 turistas potenciales en Chanchamayo (abril y mayo) [Datos no publicados]. Tesis Cosecha Viva, Apéndice A.",
    "Pampa, D. W., Quinto, H. H., Aguilar, R. H., & Estrada, F. (2026). Cosecha Viva: preventa de café tecnificado con trazabilidad para elevar los ingresos de los pequeños productores de Chanchamayo [Tesis de maestría]. Pontificia Universidad Católica del Perú.",
  ];
  function refSlide(title, arr, extra) {
    const s = content(title);
    const runs = [];
    arr.forEach((r, i) => runs.push({ text: r, options: { breakLine: i < arr.length - 1, fontSize: 10, paraSpaceAfter: 4, hanging: 0.25, indentLevel: 0 } }));
    s.addText(runs, { x: 0.5, y: 1.1, w: 9.0, h: extra ? 2.75 : 3.75, fontSize: 10, color: C.text1, margin: 0, valign: "top", fontFace: "Calibri", isTextBox: true, objectName: "Referencias" });
    return s;
  }
  {
    const s = refSlide("Referencias (1 de 2)", refs1);
    s.addNotes("Referencias en formato APA 7. Las dos entradas de El Peruano de 2026 se distinguen por fecha; en la tesis conviene rotularlas 2026a y 2026b.");
  }
  {
    const s = refSlide("Referencias (2 de 2) y créditos", refs2, true);
    card(s, 0.5, 3.95, 9.0, 0.95, HEX.lt2);
    T(s, [{ text: "Créditos de imágenes. ", options: { bold: true } }, { text: "Fotografías del trabajo de campo: archivo del equipo (Apéndice E de la tesis). Íconos: React Icons, conjunto Font Awesome (licencia CC BY 4.0, https://react-icons.github.io/react-icons). Ilustraciones de escenas: elaboración propia con formas e íconos. Los colores de los ODS siguen la paleta oficial de Naciones Unidas; no se reproducen logotipos." }], { x: 0.65, y: 3.97, w: 8.7, h: 0.9, fontSize: 10, valign: "middle" });
    s.addNotes("Créditos de imágenes e íconos.");
  }

  await pres.writeFile({ fileName: "Cosecha_Viva_PSR.pptx" });
  await applyTheme("Cosecha_Viva_PSR.pptx", THEME);
  console.log("ok");
})().catch((e) => { console.error(e); process.exit(1); });
