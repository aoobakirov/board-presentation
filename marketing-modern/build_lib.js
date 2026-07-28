// KazBioFert — modern marketing deck (RU). Rebuild of the 9-slide source.
// Layout code is shared between the real pptxgenjs emitter and the HTML
// preview recorder (for visual QA). buildAll(pres) drives both.
const path = require("path");
const A = (f) => path.join(__dirname, "assets", f);
const IC = (k, v = "green") => path.join(__dirname, "assets", "icons", `${k}_${v}.png`);

// ---------- palette ----------
const C = {
  deep:  "0F2A1F",   // deep forest (dark bg base)
  deep2: "0A1C15",
  green: "2D9C5A",   // primary accent
  greenD:"1C6E40",   // darker green
  amber: "E0A526",   // harvest accent
  ink:   "15211B",   // headline text
  body:  "3E4A43",   // body text
  mute:  "77857B",   // captions
  light: "F7FAF6",   // light bg
  white: "FFFFFF",
  tint:  "EDF4E9",   // soft green card tint
  tint2: "E3EEDD",
  line:  "D9E4D8",   // hairline
};
const FONT = "Calibri";
const FONTH = "Calibri";
const PW = 13.333, PH = 7.5, M = 0.62;

function buildAll(pres) {

function shadow(opt = {}) {
  return Object.assign({ type: "outer", color: "1A3D2A", opacity: 0.20, blur: 9, offset: 3, angle: 90 }, opt);
}

// soft card
function card(s, x, y, w, h, o = {}) {
  s.addShape("roundRect", {
    x, y, w, h, rectRadius: o.r ?? 0.11,
    fill: { color: o.fill ?? C.white },
    line: o.line ? { color: o.line, width: o.lw ?? 1 } : { type: "none" },
    shadow: o.noShadow ? undefined : shadow(o.shadow || {}),
  });
}

// icon inside a colored circle
function iconCircle(s, x, y, d, key, o = {}) {
  s.addShape("ellipse", {
    x, y, w: d, h: d,
    fill: { color: o.fill ?? C.tint },
    line: { type: "none" },
    shadow: o.shadow ? shadow({ opacity: 0.15, blur: 6, offset: 2 }) : undefined,
  });
  const pad = d * (o.pad ?? 0.30);
  s.addImage({ path: IC(key, o.variant ?? "green"), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad });
}

// content-slide header (kicker + title). returns y below header
function header(s, kicker, title, o = {}) {
  const x = o.x ?? M, y = o.y ?? 0.5;
  s.addText(kicker.toUpperCase(), {
    x, y, w: o.w ?? 11.5, h: 0.3, fontFace: FONTH, fontSize: 12, bold: true,
    color: C.green, charSpacing: 3, align: "left",
  });
  s.addText(title, {
    x, y: y + 0.32, w: o.w ?? 11.6, h: o.th ?? 0.68, fontFace: FONTH, fontSize: o.ts ?? 30, bold: true,
    color: C.ink, align: "left", lineSpacing: (o.ts ?? 30) * 1.06,
  });
  return y + 0.32 + (o.th ?? 0.68);
}

function footer(s, n) {
  s.addText([
    { text: "KAZBIOFERT", options: { color: C.green, bold: true } },
    { text: "   ·   Органоминеральные удобрения", options: { color: C.mute } },
  ], { x: M, y: PH - 0.42, w: 7, h: 0.28, fontFace: FONT, fontSize: 9.5, align: "left", charSpacing: 1 });
  s.addText(String(n).padStart(2, "0"), {
    x: PW - 1.3, y: PH - 0.42, w: 0.68, h: 0.28, fontFace: FONT, fontSize: 9.5, color: C.mute, align: "right",
  });
}

// big stat callout
function stat(s, x, y, w, num, label, o = {}) {
  s.addText(num, { x, y, w, h: o.nh ?? 0.72, fontFace: FONTH, fontSize: o.ns ?? 40, bold: true,
    color: o.color ?? C.green, align: o.align ?? "left", margin: 0 });
  s.addText(label, { x, y: y + (o.nh ?? 0.72) - 0.04, w, h: o.lh ?? 0.5, fontFace: FONT, fontSize: o.ls ?? 12,
    color: o.lcolor ?? C.body, align: o.align ?? "left", margin: 0, lineSpacing: (o.ls ?? 12) * 1.12 });
}

// bullet list block
function bullets(s, items, x, y, w, o = {}) {
  s.addText(items.map((t, i) => ({
    text: t, options: {
      bullet: { code: "2022", indent: 14 }, color: o.color ?? C.body, breakLine: true,
      paraSpaceAfter: o.gap ?? 7, fontSize: o.fs ?? 13.5,
    },
  })), { x, y, w, h: o.h ?? 3, fontFace: FONT, align: "left", valign: "top", lineSpacing: (o.fs ?? 13.5) * 1.16 });
}

/* =========================================================================
   SLIDE 1 — TITLE
   ========================================================================= */
{
  const s = pres.addSlide();
  s.background = { path: A("bg_dark.jpg") };
  // right hero panel (bleeds to right edge)
  const pw = 4.55, ph = 6.35;
  s.addImage({ path: A("r_hero_panel.png"), x: PW - pw - 0.5, y: (PH - ph) / 2, w: pw, h: ph, rounding: false,
    shadow: shadow({ color: "000000", opacity: 0.4, blur: 18, offset: 6 }) });

  // left text
  const lx = M + 0.15;
  s.addText("KAZBIOFERT", { x: lx, y: 0.72, w: 6, h: 0.5, fontFace: FONTH, fontSize: 17, bold: true,
    color: C.green, charSpacing: 4 });
  s.addText("КРАТКИЙ МАРКЕТИНГОВЫЙ ОБЗОР РЫНКА УДОБРЕНИЙ", {
    x: lx, y: 2.05, w: 7.0, h: 0.35, fontFace: FONTH, fontSize: 12.5, bold: true, color: C.amber, charSpacing: 2 });
  s.addText("Производство\nорганоминеральных\nудобрений", {
    x: lx, y: 2.45, w: 7.2, h: 2.5, fontFace: FONTH, fontSize: 47, bold: true, color: C.white, lineSpacing: 48 });
  // divider
  s.addShape("line", { x: lx + 0.02, y: 5.28, w: 2.0, h: 0, line: { color: C.green, width: 2.5 } });
  s.addText("Биогазовый завод с интегрированным производством ОМУ  ·  75 000 тонн в год", {
    x: lx, y: 5.45, w: 7.0, h: 0.7, fontFace: FONT, fontSize: 14.5, color: "CFE3D6", lineSpacing: 20 });
}

/* =========================================================================
   SLIDE 2 — ВИДЫ УДОБРЕНИЙ
   ========================================================================= */
{
  const s = pres.addSlide();
  header(s, "Основы", "Виды удобрений и элементы питания");

  // left: NPK elements
  const lx = M, ly = 2.0;
  s.addText("Основные элементы питания", { x: lx, y: ly - 0.42, w: 4.4, h: 0.3, fontFace: FONTH,
    fontSize: 13, bold: true, color: C.ink });
  const el = [
    ["N", "Азот", C.green],
    ["P", "Фосфор", C.greenD],
    ["K", "Калий", C.amber],
  ];
  el.forEach((e, i) => {
    const y = ly + i * 1.02;
    s.addShape("ellipse", { x: lx, y, w: 0.8, h: 0.8, fill: { color: e[2] }, line: { type: "none" },
      shadow: shadow({ opacity: 0.18, blur: 6, offset: 2 }) });
    s.addText(e[0], { x: lx, y, w: 0.8, h: 0.8, fontFace: FONTH, fontSize: 30, bold: true, color: C.white,
      align: "center", valign: "middle", margin: 0 });
    s.addText(e[1], { x: lx + 1.0, y: y + 0.06, w: 3.2, h: 0.7, fontFace: FONTH, fontSize: 18, bold: true,
      color: C.ink, valign: "middle" });
  });
  s.addText("N · P · K — три ключевых макроэлемента, определяющих питание и урожайность культур.", {
    x: lx, y: ly + 3.2, w: 4.4, h: 0.9, fontFace: FONT, fontSize: 12.5, color: C.body, italic: true, lineSpacing: 17 });

  // right: 4 type rows
  const rx = 5.6, rw = PW - rx - M;
  const types = [
    ["leaf",   "Органические",            "Из продуктов животного и растительного происхождения.", false],
    ["atom",   "Простые",                 "Содержат только один элемент — N, P или K.", false],
    ["flask",  "Минеральные (неорг.)",    "Из химических веществ; могут содержать NPK.", false],
    ["seedling","Органоминеральные · ОМУ","Смесь органических и минеральных в одной грануле. NPK корректируется под запрос.", true],
  ];
  const rh = 1.12, gap = 0.14, ry0 = 1.62;
  types.forEach((t, i) => {
    const y = ry0 + i * (rh + gap);
    const hi = t[3];
    card(s, rx, y, rw, rh, { fill: hi ? C.tint : C.white, r: 0.09,
      shadow: { opacity: hi ? 0.16 : 0.12 } });
    iconCircle(s, rx + 0.24, y + (rh - 0.72) / 2, 0.72, t[0], { fill: hi ? C.green : C.tint, variant: hi ? "white" : "green" });
    s.addText(t[1], { x: rx + 1.16, y: y + 0.16, w: rw - 1.35, h: 0.36, fontFace: FONTH, fontSize: 15.5, bold: true,
      color: hi ? C.greenD : C.ink });
    s.addText(t[2], { x: rx + 1.16, y: y + 0.5, w: rw - 1.35, h: 0.5, fontFace: FONT, fontSize: 12, color: C.body, lineSpacing: 15 });
    if (hi) s.addText("ОМУ", { x: rx + rw - 1.15, y: y + 0.14, w: 0.95, h: 0.32, fontFace: FONTH, fontSize: 11,
      bold: true, color: C.white, align: "center", valign: "middle", fill: { color: C.amber }, rectRadius: 0.1 });
  });
  // ОМУ subtypes note
  s.addText([
    { text: "Виды ОМУ:  ", options: { bold: true, color: C.greenD } },
    { text: "комплексные", options: { bold: true, color: C.ink } },
    { text: " — каждая гранула содержит NPK + Fe, Mg, S;   ", options: { color: C.body } },
    { text: "смешанные", options: { bold: true, color: C.ink } },
    { text: " — смесь простых гранул, каждая с одним элементом.", options: { color: C.body } },
  ], { x: rx, y: ry0 + 4 * (rh + gap) + 0.02, w: rw, h: 0.5, fontFace: FONT, fontSize: 11.5, lineSpacing: 15, align: "left" });

  footer(s, 2);
}

/* =========================================================================
   SLIDE 3 — БИОГАЗОВЫЙ ЗАВОД
   ========================================================================= */
{
  const s = pres.addSlide();
  header(s, "Производство", "Биогазовый завод с производством удобрений");

  // left: 3D render
  const iw = 6.0, ih = 4.05, ix = M, iy = 1.95;
  s.addImage({ path: A("r_plant3d.png"), x: ix, y: iy, w: iw, h: ih,
    shadow: shadow({ opacity: 0.25, blur: 12, offset: 5 }) });
  s.addText("Визуализация производственного комплекса", { x: ix, y: iy + ih + 0.08, w: iw, h: 0.3,
    fontFace: FONT, fontSize: 10.5, italic: true, color: C.mute });

  // right: stat cards 2x2 + compost text
  const rx = 6.95, rw = PW - rx - M;
  const stats = [
    ["17 га", "площадь объекта", C.green],
    ["70 000 м²", "производственные здания", C.greenD],
    ["75 000 т", "ОМУ в год", C.amber],
    ["7 560 м²", "площадь производства", C.green],
  ];
  const cw = (rw - 0.24) / 2, ch = 1.12;
  stats.forEach((st, i) => {
    const cx = rx + (i % 2) * (cw + 0.24);
    const cy = 1.95 + Math.floor(i / 2) * (ch + 0.24);
    card(s, cx, cy, cw, ch, { r: 0.1, shadow: { opacity: 0.12 } });
    s.addText(st[0], { x: cx + 0.22, y: cy + 0.16, w: cw - 0.4, h: 0.55, fontFace: FONTH, fontSize: 24, bold: true,
      color: st[2], margin: 0 });
    s.addText(st[1], { x: cx + 0.22, y: cy + 0.7, w: cw - 0.4, h: 0.34, fontFace: FONT, fontSize: 11.5, color: C.body, margin: 0 });
  });
  // compost text block
  const ty = 1.95 + 2 * (ch + 0.24) + 0.05;
  card(s, rx, ty, rw, 1.5, { fill: C.tint, r: 0.1, noShadow: true });
  s.addText([
    { text: "Сырьё — компост животного происхождения.  ", options: { bold: true, color: C.greenD } },
    { text: "Готовится ~3 мес. мезофильным способом (анаэробно) и ~3 мес. аэробно в температурных ячейках с термообработкой. Итог: до ", options: { color: C.body } },
    { text: "NPK 2-2-2", options: { bold: true, color: C.ink } },
    { text: ", без запаха и примесей, обеззаражен, без патогенов и грибов.", options: { color: C.body } },
  ], { x: rx + 0.24, y: ty + 0.16, w: rw - 0.48, h: 1.2, fontFace: FONT, fontSize: 12, lineSpacing: 16, valign: "middle", align: "left" });

  footer(s, 3);
}

/* =========================================================================
   SLIDE 4 — СХЕМА ПРОИЗВОДСТВА
   ========================================================================= */
{
  const s = pres.addSlide();
  header(s, "Технология", "Схема производства органоминеральных удобрений");

  // left: process diagram
  const iw = 5.7, ih = 4.3, ix = M, iy = 1.9;
  s.addImage({ path: A("r_process.png"), x: ix, y: iy, w: iw, h: ih,
    shadow: shadow({ opacity: 0.25, blur: 12, offset: 5 }) });

  const rx = 6.7, rw = PW - rx - M;
  s.addText([
    { text: "Компост животного происхождения ", options: { bold: true, color: C.greenD } },
    { text: "смешивается с минеральными компонентами (N, P, K) в зависимости от нужной формы конечного продукта.", options: { color: C.body } },
  ], { x: rx, y: 1.95, w: rw, h: 0.85, fontFace: FONT, fontSize: 13.5, lineSpacing: 18, align: "left" });

  s.addText("ПЛАНИРУЕМЫЕ ПРОДУКТЫ", { x: rx, y: 2.85, w: rw, h: 0.28, fontFace: FONTH, fontSize: 11, bold: true,
    color: C.mute, charSpacing: 2 });
  ["NPK 12-12-12", "NP 8-21 (без калия)"].forEach((t, i) => {
    const cx = rx + i * 2.85;
    s.addShape("roundRect", { x: cx, y: 3.16, w: 2.65, h: 0.56, rectRadius: 0.28,
      fill: { color: i === 0 ? C.green : C.greenD }, line: { type: "none" }, shadow: shadow({ opacity: 0.16, blur: 5, offset: 2 }) });
    s.addText(t, { x: cx, y: 3.16, w: 2.65, h: 0.56, fontFace: FONTH, fontSize: 14, bold: true, color: C.white,
      align: "center", valign: "middle", margin: 0 });
  });

  // FAO efficiency callout
  const fy = 4.0;
  card(s, rx, fy, rw, 2.35, { fill: C.deep, r: 0.11, noShadow: false, shadow: { opacity: 0.22, blur: 10, offset: 4 } });
  iconCircle(s, rx + 0.28, fy + 0.26, 0.64, "balance", { fill: C.green, variant: "white" });
  s.addText("Эффективность 1 : 3 по данным FAO", { x: rx + 1.06, y: fy + 0.28, w: rw - 1.3, h: 0.6, fontFace: FONTH,
    fontSize: 15, bold: true, color: C.white, valign: "middle" });
  s.addText("Питательные вещества в ОМУ приравнены 1 к 3 к минеральным удобрениям — за счёт растворимости, органических связей, усвояемости и удержания влаги «органической губкой» в почве.", {
    x: rx + 0.32, y: fy + 0.92, w: rw - 0.6, h: 0.72, fontFace: FONT, fontSize: 11.5, color: "CFE3D6", lineSpacing: 15, align: "left" });
  // two equivalences
  const eq = [["NPK 36-36-36 мин.", "NPK 12-12-12 ОМУ"], ["NP 12-52 аммофос", "NP 8-21 ОМУ"]];
  eq.forEach((e, i) => {
    const ey = fy + 1.62 + i * 0.36;
    s.addText([
      { text: e[0], options: { color: "9FC7AE" } },
      { text: "   =   ", options: { color: C.amber, bold: true } },
      { text: e[1], options: { color: C.white, bold: true } },
    ], { x: rx + 0.34, y: ey, w: rw - 0.6, h: 0.32, fontFace: FONT, fontSize: 12, align: "left", valign: "middle", margin: 0 });
  });

  footer(s, 4);
}

/* =========================================================================
   SLIDE 5 — ПОЛЕВЫЕ ОПЫТЫ (обзор)
   ========================================================================= */
{
  const s = pres.addSlide();
  header(s, "Результаты опытов", "Полевые опыты: прибавка урожайности");

  const cy = 1.95, ch = 4.55, cw = (PW - 2 * M - 0.4) / 2;
  const regions = [
    { map: "r_map_pukhalsk.png", stat: "+19%", crop: "Кормовая кукуруза",
      loc: "Акмолинская обл., с. Пухальское",
      lines: [
        "ОМУ NPK 12-12-12 (80 кг/га) — 13,8 т/га против ЖКУ, при внесении один раз вместе с семенами.",
        "К ЖКУ + Карбамид прибавка +7% (13,8 против 12,9 т/га) — но без дополнительной техники и расходов.",
      ] },
    { map: "r_map_arykbalyk.png", stat: "+25%", crop: "Ячмень",
      loc: "СКО, с. Арыкбалык",
      lines: [
        "ОМУ NPK 12-12-12 и NP 8-21 (80 кг/га) — 2,82 т/га против 2,26 т/га контроля.",
        "Аммофос 10-46 и Яра Милла 9-12-25 (80 кг/га) дали лишь +2% к контролю (2,3 т/га).",
      ] },
  ];
  regions.forEach((r, i) => {
    const x = M + i * (cw + 0.4);
    card(s, x, cy, cw, ch, { r: 0.11, shadow: { opacity: 0.14 } });
    // map
    s.addImage({ path: A(r.map), x: x + 0.28, y: cy + 0.28, w: cw - 0.56, h: 1.7 });
    // stat + crop
    s.addShape("ellipse", { x: x + 0.28, y: cy + 2.16, w: 0.16, h: 0.16, fill: { color: C.green }, line: { type: "none" } });
    s.addText(r.loc, { x: x + 0.52, y: cy + 2.08, w: cw - 0.8, h: 0.3, fontFace: FONT, fontSize: 11, bold: true, color: C.mute, valign: "middle" });
    s.addText(r.stat, { x: x + 0.24, y: cy + 2.42, w: 2.2, h: 0.8, fontFace: FONTH, fontSize: 46, bold: true, color: C.green, margin: 0 });
    s.addText([{ text: r.crop, options: { bold: true, color: C.ink } }, { text: "  ·  к урожайности", options: { color: C.body } }],
      { x: x + 0.28, y: cy + 3.24, w: cw - 0.56, h: 0.34, fontFace: FONTH, fontSize: 15, align: "left", margin: 0 });
    s.addText(r.lines.map((t, k) => ({ text: t, options: { bullet: { code: "2022", indent: 12 }, breakLine: true, paraSpaceAfter: 6, color: C.body } })),
      { x: x + 0.3, y: cy + 3.66, w: cw - 0.6, h: 0.8, fontFace: FONT, fontSize: 11.5, lineSpacing: 15, align: "left", valign: "top" });
  });

  footer(s, 5);
}

/* =========================================================================
   SLIDE 6 — ЯЧМЕНЬ: ЭКОНОМИКА (таблица)
   ========================================================================= */
{
  const s = pres.addSlide();
  header(s, "Результаты опытов", "Яровой ячмень: экономика внесения", { ts: 28 });
  s.addText("Посев 01.06.2023 · сорт Крешендо · густота 3,5 млн/га · анализ по ПО OneSoil / FieldView (зоны однородности поля за 6 лет)", {
    x: M, y: 1.42, w: PW - 2 * M, h: 0.3, fontFace: FONT, fontSize: 11.5, italic: true, color: C.mute });

  const head = ["Вариант", "Цена\n$/т", "Норма\nкг/га", "Доп.\n$/га", "Урож.\nц/га", "Δ ц", "Δ %", "Эконом.\n$/га", "FieldView\nц/га"];
  const rows = [
    ["Без удобрений (контроль)", "—", "—", "—", "32,7", "—", "—", "—", "22,6", 0],
    ["Аммофос NP 10-43", "413", "80", "33", "29,77", "−2,93", "−9,0%", "−78", "23,0", 0],
    ["Аммиачная селитра N 34", "381", "80", "30", "29,1", "−3,60", "−11,0%", "−85", "25,4", 0],
    ["Яра Мила NPK 9-12-25", "1 152", "80", "92", "34,8", "2,07", "6,3%", "−61", "23,1", 0],
    ["ОМУ NPK 12-12-12", "652", "80", "52", "29,9", "−2,83", "−8,7%", "−95", "28,2", 1],
    ["ОМУ NP 8-21", "652", "80", "52", "30,6", "−2,07", "−6,3%", "−84", "28,1", 1],
    ["План стар NP 10-46", "3 478", "80", "278", "32,1", "2,37", "−1,7%", "−242", "18,0", 0],
  ];
  const colW = [3.05, 0.92, 0.92, 0.92, 0.95, 0.9, 0.95, 1.05, 1.44];
  const tblW = colW.reduce((a, b) => a + b, 0);
  const tx = M, ty = 1.9;
  const tableRows = [];
  tableRows.push(head.map((h, i) => ({
    text: h, options: { fill: { color: C.deep }, color: C.white, bold: true, fontSize: 10.5,
      align: i === 0 ? "left" : "center", valign: "middle", fontFace: FONTH, margin: [3, 4, 3, 6] },
  })));
  rows.forEach((r) => {
    const hi = r[9] === 1;
    for (let i = 0; i < 9; i++) {
      const isNum = i >= 1;
      tableRows.push; // noop
    }
    tableRows.push(r.slice(0, 9).map((cell, i) => {
      let col = C.body;
      if (i === 8) col = C.greenD; // fieldview
      const emph = hi && (i === 0 || i === 8);
      return {
        text: cell, options: {
          fill: { color: hi ? C.tint : C.white }, color: emph ? C.greenD : col,
          bold: (i === 0) || (i === 8), fontSize: i === 0 ? 11 : 10.5,
          align: i === 0 ? "left" : "center", valign: "middle", fontFace: FONT,
          margin: [2, 4, 2, i === 0 ? 6 : 4],
        },
      };
    }));
  });
  s.addTable(tableRows, {
    x: tx, y: ty, w: tblW, colW, border: { type: "solid", color: C.line, pt: 1 },
    rowH: [0.5, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42], autoPage: false,
  });

  // callout note
  const noteY = ty + 0.5 + 7 * 0.42 + 0.18;
  card(s, tx, noteY, tblW, 0.86, { fill: C.tint, r: 0.09, noShadow: true });
  iconCircle(s, tx + 0.22, noteY + 0.19, 0.48, "microscope", { fill: C.green, variant: "white", pad: 0.26 });
  s.addText([
    { text: "По методике FieldView ", options: { bold: true, color: C.greenD } },
    { text: "органоминеральные удобрения показали максимальную урожайность — до ", options: { color: C.body } },
    { text: "28,2 ц/га", options: { bold: true, color: C.ink } },
    { text: " (против 22,6 ц/га на контроле). Результаты двух методик оказались противоречивыми — приведены обе оценки.", options: { color: C.body } },
  ], { x: tx + 0.86, y: noteY + 0.1, w: tblW - 1.1, h: 0.66, fontFace: FONT, fontSize: 11, lineSpacing: 14, valign: "middle", align: "left" });

  footer(s, 6);
}

/* =========================================================================
   SLIDE 7 — ИССЛЕДОВАНИЯ НПЦЗХ им. БАРАЕВА
   ========================================================================= */
{
  const s = pres.addSlide();
  header(s, "Независимые исследования", "НПЦ зернового хозяйства им. А.И. Бараева", { ts: 28 });

  // left column: findings on spring wheat
  const lx = M, lw = 6.35;
  s.addText("Яровая пшеница · заключение из отчёта", { x: lx, y: 1.65, w: lw, h: 0.3, fontFace: FONTH,
    fontSize: 12.5, bold: true, color: C.greenD });

  // three stat chips
  const chips = [["+21%", "NPK 12-12-12\nдоза 200 кг/га"], ["+22%", "NP 8-21-15SO₃\nN8P20"], ["1 класс", "качество\nзерна пшеницы"]];
  const chw = (lw - 0.4) / 3;
  chips.forEach((c, i) => {
    const x = lx + i * (chw + 0.2);
    card(s, x, 2.05, chw, 1.2, { fill: C.tint, r: 0.1, noShadow: true });
    s.addText(c[0], { x: x + 0.1, y: 2.16, w: chw - 0.2, h: 0.55, fontFace: FONTH, fontSize: 26, bold: true,
      color: C.green, align: "center", margin: 0 });
    s.addText(c[1], { x: x + 0.1, y: 2.72, w: chw - 0.2, h: 0.44, fontFace: FONT, fontSize: 10.5, color: C.body,
      align: "center", margin: 0, lineSpacing: 12 });
  });

  bullets(s, [
    "Контроль — 9,6 ц/га; аммофос (N4P20) достоверно прибавлял +1,6 ц/га.",
    "NPK 12-12-12 при 100 кг/га — +11% к контролю; при 200 кг/га — +21%.",
    "NP 8-21-15SO₃ (N8P20) дал +2,1 ц/га зерна — +22% к контролю.",
    "Клейковина выросла до 28,7–30,4% и 74–75 ед. ИДК (контроль — 25,5% и 64 ед.).",
  ], lx, 3.45, lw, { fs: 12.5, gap: 8 });

  // recommendation badge
  s.addShape("roundRect", { x: lx, y: 5.85, w: lw, h: 0.72, rectRadius: 0.1, fill: { color: C.deep }, line: { type: "none" },
    shadow: shadow({ opacity: 0.2, blur: 8, offset: 3 }) });
  iconCircle(s, lx + 0.2, 5.99, 0.44, "award", { fill: C.amber, variant: "white", pad: 0.26 });
  s.addText([
    { text: "НПЦЗХ им. Бараева рекомендует  ", options: { bold: true, color: C.white } },
    { text: "NP 8-21-15SO₃ и NPK 12-12-12", options: { bold: true, color: C.amber } },
  ], { x: lx + 0.78, y: 5.85, w: lw - 0.95, h: 0.72, fontFace: FONT, fontSize: 12.5, valign: "middle", align: "left" });

  // right column: our comment card
  const rx = 7.35, rw = PW - rx - M;
  card(s, rx, 1.65, rw, 4.92, { fill: C.tint, r: 0.11, noShadow: true });
  s.addText("Наш комментарий", { x: rx + 0.3, y: 1.85, w: rw - 0.6, h: 0.35, fontFace: FONTH, fontSize: 14, bold: true, color: C.greenD });
  const comm = [
    "ОМУ NPK 12-12-12 и NP 8-21 при дозах 100 и 200 кг/га сопоставимы по урожайности с аммофосом.",
    "В долгосрочной перспективе ОМУ восстанавливают почву, препятствуют деградации и насыщают её органикой.",
    "Не нужно вносить дополнительные удобрения — фермер экономит на технике и персонале.",
    "Стоимость ОМУ значительно ниже: 40–60% гранулы — органика, доход от которой уже получен в виде электроэнергии.",
  ];
  s.addText(comm.map((t) => ({ text: t, options: { bullet: { code: "2022", indent: 13 }, breakLine: true, paraSpaceAfter: 12, color: C.body } })),
    { x: rx + 0.32, y: 2.32, w: rw - 0.62, h: 4.1, fontFace: FONT, fontSize: 12.5, lineSpacing: 17, align: "left", valign: "top" });

  footer(s, 7);
}

/* =========================================================================
   SLIDE 8 — ПРОДУКЦИЯ
   ========================================================================= */
{
  const s = pres.addSlide();
  header(s, "Продукция", "Органоминеральные удобрения · от 75 000 т/год");

  const cy = 1.9, ch = 4.75, cw = (PW - 2 * M - 0.4) / 2;
  const prods = [
    { name: "ОМУ NPK 12-12-12", bag: "p_bag_npk.png", accent: C.green,
      rows: [["Органическое вещество", "40%"], ["Общий азот (N)", "12%"], ["  · аммонийный / мочевинный / орг.", "5 / 5 / 2"],
             ["Фосфор (P₂O₅), водораств.", "12 / 10%"], ["Калий (K₂O), водораств.", "12%"],
             ["Гуминовая + фульвовая к-та", "15%"], ["Влажность / pH", "≤20% / 5–7"]] },
    { name: "ОМУ NP 8-21", bag: "p_bag_np.png", accent: C.greenD,
      rows: [["Органическое вещество", "40%"], ["Общий азот (N)", "8%"], ["  · аммонийный / мочевинный / орг.", "3 / 3 / 2"],
             ["Фосфор (P₂O₅), водораств.", "21 / 19%"], ["Калий (K₂O)", "—"],
             ["Гуминовая + фульвовая к-та", "15%"], ["Влажность / pH", "≤20% / 5–7"]] },
  ];
  prods.forEach((p, i) => {
    const x = M + i * (cw + 0.4);
    card(s, x, cy, cw, ch, { r: 0.11, shadow: { opacity: 0.14 } });
    // bag image
    s.addImage({ path: A(p.bag), x: x + 0.28, y: cy + 0.3, w: 1.55, h: 2.1 });
    // name + tag
    s.addText(p.name, { x: x + 2.0, y: cy + 0.42, w: cw - 2.3, h: 0.5, fontFace: FONTH, fontSize: 19, bold: true, color: C.ink });
    s.addShape("roundRect", { x: x + 2.0, y: cy + 0.98, w: 2.35, h: 0.42, rectRadius: 0.21, fill: { color: p.accent }, line: { type: "none" } });
    s.addText("Гарантированный состав", { x: x + 2.0, y: cy + 0.98, w: 2.35, h: 0.42, fontFace: FONT, fontSize: 11, bold: true,
      color: C.white, align: "center", valign: "middle", margin: 0 });
    s.addText("% (w/w)", { x: x + 2.0, y: cy + 1.5, w: cw - 2.3, h: 0.3, fontFace: FONT, fontSize: 10.5, italic: true, color: C.mute });

    // spec table (below bag, full width)
    const trows = p.rows.map((r, k) => ([
      { text: r[0], options: { fill: { color: k % 2 ? C.white : C.tint }, color: C.body, fontSize: 11, align: "left",
        valign: "middle", fontFace: FONT, margin: [2, 6, 2, 6], bold: r[0].startsWith("  ") ? false : false } },
      { text: r[1], options: { fill: { color: k % 2 ? C.white : C.tint }, color: C.ink, bold: true, fontSize: 11,
        align: "right", valign: "middle", fontFace: FONT, margin: [2, 8, 2, 4] } },
    ]));
    s.addTable(trows, { x: x + 0.28, y: cy + 2.62, w: cw - 0.56, colW: [(cw - 0.56) * 0.68, (cw - 0.56) * 0.32],
      border: { type: "solid", color: C.line, pt: 0.5 }, rowH: 0.28, autoPage: false });
  });

  footer(s, 8);
}

/* =========================================================================
   SLIDE 9 — ПРЕИМУЩЕСТВА ОМУ
   ========================================================================= */
{
  const s = pres.addSlide();
  header(s, "Почему ОМУ", "Преимущества органоминеральных удобрений");

  const tiles = [
    ["chart",   "Прирост урожайности от 20%", "Стабильная прибавка на всех испытанных культурах."],
    ["leaf",    "До 40% органики", "Естественное содержание гуминовой и фульвовой кислот."],
    ["drop",    "Удержание влаги", "Губчатая структура органической части гранулы."],
    ["bolt",    "NPK при посеве", "Быстрорастворимая гранула — все элементы одномоментно."],
    ["recycle", "Восстановление почвы", "+2% плодородного слоя в год при ежегодном применении."],
    ["layer",   "Отказ от севооборота", "Богатая органика поддерживает плодородие поля."],
    ["flask",   "Любая формула NPK", "Под запрос и результаты анализа почвы."],
    ["check",   "Чистая гранула", "Без патогенов, грибов и семян сорных трав."],
  ];
  const cols = 4, rows = 2, gx = 0.28, gy = 0.3;
  const gw = (PW - 2 * M - (cols - 1) * gx) / cols;
  const gh = 2.0, y0 = 2.0;
  tiles.forEach((t, i) => {
    const cx = M + (i % cols) * (gw + gx);
    const cyy = y0 + Math.floor(i / cols) * (gh + gy);
    card(s, cx, cyy, gw, gh, { r: 0.1, shadow: { opacity: 0.11 } });
    iconCircle(s, cx + 0.28, cyy + 0.28, 0.68, t[0], { fill: C.tint, variant: "green" });
    s.addText(t[1], { x: cx + 0.26, y: cyy + 1.02, w: gw - 0.5, h: 0.5, fontFace: FONTH, fontSize: 13.5, bold: true,
      color: C.ink, lineSpacing: 16 });
    s.addText(t[2], { x: cx + 0.26, y: cyy + 1.44, w: gw - 0.5, h: 0.5, fontFace: FONT, fontSize: 10.8, color: C.body, lineSpacing: 13.5 });
  });

  footer(s, 9);
}

/* =========================================================================
   SLIDE 10 — МОДЕЛИ ДИСТРИБУЦИИ
   ========================================================================= */
{
  const s = pres.addSlide();
  header(s, "Каналы продаж", "Модели работы с дистрибьюторами", { ts: 28 });

  // assumptions strip (EXW definition folded in — removes need for a bottom legend)
  s.addText([
    { text: "Базовые допущения:  ", options: { bold: true, color: C.greenD } },
    { text: "весь РК + экспорт  ·  EXW — самовывоз со склада Kazbiofert  ·  цена «по рынку» (окно фиксации 3–7 дней, матрица скидок и ретро-бонусов)", options: { color: C.body } },
  ], { x: M, y: 1.42, w: PW - 2 * M, h: 0.32, fontFace: FONT, fontSize: 12, align: "left" });

  const colTitles = ["Объём и охват", "Запас, оборот, экспорт", "Ключевые клиенты"];
  const colIcons = ["boxes", "warehouse", "user"];
  const models = [
    [ // col 1
      ["1", "Buy–Sell (неэксклюзив)", "Дистрибьютор покупает у Kazbiofert и перепродаёт (EXW). Риск дебиторки и демпинга — на нём; управление матрицей скидок."],
      ["2", "Эксклюзив за KPI", "Эксклюзив по макрорегиону только при выполнении плана, покрытия и демо-полей. Не выполнил — делится/снимается."],
      ["3", "Мастер-дистриб + дилеры", "1–2 «якоря» масштабируют дилерскую сеть по РК; прямые Key Accounts остаются за Kazbiofert."],
    ],
    [ // col 2
      ["6", "Консигнация / ответхранение", "Товар на складе партнёра, право собственности у Kazbiofert до реализации. Оплата по отчёту продаж (7–14 дней)."],
      ["7", "Экспортный трейдер", "Продажа EXW трейдеру; он ведёт экспорт, логистику и страновые риски. Быстрый выход на рынки с сертификацией."],
    ],
    [ // col 3
      ["4", "Агент / комиссия", "Партнёр приводит клиента (контракт у Kazbiofert) и получает % от оплаты. Контроль клиентской базы сохраняется."],
      ["5", "Key Accounts + сервис", "Агрохолдинги и теплицы — прямой контракт (EXW). Партнёр зарабатывает на логистике, складе и агросопровождении."],
    ],
  ];
  const colW = (PW - 2 * M - 2 * 0.3) / 3;
  const colX = [M, M + colW + 0.3, M + 2 * (colW + 0.3)];
  const topY = 1.94;
  const cardsTop = topY + 0.66;
  models.forEach((col, ci) => {
    const x = colX[ci];
    // column header
    iconCircle(s, x, topY, 0.5, colIcons[ci], { fill: C.green, variant: "white", pad: 0.27 });
    s.addText(colTitles[ci], { x: x + 0.62, y: topY, w: colW - 0.62, h: 0.5, fontFace: FONTH, fontSize: 13.5, bold: true,
      color: C.greenD, valign: "middle" });
    let yy = cardsTop;
    const isThree = col.length === 3;
    const h = isThree ? 1.30 : 2.02;
    const gap = 0.15;
    col.forEach((m) => {
      card(s, x, yy, colW, h, { r: 0.09, shadow: { opacity: 0.1 } });
      // number badge
      s.addShape("ellipse", { x: x + 0.2, y: yy + 0.2, w: 0.42, h: 0.42, fill: { color: C.tint }, line: { type: "none" } });
      s.addText(m[0], { x: x + 0.2, y: yy + 0.2, w: 0.42, h: 0.42, fontFace: FONTH, fontSize: 15, bold: true, color: C.greenD,
        align: "center", valign: "middle", margin: 0 });
      s.addText(m[1], { x: x + 0.74, y: yy + 0.16, w: colW - 0.9, h: 0.5, fontFace: FONTH, fontSize: 12.5, bold: true,
        color: C.ink, valign: "middle", lineSpacing: 14 });
      s.addText(m[2], { x: x + 0.24, y: yy + 0.66, w: colW - 0.46, h: h - 0.78, fontFace: FONT, fontSize: 10.5, color: C.body,
        lineSpacing: 13, align: "left", valign: "top" });
      yy += h + gap;
    });
  });

  footer(s, 10);
}

/* =========================================================================
   SLIDE 11 — CLOSING
   ========================================================================= */
{
  const s = pres.addSlide();
  s.background = { path: A("bg_closing.jpg") };
  const lx = M + 0.15;
  s.addText("KAZBIOFERT", { x: lx, y: 1.7, w: 7, h: 0.5, fontFace: FONTH, fontSize: 16, bold: true, color: C.green, charSpacing: 4 });
  s.addText("Органоминеральные удобрения\nнового поколения", {
    x: lx, y: 2.25, w: 7.6, h: 1.7, fontFace: FONTH, fontSize: 38, bold: true, color: C.white, lineSpacing: 44 });
  s.addShape("line", { x: lx + 0.02, y: 4.05, w: 2.0, h: 0, line: { color: C.green, width: 2.5 } });

  const kpis = [["75 000 т", "ОМУ в год"], ["от +20%", "к урожайности"], ["1 : 3", "к минеральным"]];
  kpis.forEach((k, i) => {
    const x = lx + i * 2.55;
    s.addText(k[0], { x, y: 4.4, w: 2.4, h: 0.7, fontFace: FONTH, fontSize: 32, bold: true, color: C.amber, margin: 0 });
    s.addText(k[1], { x, y: 5.08, w: 2.4, h: 0.4, fontFace: FONT, fontSize: 13, color: "CFE3D6", margin: 0 });
  });
  s.addText("Биогазовый завод с интегрированным производством удобрений  ·  восстановление почвы  ·  экспортный потенциал", {
    x: lx, y: 5.95, w: 7.2, h: 0.7, fontFace: FONT, fontSize: 12.5, italic: true, color: "AFC9BA", lineSpacing: 17 });
}

} // end buildAll

module.exports = { buildAll, C, PW, PH, M, A, IC, FONT, FONTH };
