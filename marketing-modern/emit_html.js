// HTML preview recorder: implements the subset of the pptxgenjs slide API used
// by build_lib.js, and renders each slide as an absolutely-positioned HTML block
// at 96px/inch. Font falls back to Liberation Sans (WIDER than Calibri) so any
// text that fits here is guaranteed to fit in real PowerPoint — conservative
// overflow QA. Used only for visual inspection, never shipped.
const fs = require("fs");
const path = require("path");
const { buildAll, PW, PH } = require("./build_lib.js");

const PXI = 96;                 // px per inch
const in2px = (v) => v * PXI;
const pt2px = (v) => v * (96 / 72);
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const col = (c) => (c && c[0] === "#" ? c : "#" + (c || "000000"));
const fileUrl = (p) => "file://" + p;

function fontFamily(face) {
  // map Calibri -> metric-ish web fallbacks (wider => conservative)
  return `'Carlito','${face || "Calibri"}','Liberation Sans','DejaVu Sans',Arial,sans-serif`;
}

function alignItems(valign) {
  return valign === "middle" ? "center" : valign === "bottom" ? "flex-end" : "flex-start";
}

// build runs -> html, handling {text, options} arrays and plain strings + bullets
function runsHtml(text, base) {
  let items;
  if (Array.isArray(text)) items = text;
  else items = [{ text, options: {} }];

  // detect bullet mode (paragraph list): each item with bullet becomes a line
  const anyBullet = items.some((r) => r.options && r.options.bullet);
  const out = [];
  let curLine = [];
  const flush = (bulleted, spaceAfter) => {
    if (!curLine.length) return;
    const style = bulleted
      ? `display:flex;gap:6px;margin-bottom:${spaceAfter || 0}px;`
      : `margin-bottom:${spaceAfter || 0}px;`;
    const dot = bulleted ? `<span style="color:${bulleted};flex:0 0 auto;">&#8226;</span>` : "";
    out.push(`<div style="${style}">${dot}<span>${curLine.join("")}</span></div>`);
    curLine = [];
  };

  items.forEach((r) => {
    const o = r.options || {};
    const parts = String(r.text).split("\n");
    parts.forEach((seg, i) => {
      let s = esc(seg);
      const st = [];
      if (o.color) st.push(`color:${col(o.color)}`);
      if (o.bold) st.push("font-weight:700");
      if (o.italic) st.push("font-style:italic");
      if (o.fontSize) st.push(`font-size:${pt2px(o.fontSize)}px`);
      if (o.charSpacing) st.push(`letter-spacing:${o.charSpacing * 0.6}px`);
      curLine.push(`<span style="${st.join(";")}">${s}</span>`);
      if (i < parts.length - 1) { // hard newline
        flush(o.bullet ? "inherit" : false, 0);
      }
    });
    if (o.breakLine) {
      const bulletColor = o.bullet ? (base.bulletColor || "#2D9C5A") : false;
      flush(bulletColor, o.paraSpaceAfter ? o.paraSpaceAfter : 0);
    }
  });
  // final line
  const lastBullet = items[items.length - 1] && items[items.length - 1].options && items[items.length - 1].options.bullet;
  flush(lastBullet ? (base.bulletColor || "#2D9C5A") : false, 0);
  return out.join("");
}

class Slide {
  constructor() { this.parts = []; this.bg = "#F7FAF6"; }
  set background(v) {
    if (v.color) this.bg = col(v.color);
    else if (v.path) this.bg = `url('${fileUrl(v.path)}') center/cover no-repeat`;
  }
  addText(text, o = {}) {
    const x = in2px(o.x), y = in2px(o.y), w = in2px(o.w), h = in2px(o.h || 0.4);
    const st = [
      `position:absolute`, `left:${x}px`, `top:${y}px`, `width:${w}px`, `height:${h}px`,
      `font-family:${fontFamily(o.fontFace)}`,
      `font-size:${pt2px(o.fontSize || 14)}px`,
      `color:${col(o.color || "15211B")}`,
      `font-weight:${o.bold ? 700 : 400}`,
      o.italic ? "font-style:italic" : "",
      `text-align:${o.align || "left"}`,
      `display:flex`, `flex-direction:column`,
      `justify-content:${alignItems(o.valign || "top")}`,
      o.charSpacing ? `letter-spacing:${o.charSpacing * 0.6}px` : "",
      o.lineSpacing ? `line-height:${pt2px(o.lineSpacing)}px` : "line-height:1.15",
      `overflow:visible`, `box-sizing:border-box`,
      // margin: pptx default ~0.05in text inset unless margin:0
      o.margin === 0 ? "padding:0" : "padding:2px 4px",
      o.fill ? `background:${col(o.fill.color)}` : "",
      o.rectRadius ? `border-radius:${in2px(o.rectRadius)}px` : "",
    ].filter(Boolean).join(";");
    const align = o.align || "left";
    const inner = runsHtml(text, { bulletColor: col(o.color || "#2D9C5A") });
    // wrap: apply text-align inside
    this.parts.push(`<div style="${st};text-align:${align}"><div style="width:100%">${inner}</div></div>`);
  }
  addShape(type, o = {}) {
    const x = in2px(o.x), y = in2px(o.y), w = in2px(o.w), h = in2px(o.h);
    if (type === "line") {
      const color = col((o.line && o.line.color) || "#000");
      const tw = (o.line && o.line.width) || 1;
      this.parts.push(`<div style="position:absolute;left:${x}px;top:${y}px;width:${Math.max(w, tw)}px;height:${Math.max(h, tw)}px;background:${color}"></div>`);
      return;
    }
    const st = [
      `position:absolute`, `left:${x}px`, `top:${y}px`, `width:${w}px`, `height:${h}px`,
      o.fill && o.fill.color ? `background:${col(o.fill.color)}` : "background:transparent",
      o.line && o.line.color ? `border:${(o.line.width || 1)}px solid ${col(o.line.color)}` : "",
      type === "ellipse" ? "border-radius:50%" : (o.rectRadius ? `border-radius:${in2px(o.rectRadius)}px` : ""),
      o.shadow ? "box-shadow:0 3px 9px rgba(20,50,30,0.18)" : "",
      "box-sizing:border-box",
    ].filter(Boolean).join(";");
    this.parts.push(`<div style="${st}"></div>`);
  }
  addImage(o = {}) {
    const x = in2px(o.x), y = in2px(o.y), w = in2px(o.w), h = in2px(o.h);
    const rad = o.rounding ? "border-radius:50%;" : "";
    const sh = o.shadow ? "box-shadow:0 6px 14px rgba(0,0,0,0.28);" : "";
    this.parts.push(`<img src="${fileUrl(o.path)}" style="position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;object-fit:contain;${rad}${sh}"/>`);
  }
  addTable(rows, o = {}) {
    const x = in2px(o.x), y = in2px(o.y), w = in2px(o.w);
    const colW = (o.colW || []).map(in2px);
    const rowH = Array.isArray(o.rowH) ? o.rowH.map(in2px) : null;
    const defRowH = typeof o.rowH === "number" ? in2px(o.rowH) : 26;
    let html = `<table style="position:absolute;left:${x}px;top:${y}px;width:${w}px;border-collapse:collapse;table-layout:fixed">`;
    if (colW.length) {
      html += "<colgroup>" + colW.map((cw) => `<col style="width:${cw}px"/>`).join("") + "</colgroup>";
    }
    rows.forEach((r, ri) => {
      const rh = rowH ? (rowH[ri] || defRowH) : defRowH;
      html += `<tr style="height:${rh}px">`;
      r.forEach((cell) => {
        const co = cell.options || {};
        const st = [
          co.fill ? `background:${col(co.fill.color)}` : "",
          `color:${col(co.color || "#15211B")}`,
          co.bold ? "font-weight:700" : "font-weight:400",
          `font-size:${pt2px(co.fontSize || 11)}px`,
          `text-align:${co.align || "left"}`,
          `vertical-align:${co.valign === "middle" ? "middle" : "top"}`,
          `font-family:${fontFamily(co.fontFace)}`,
          "border:1px solid #D9E4D8",
          "padding:3px 6px", "overflow:hidden", "white-space:normal", "word-break:break-word",
          "line-height:1.1",
        ].filter(Boolean).join(";");
        html += `<td style="${st}">${esc(cell.text).replace(/\n/g, "<br>")}</td>`;
      });
      html += "</tr>";
    });
    html += "</table>";
    this.parts.push(html);
  }
  html() {
    return `<section class="slide" style="background:${this.bg}">${this.parts.join("\n")}</section>`;
  }
}

class Pres {
  constructor() { this.slides = []; }
  addSlide() { const s = new Slide(); this.slides.push(s); return s; }
  defineSlideMaster() {}
  set layout(v) {}
}

const pres = new Pres();
buildAll(pres);

const W = in2px(PW), H = in2px(PH);
const page = `<!doctype html><html><head><meta charset="utf-8"><style>
  *{margin:0;box-sizing:border-box}
  body{background:#333;padding:0}
  .slide{position:relative;width:${W}px;height:${H}px;overflow:hidden;margin:0 auto 2px;}
</style></head><body>
${pres.slides.map((s) => s.html()).join("\n")}
</body></html>`;

fs.writeFileSync(path.join(__dirname, "preview.html"), page);
console.log("HTML slides:", pres.slides.length);
