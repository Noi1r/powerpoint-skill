# Color Themes & Visual System

## Theme Definitions

**Academic Light** (default):
```json
{"name": "academic_light", "background": {"title": "1E2761", "content": "F8F9FA", "section": "1E2761"}, "text": {"primary": "1A1A2E", "secondary": "4A4A5A", "muted": "8C8C8C"}, "accent": {"positive": "0173B2", "negative": "DE8F05", "emphasis": "029E73"}, "formula": {"fg": "1A1A2E"}, "chart": ["0173B2", "DE8F05", "029E73", "CC78BC"], "card_fill": "0173B215", "divider": "D0D0D0"}
```

**Midnight**:
```json
{"name": "midnight", "background": {"title": "16213E", "content": "1A1A2E", "section": "16213E"}, "text": {"primary": "EAEAEA", "secondary": "B0B0C0", "muted": "707080"}, "accent": {"positive": "4FC3F7", "negative": "FFB74D", "emphasis": "81C784"}, "formula": {"fg": "FFFFFF"}, "chart": ["4FC3F7", "FFB74D", "81C784", "CE93D8"], "card_fill": "4FC3F715", "divider": "3A3A5E"}
```

**Ocean**:
```json
{"name": "ocean", "background": {"title": "065A82", "content": "FFFFFF", "section": "065A82"}, "text": {"primary": "1A1A2E", "secondary": "3A5A6E", "muted": "8C8C8C"}, "accent": {"positive": "0173B2", "negative": "DE8F05", "emphasis": "029E73"}, "formula": {"fg": "1A1A2E"}, "chart": ["065A82", "1C7293", "029E73", "DE8F05"], "card_fill": "065A8215", "divider": "C8DDE8"}
```

**Forest**:
```json
{"name": "forest", "background": {"title": "2C5F2D", "content": "F5F5F0", "section": "2C5F2D"}, "text": {"primary": "2C3E2D", "secondary": "4A5A4A", "muted": "8C8C8C"}, "accent": {"positive": "0173B2", "negative": "DE8F05", "emphasis": "029E73"}, "formula": {"fg": "2C3E2D"}, "chart": ["2C5F2D", "97BC62", "0173B2", "DE8F05"], "card_fill": "2C5F2D15", "divider": "C8D8C0"}
```

**Sandwich** (dark title/conclusion + light content):
```json
{"name": "sandwich", "background": {"title": "2D3436", "content": "FFFFFF", "section": "2D3436"}, "text": {"primary": "2D3436", "secondary": "555555", "muted": "8C8C8C"}, "accent": {"positive": "0173B2", "negative": "DE8F05", "emphasis": "029E73"}, "formula": {"fg_light": "2D3436", "fg_dark": "FFFFFF"}, "chart": ["0173B2", "DE8F05", "029E73", "CC78BC"], "card_fill": "0173B215", "divider": "D0D0D0"}
```

## Slide Masters

4 masters per theme. Define via `pres.defineSlideMaster()`:

| Master | Background | Typography | Usage |
|--------|-----------|------------|-------|
| `TITLE_SLIDE` | `theme.background.title` (dark) | Title 44pt white, subtitle 18pt white 80% opacity | Cover slide |
| `SECTION_SLIDE` | `theme.background.section` (dark) | Section title 36pt white | Chapter transitions |
| `CONTENT_SLIDE` | `theme.background.content` (light) | Title bar 0.7" at top with dark bg, body below | All content slides |
| `THANK_YOU` | `theme.background.title` (dark) | "Thank You" 44pt white, contact 14pt | Closing slide |

Content slides use dynamic layouts built on `CONTENT_SLIDE` master for maximum flexibility.

## Typography

Default pairing: **Georgia** (titles) + **Calibri** (body).
Fallback chain: Georgia → Cambria → DejaVu Serif → system serif. Calibri → Segoe UI → DejaVu Sans → system sans.

Default baselines — use `F` constant in script template. Scale up 2-4pt when content is sparse (text fills <50% of card/region). Never go below 10pt.

| Role | Key | Font | Size | Weight |
|------|-----|------|------|--------|
| Cover title | `F.cover` | Georgia | 48pt | Bold |
| Section divider | `F.section` | Georgia | 36pt | Bold |
| Content title bar | `F.title` | Georgia | 22pt | Bold |
| Cover subtitle | `F.subtitle` | Calibri | 18pt | Regular |
| Body / bullets | `F.body` | Calibri | 15pt | Regular |
| Card body / small | `F.small` | Calibri | 12pt | Regular |
| Caption / footnote | `F.caption` | Calibri | 11pt | Regular |
| Stat callout number | `F.stat` | Georgia | 60pt | Bold |
| Card title | `F.cardTitle` | Georgia | 18pt | Bold |
| Table header | `F.tblHead` | Calibri | 13pt | Bold |
| Table cell | `F.tblCell` | Calibri | 12pt | Regular |

## Semantic Colors

| Role | Hex | Usage |
|------|-----|-------|
| positive | `0173B2` (blue) | Correct, advantage, positive result |
| negative | `DE8F05` (orange) | Limitation, drawback, negative result |
| emphasis | `029E73` (green) | Key finding, highlight |
| neutral | `8C8C8C` (gray) | De-emphasized, context |
| card-fill | Theme-derived, 10-15% opacity | Accent card/box background |
| divider | Theme-derived, lighter | Section dividers, horizontal rules |

In formulas: `\textcolor{positive}{...}`, `\textcolor{negative}{...}`, `\textcolor{emphasis}{...}`.

## Script Template (Academic Light)

```javascript
const pptxgen = require("pptxgenjs");
const fs = require("fs");
let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Presenter Name";
pres.title = "Presentation Title";

// Theme — short alias T
const T = {
  bg: { title: "1E2761", content: "F8F9FA", section: "1E2761" },
  tx: { pri: "1A1A2E", sec: "4A4A5A", mut: "8C8C8C", wh: "FFFFFF" },
  ac: { pos: "0173B2", neg: "DE8F05", emp: "029E73" },
  chart: ["0173B2", "DE8F05", "029E73", "CC78BC"],
  cardFill: "0173B215", divider: "D0D0D0"
};

// Typography — default baselines (scale up when content is sparse, see QA density check)
const F = {
  cover:    { face: "Georgia", size: 48 },
  section:  { face: "Georgia", size: 36 },
  title:    { face: "Georgia", size: 22 },
  subtitle: { face: "Calibri", size: 18 },
  body:     { face: "Calibri", size: 15 },
  small:    { face: "Calibri", size: 12 },
  caption:  { face: "Calibri", size: 11 },
  stat:     { face: "Georgia", size: 60 },
  cardTitle:{ face: "Georgia", size: 18 },
  tblHead:  { face: "Calibri", size: 13 },
  tblCell:  { face: "Calibri", size: 12 },
};

// Layout constants
const SW = 10, SH = 5.625, M = 0.5;
const CW = SW - 2 * M;       // content width (9")
const CY = 0.9;              // content top Y (below title bar)
const CH = SH - CY - M;      // content height (4.225")

// Slide Masters
pres.defineSlideMaster({
  title: "TITLE_SLIDE",
  background: { color: T.bg.title },
  objects: [
    { rect: { x: 0, y: 4.8, w: SW, h: 0.825, fill: { color: "000000", transparency: 30 } } }
  ]
});
pres.defineSlideMaster({
  title: "SECTION_SLIDE",
  background: { color: T.bg.section }
});
pres.defineSlideMaster({
  title: "CONTENT_SLIDE",
  background: { color: T.bg.content },
  objects: [
    { rect: { x: 0, y: 0, w: SW, h: 0.7, fill: { color: T.bg.title } } }
  ]
});
pres.defineSlideMaster({
  title: "THANK_YOU",
  background: { color: T.bg.title }
});

// Formula manifest — supports OMML (native math) and image (PNG/SVG) entries
const manifest = JSON.parse(fs.readFileSync("formulas/manifest.json", "utf8"));

// OMML placeholder — inject_omml.py replaces these post-generation
function addMathText(slide, id, x, y, w, h, opts = {}) {
  slide.addText(`{{MATH:${id}}}`, {
    x, y, w, h,
    fontFace: opts.fontFace || F.body.face,
    fontSize: opts.fontSize || F.body.size,
    color: opts.color || T.tx.pri,
    align: opts.align || "center",
    valign: opts.valign || "middle",
    margin: 0
  });
}

// Centered formula — auto-routes OMML vs image based on manifest
function addFormula(slide, id, y, targetH = 0.4) {
  const f = manifest[id];
  if (!f || f.error) return;
  if (f.render === "omml") {
    addMathText(slide, id, M, y, CW, targetH);
    return;
  }
  if (f.format === "svg") {
    const b64 = Buffer.from(fs.readFileSync(f.file, "utf8")).toString("base64");
    const scale = targetH / (f.height_pt / 72);
    const w = (f.width_pt / 72) * scale;
    slide.addImage({ data: "image/svg+xml;base64," + b64, x: (SW - w) / 2, y, w, h: targetH });
  } else {
    const dpi = f.dpi || 600;
    const scale = targetH / (f.height_px / dpi);
    const w = (f.width_px / dpi) * scale;
    slide.addImage({ path: f.file, x: (SW - w) / 2, y, w, h: targetH });
  }
}

// Left-aligned formula — auto-routes OMML vs image
function addFormulaAt(slide, id, x, y, maxW, targetH) {
  const f = manifest[id];
  if (!f || f.error) return;
  if (f.render === "omml") {
    addMathText(slide, id, x, y, maxW, targetH, { align: "left" });
    return;
  }
  if (f.format === "svg") {
    const b64 = Buffer.from(fs.readFileSync(f.file, "utf8")).toString("base64");
    let scale = targetH / (f.height_pt / 72);
    let w = (f.width_pt / 72) * scale, h = targetH;
    if (w > maxW) { const s = maxW / w; w = maxW; h *= s; }
    slide.addImage({ data: "image/svg+xml;base64," + b64, x, y, w, h });
  } else {
    const dpi = f.dpi || 600;
    let scale = targetH / (f.height_px / dpi);
    let w = (f.width_px / dpi) * scale, h = targetH;
    if (w > maxW) { const s = maxW / w; w = maxW; h *= s; }
    slide.addImage({ path: f.file, x, y, w, h });
  }
}

// Card with formula body — auto-routes OMML vs image
function addCardFormula(sl, x, y, w, h, color, title, formulaId, fH) {
  chk("cardFormula", y, h);
  sl.addShape(pres.shapes.RECTANGLE, { x, y, w, h,
    fill: { color: "FFFFFF" }, shadow: mkSh() });
  sl.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h,
    fill: { color } });
  const px = x + 0.25, pw = w - 0.4;
  sl.addText(title, { x: px, y: y + 0.1, w: pw, h: 0.4,
    fontFace: F.cardTitle.face, fontSize: F.cardTitle.size, bold: true,
    color: T.tx.pri, valign: "top", margin: [0, 0, 0, 0], shrinkText: true });
  addFormulaAt(sl, formulaId, px, y + 0.55, pw, fH || (h - 0.7));
}

// Factory: card shadow (NEVER reuse option objects directly)
const mkSh = () => ({ type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.15 });

// Diagram manifest (optional — only when diagrams/ exists)
const dManifest = (() => {
  try { return JSON.parse(fs.readFileSync("diagrams/manifest.json", "utf8")); }
  catch { return {}; }
})();

// Centered diagram
function addDiagram(slide, id, y, targetH, maxW = CW) {
  const d = dManifest[id];
  if (!d || d.error) return;
  let w, h = targetH;
  if (d.format === "svg") {
    const b64 = Buffer.from(fs.readFileSync(d.file, "utf8")).toString("base64");
    const scale = targetH / (d.height_pt / 72);
    w = (d.width_pt / 72) * scale;
    if (w > maxW) { const s = maxW / w; w = maxW; h *= s; }
    slide.addImage({ data: "image/svg+xml;base64," + b64, x: (SW - w) / 2, y, w, h });
  } else {
    const dpi = d.dpi || 300;
    const scale = targetH / (d.height_px / dpi);
    w = (d.width_px / dpi) * scale;
    if (w > maxW) { const s = maxW / w; w = maxW; h *= s; }
    slide.addImage({ path: d.file, x: (SW - w) / 2, y, w, h });
  }
}

// Positioned diagram (for two-column layouts)
function addDiagramAt(slide, id, x, y, maxW, targetH) {
  const d = dManifest[id];
  if (!d || d.error) return;
  let w, h = targetH;
  if (d.format === "svg") {
    const b64 = Buffer.from(fs.readFileSync(d.file, "utf8")).toString("base64");
    const scale = targetH / (d.height_pt / 72);
    w = (d.width_pt / 72) * scale;
    if (w > maxW) { const s = maxW / w; w = maxW; h *= s; }
    slide.addImage({ data: "image/svg+xml;base64," + b64, x, y, w, h });
  } else {
    const dpi = d.dpi || 300;
    const scale = targetH / (d.height_px / dpi);
    w = (d.width_px / dpi) * scale;
    if (w > maxW) { const s = maxW / w; w = maxW; h *= s; }
    slide.addImage({ path: d.file, x, y, w, h });
  }
}

// Extracted figure with caption (source attribution)
function addFigure(slide, id, y, targetH, caption) {
  addDiagram(slide, id, y, targetH - 0.3);
  if (caption) {
    slide.addText(caption, {
      x: M, y: y + targetH - 0.25, w: CW, h: 0.25,
      fontFace: F.caption.face, fontSize: F.caption.size,
      color: T.tx.mut, align: "center", italic: true, margin: 0
    });
  }
}

// Themed chart (auto-applies theme colors and clean styling)
function addChart(sl, type, data, x, y, w, h, opts = {}) {
  chk("chart", y, h);
  sl.addChart(type, data, {
    x, y, w, h,
    chartColors: opts.chartColors || T.chart,
    chartArea: { fill: { color: T.bg.content }, roundedCorners: true },
    catAxisLabelColor: T.tx.sec, valAxisLabelColor: T.tx.sec,
    catAxisLabelFontFace: F.tblCell.face, catAxisLabelFontSize: F.tblCell.size,
    valAxisLabelFontFace: F.tblCell.face, valAxisLabelFontSize: F.tblCell.size,
    valGridLine: { color: T.divider, size: 0.5 },
    catGridLine: { style: "none" },
    showLegend: opts.showLegend !== undefined ? opts.showLegend : data.length > 1,
    legendPos: opts.legendPos || "b",
    legendFontFace: F.caption.face, legendFontSize: F.caption.size, legendColor: T.tx.sec,
    ...opts,
  });
}

// --- Slide builders ---

function sTitle(sl, txt) {
  sl.addText(txt, { x: M, y: 0.08, w: CW, h: 0.55,
    fontFace: F.title.face, fontSize: F.title.size, bold: true,
    color: T.tx.wh, margin: [0, 6, 0, 0], shrinkText: true });
}

function sectionSlide(title) {
  const sl = pres.addSlide({ masterName: "SECTION_SLIDE" });
  sl.addText(title, { x: M, y: SH / 2 - 0.5, w: CW, h: 1,
    fontFace: F.section.face, fontSize: F.section.size, bold: true,
    color: T.tx.wh, align: "center", valign: "middle" });
  sl.addShape(pres.shapes.RECTANGLE, { x: SW / 2 - 1.5, y: SH / 2 + 0.4,
    w: 3, h: 0.04, fill: { color: T.ac.pos } });
  return sl;
}

// --- Content helpers ---

function addBullets(sl, items, x, y, w, h, opts = {}) {
  chk("bullets", y, h);
  const arr = items.map((t, i) => ({
    text: t, options: {
      bullet: true, breakLine: i < items.length - 1,
      fontSize: opts.fontSize || F.body.size,
      fontFace: opts.fontFace || F.body.face,
      color: opts.color || T.tx.pri,
    }
  }));
  sl.addText(arr, { x, y, w, h, valign: "top", margin: [4, 6, 4, 4], shrinkText: true });
}

function addCard(sl, x, y, w, h, color, title, body) {
  chk("card", y, h);
  sl.addShape(pres.shapes.RECTANGLE, { x, y, w, h,
    fill: { color: "FFFFFF" }, shadow: mkSh() });
  sl.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h,
    fill: { color } });
  // px: 0.19" clear of accent bar. pw: 0.15" from right edge.
  const px = x + 0.25, pw = w - 0.4;
  sl.addText(title, { x: px, y: y + 0.1, w: pw, h: 0.4,
    fontFace: F.cardTitle.face, fontSize: F.cardTitle.size, bold: true,
    color: T.tx.pri, valign: "top", margin: [0, 0, 0, 0], shrinkText: true });
  sl.addText(body, { x: px, y: y + 0.48, w: pw, h: h - 0.62,
    fontFace: F.body.face, fontSize: F.body.size,
    color: T.tx.sec, valign: "top", margin: [2, 0, 8, 0], shrinkText: true });
}

function addTable(sl, headers, rows, x, y, w, colW, rowH = 0.35) {
  chk("table", y, rowH * (rows.length + 1));
  const hRow = headers.map(h => ({ text: h, options: {
    fill: { color: T.bg.title }, color: T.tx.wh,
    bold: true, fontFace: F.tblHead.face, fontSize: F.tblHead.size,
    align: "center", valign: "middle"
  }}));
  const dRows = rows.map((row, i) => row.map(cell => {
    const isObj = typeof cell === "object" && cell !== null;
    const txt = isObj ? cell.text : String(cell);
    const base = {
      fill: { color: i % 2 === 0 ? "FFFFFF" : "F0F2F5" },
      fontFace: F.tblCell.face, fontSize: F.tblCell.size,
      color: T.tx.pri, valign: "middle"
    };
    return { text: txt, options: isObj ? { ...base, ...cell.options } : base };
  }));
  sl.addTable([hRow, ...dRows], { x, y, w, colW, border: { pt: 0.5, color: "E0E0E0" } });
}

// Continuation table page — repeats header, adds " (cont'd)" to slide title
function addTableCont(sl, title, headers, rows, x, y, w, colW, rowH = 0.35) {
  sTitle(sl, title + " (cont'd)");
  addTable(sl, headers, rows, x, y, w, colW, rowH);
}

function addNumCard(sl, num, title, body, x, y, w, h, color) {
  addCard(sl, x, y, w, h, color, title, body);
  sl.addShape(pres.shapes.OVAL, { x: x + w - 0.4, y: y - 0.15,
    w: 0.35, h: 0.35, fill: { color } });
  sl.addText(String(num), { x: x + w - 0.4, y: y - 0.15, w: 0.35, h: 0.35,
    fontFace: F.title.face, fontSize: 14, bold: true,
    color: "FFFFFF", align: "center", valign: "middle", margin: 0 });
}

function addFlow(sl, steps, y, h = 0.6) {
  const n = steps.length;
  const gap = 0.25, arrowW = 0.2;
  const totalArrows = (n - 1) * (gap + arrowW);
  const boxW = (CW - totalArrows) / n;
  let cx = M;
  steps.forEach((step, i) => {
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y, w: boxW, h,
      fill: { color: T.ac.pos, transparency: 88 },
      line: { color: T.ac.pos, width: 1.5 }, rectRadius: 0.08 });
    sl.addText(step, { x: cx, y, w: boxW, h,
      fontFace: F.small.face, fontSize: F.small.size,
      color: T.tx.pri, align: "center", valign: "middle", margin: 0 });
    if (i < n - 1) {
      const ax = cx + boxW + gap / 2;
      sl.addText("→", { x: ax, y, w: arrowW, h,
        fontFace: F.body.face, fontSize: 18, color: T.tx.mut,
        align: "center", valign: "middle", margin: 0 });
    }
    cx += boxW + gap + arrowW;
  });
}

function cols2() {
  const gap = 0.3, colW = (CW - gap) / 2;
  return {
    left:  { x: M, w: colW },
    right: { x: M + colW + gap, w: colW },
    colW, gap
  };
}

// Boundary check: warn if element bottom exceeds slide content area
function chk(label, y, h) {
  const bot = y + h;
  if (bot > SH - M + 0.05) console.warn(`OVERFLOW: ${label} bottom=${bot.toFixed(2)}" > limit=${(SH-M).toFixed(2)}"`);
}

// Gap check: verify element starts below previous element's bottom with minimum gap
let _lastBot = 0;
function gap(label, y, minGap = 0.08) {
  if (_lastBot > 0 && y < _lastBot + minGap)
    console.warn(`OVERLAP: ${label} y=${y.toFixed(2)}" < prev bottom + gap=${(_lastBot+minGap).toFixed(2)}"`);
}
function trackBot(y, h) { _lastBot = y + h; }
function resetBot() { _lastBot = 0; }

// === SLIDES ===
// ... generate all slides here ...

pres.writeFile({ fileName: "output.pptx" });
```

| Major | Diagram theme mismatch — Graphviz/Mermaid default colors in output | -8/instance |
| Major | Diagram labels unreadable or nodes overlapping | -8/instance |
| Major | Extracted figure without source attribution (R26) | -5/instance |
| Major | Chart colors don't match theme `T.chart[]` palette | -5/instance |
| Major | Chart axis labels unreadable or overlapping | -5/instance |
| Major | Table cell math as Unicode text instead of OMML (R30) | -5/instance |
| Major | Long table missing header repeat or "(cont'd)" on continuation page (R31) | -8/table |
| Major | Table without caption or takeaway line below | -5/instance |

## Quality Scoring Rubric

Start at 100, deduct for issues found:

| Severity | Issue | Deduction |
|----------|-------|-----------|
| Critical | JS script execution failure | -100 |
| Critical | Formula image missing or corrupted | -20/slide (cap: -60) |
| Critical | Element overlap (card on table, title on body, etc.) | -15/instance (cap: -45) |
| Critical | Text/card overflow beyond slide boundary (bottom-edge clipping) | -15/instance (cap: -45) |
| Major | Formula image blurry (insufficient DPI) | -10/slide |
| Major | 3+ consecutive slides with same layout | -10 |
| Major | Color inconsistent with theme | -8/instance |
| Major | Missing references slide | -5 |
| Major | Uneven spacing | -5/instance |
| Major | Visual blank space — rendered content occupies <75% of CH (R19) | -8/slide |
| Major | Sparse content — text fills <50% of card/region area | -5/slide |
| Minor | Tight margins (0.3"-0.5") | -2/slide |
| Minor | Font size inconsistency | -2/instance |

Score floor is 0. Per-category caps: formula issues -60, overlap -45, overflow -45.

**Thresholds**: ≥90 deliver, 80-89 acceptable with warnings, <80 must fix.
