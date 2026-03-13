# PptxGenJS Tutorial

## Setup & Basic Structure

```javascript
const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';  // or 'LAYOUT_16x10', 'LAYOUT_4x3', 'LAYOUT_WIDE'
pres.author = 'Your Name';
pres.title = 'Presentation Title';

let slide = pres.addSlide();
slide.addText("Hello World!", { x: 0.5, y: 0.5, fontSize: 36, color: "363636" });

pres.writeFile({ fileName: "Presentation.pptx" });
```

---

## Layout Dimensions

Slide dimensions (coordinates in inches):
- `LAYOUT_16x9`: 10" × 5.625" (default)
- `LAYOUT_16x10`: 10" × 6.25"
- `LAYOUT_4x3`: 10" × 7.5"
- `LAYOUT_WIDE`: 13.3" × 7.5"

---

## Text & Formatting

```javascript
// Basic text
slide.addText("Simple Text", {
  x: 1, y: 1, w: 8, h: 2, fontSize: 24, fontFace: "Arial",
  color: "363636", bold: true, align: "center", valign: "middle"
});

// Character spacing (use charSpacing, not letterSpacing which is silently ignored)
slide.addText("SPACED TEXT", { x: 1, y: 1, w: 8, h: 1, charSpacing: 6 });

// Rich text arrays
slide.addText([
  { text: "Bold ", options: { bold: true } },
  { text: "Italic ", options: { italic: true } }
], { x: 1, y: 3, w: 8, h: 1 });

// Multi-line text (requires breakLine: true)
slide.addText([
  { text: "Line 1", options: { breakLine: true } },
  { text: "Line 2", options: { breakLine: true } },
  { text: "Line 3" }  // Last item doesn't need breakLine
], { x: 0.5, y: 0.5, w: 8, h: 2 });

// Text box margin (internal padding)
slide.addText("Title", {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  margin: 0  // Use 0 when aligning text with other elements like shapes or icons
});
```

**Tip:** Text boxes have internal margin by default. Set `margin: 0` when you need text to align precisely with shapes, lines, or icons at the same x-position.

---

## Lists & Bullets

```javascript
// ✅ CORRECT: Multiple bullets
slide.addText([
  { text: "First item", options: { bullet: true, breakLine: true } },
  { text: "Second item", options: { bullet: true, breakLine: true } },
  { text: "Third item", options: { bullet: true } }
], { x: 0.5, y: 0.5, w: 8, h: 3 });

// ❌ WRONG: Never use unicode bullets
slide.addText("• First item", { ... });  // Creates double bullets

// Sub-items and numbered lists
{ text: "Sub-item", options: { bullet: true, indentLevel: 1 } }
{ text: "First", options: { bullet: { type: "number" }, breakLine: true } }
```

---

## Shapes

```javascript
slide.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 0.8, w: 1.5, h: 3.0,
  fill: { color: "FF0000" }, line: { color: "000000", width: 2 }
});

slide.addShape(pres.shapes.OVAL, { x: 4, y: 1, w: 2, h: 2, fill: { color: "0000FF" } });

slide.addShape(pres.shapes.LINE, {
  x: 1, y: 3, w: 5, h: 0, line: { color: "FF0000", width: 3, dashType: "dash" }
});

// With transparency
slide.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "0088CC", transparency: 50 }
});

// Rounded rectangle (rectRadius only works with ROUNDED_RECTANGLE, not RECTANGLE)
// ⚠️ Don't pair with rectangular accent overlays — they won't cover rounded corners. Use RECTANGLE instead.
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "FFFFFF" }, rectRadius: 0.1
});

// With shadow
slide.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "FFFFFF" },
  shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.15 }
});
```

Shadow options:

| Property | Type | Range | Notes |
|----------|------|-------|-------|
| `type` | string | `"outer"`, `"inner"` | |
| `color` | string | 6-char hex (e.g. `"000000"`) | No `#` prefix, no 8-char hex — see Common Pitfalls |
| `blur` | number | 0-100 pt | |
| `offset` | number | 0-200 pt | **Must be non-negative** — negative values corrupt the file |
| `angle` | number | 0-359 degrees | Direction the shadow falls (135 = bottom-right, 270 = upward) |
| `opacity` | number | 0.0-1.0 | Use this for transparency, never encode in color string |

To cast a shadow upward (e.g. on a footer bar), use `angle: 270` with a positive offset — do **not** use a negative offset.

**Note**: Gradient fills are not natively supported. Use a gradient image as a background instead.

---

## Images

### Image Sources

```javascript
// From file path
slide.addImage({ path: "images/chart.png", x: 1, y: 1, w: 5, h: 3 });

// From URL
slide.addImage({ path: "https://example.com/image.jpg", x: 1, y: 1, w: 5, h: 3 });

// From base64 (faster, no file I/O)
slide.addImage({ data: "image/png;base64,iVBORw0KGgo...", x: 1, y: 1, w: 5, h: 3 });
```

### Image Options

```javascript
slide.addImage({
  path: "image.png",
  x: 1, y: 1, w: 5, h: 3,
  rotate: 45,              // 0-359 degrees
  rounding: true,          // Circular crop
  transparency: 50,        // 0-100
  flipH: true,             // Horizontal flip
  flipV: false,            // Vertical flip
  altText: "Description",  // Accessibility
  hyperlink: { url: "https://example.com" }
});
```

### Image Sizing Modes

```javascript
// Contain - fit inside, preserve ratio
{ sizing: { type: 'contain', w: 4, h: 3 } }

// Cover - fill area, preserve ratio (may crop)
{ sizing: { type: 'cover', w: 4, h: 3 } }

// Crop - cut specific portion
{ sizing: { type: 'crop', x: 0.5, y: 0.5, w: 2, h: 2 } }
```

### Calculate Dimensions (preserve aspect ratio)

```javascript
const origWidth = 1978, origHeight = 923, maxHeight = 3.0;
const calcWidth = maxHeight * (origWidth / origHeight);
const centerX = (10 - calcWidth) / 2;

slide.addImage({ path: "image.png", x: centerX, y: 1.2, w: calcWidth, h: maxHeight });
```

### Supported Formats

- **Standard**: PNG, JPG, GIF (animated GIFs work in Microsoft 365)
- **SVG**: Works in modern PowerPoint/Microsoft 365

---

## Icons

Use react-icons to generate SVG icons, then rasterize to PNG for universal compatibility.

### Setup

```javascript
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { FaCheckCircle, FaChartLine } = require("react-icons/fa");

function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}
```

### Add Icon to Slide

```javascript
const iconData = await iconToBase64Png(FaCheckCircle, "#4472C4", 256);

slide.addImage({
  data: iconData,
  x: 1, y: 1, w: 0.5, h: 0.5  // Size in inches
});
```

**Note**: Use size 256 or higher for crisp icons. The size parameter controls the rasterization resolution, not the display size on the slide (which is set by `w` and `h` in inches).

### Icon Libraries

Install: `npm install -g react-icons react react-dom sharp`

Popular icon sets in react-icons:
- `react-icons/fa` - Font Awesome
- `react-icons/md` - Material Design
- `react-icons/hi` - Heroicons
- `react-icons/bi` - Bootstrap Icons

---

## Slide Backgrounds

```javascript
// Solid color
slide.background = { color: "F1F1F1" };

// Color with transparency
slide.background = { color: "FF3399", transparency: 50 };

// Image from URL
slide.background = { path: "https://example.com/bg.jpg" };

// Image from base64
slide.background = { data: "image/png;base64,iVBORw0KGgo..." };
```

---

## Tables

```javascript
slide.addTable([
  ["Header 1", "Header 2"],
  ["Cell 1", "Cell 2"]
], {
  x: 1, y: 1, w: 8, h: 2,
  border: { pt: 1, color: "999999" }, fill: { color: "F1F1F1" }
});

// Advanced with merged cells
let tableData = [
  [{ text: "Header", options: { fill: { color: "6699CC" }, color: "FFFFFF", bold: true } }, "Cell"],
  [{ text: "Merged", options: { colspan: 2 } }]
];
slide.addTable(tableData, { x: 1, y: 3.5, w: 8, colW: [4, 4] });
```

---

## Charts

```javascript
// Bar chart
slide.addChart(pres.charts.BAR, [{
  name: "Sales", labels: ["Q1", "Q2", "Q3", "Q4"], values: [4500, 5500, 6200, 7100]
}], {
  x: 0.5, y: 0.6, w: 6, h: 3, barDir: 'col',
  showTitle: true, title: 'Quarterly Sales'
});

// Line chart
slide.addChart(pres.charts.LINE, [{
  name: "Temp", labels: ["Jan", "Feb", "Mar"], values: [32, 35, 42]
}], { x: 0.5, y: 4, w: 6, h: 3, lineSize: 3, lineSmooth: true });

// Pie chart
slide.addChart(pres.charts.PIE, [{
  name: "Share", labels: ["A", "B", "Other"], values: [35, 45, 20]
}], { x: 7, y: 1, w: 5, h: 4, showPercent: true });
```

### Better-Looking Charts

Default charts look dated. Apply these options for a modern, clean appearance:

```javascript
slide.addChart(pres.charts.BAR, chartData, {
  x: 0.5, y: 1, w: 9, h: 4, barDir: "col",

  // Custom colors (match your presentation palette)
  chartColors: ["0D9488", "14B8A6", "5EEAD4"],

  // Clean background
  chartArea: { fill: { color: "FFFFFF" }, roundedCorners: true },

  // Muted axis labels
  catAxisLabelColor: "64748B",
  valAxisLabelColor: "64748B",

  // Subtle grid (value axis only)
  valGridLine: { color: "E2E8F0", size: 0.5 },
  catGridLine: { style: "none" },

  // Data labels on bars
  showValue: true,
  dataLabelPosition: "outEnd",
  dataLabelColor: "1E293B",

  // Hide legend for single series
  showLegend: false,
});
```

**Key styling options:**
- `chartColors: [...]` - hex colors for series/segments
- `chartArea: { fill, border, roundedCorners }` - chart background
- `catGridLine/valGridLine: { color, style, size }` - grid lines (`style: "none"` to hide)
- `lineSmooth: true` - curved lines (line charts)
- `legendPos: "r"` - legend position: "b", "t", "l", "r", "tr"

---

## Slide Masters

```javascript
pres.defineSlideMaster({
  title: 'TITLE_SLIDE', background: { color: '283A5E' },
  objects: [{
    placeholder: { options: { name: 'title', type: 'title', x: 1, y: 2, w: 8, h: 2 } }
  }]
});

let titleSlide = pres.addSlide({ masterName: "TITLE_SLIDE" });
titleSlide.addText("My Title", { placeholder: "title" });
```

---

## Common Pitfalls

⚠️ These issues cause file corruption, visual bugs, or broken output. Avoid them.

1. **NEVER use "#" with hex colors** - causes file corruption
   ```javascript
   color: "FF0000"      // ✅ CORRECT
   color: "#FF0000"     // ❌ WRONG
   ```

2. **NEVER encode opacity in hex color strings** - 8-char colors (e.g., `"00000020"`) corrupt the file. Use the `opacity` property instead.
   ```javascript
   shadow: { type: "outer", blur: 6, offset: 2, color: "00000020" }          // ❌ CORRUPTS FILE
   shadow: { type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.12 }  // ✅ CORRECT
   ```

3. **Use `bullet: true`** - NEVER unicode symbols like "•" (creates double bullets)

4. **Use `breakLine: true`** between array items or text runs together

5. **Avoid `lineSpacing` with bullets** - causes excessive gaps; use `paraSpaceAfter` instead

6. **Each presentation needs fresh instance** - don't reuse `pptxgen()` objects

7. **NEVER reuse option objects across calls** - PptxGenJS mutates objects in-place (e.g. converting shadow values to EMU). Sharing one object between multiple calls corrupts the second shape.
   ```javascript
   const shadow = { type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.15 };
   slide.addShape(pres.shapes.RECTANGLE, { shadow, ... });  // ❌ second call gets already-converted values
   slide.addShape(pres.shapes.RECTANGLE, { shadow, ... });

   const makeShadow = () => ({ type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.15 });
   slide.addShape(pres.shapes.RECTANGLE, { shadow: makeShadow(), ... });  // ✅ fresh object each time
   slide.addShape(pres.shapes.RECTANGLE, { shadow: makeShadow(), ... });
   ```

8. **Don't use `ROUNDED_RECTANGLE` with accent borders** - rectangular overlay bars won't cover rounded corners. Use `RECTANGLE` instead.
   ```javascript
   // ❌ WRONG: Accent bar doesn't cover rounded corners
   slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1, y: 1, w: 3, h: 1.5, fill: { color: "FFFFFF" } });
   slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 1, w: 0.08, h: 1.5, fill: { color: "0891B2" } });

   // ✅ CORRECT: Use RECTANGLE for clean alignment
   slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 1, w: 3, h: 1.5, fill: { color: "FFFFFF" } });
   slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 1, w: 0.08, h: 1.5, fill: { color: "0891B2" } });
   ```

---

## Quick Reference

- **Shapes**: RECTANGLE, OVAL, LINE, ROUNDED_RECTANGLE
- **Charts**: BAR, LINE, PIE, DOUGHNUT, SCATTER, BUBBLE, RADAR
- **Layouts**: LAYOUT_16x9 (10"×5.625"), LAYOUT_16x10, LAYOUT_4x3, LAYOUT_WIDE
- **Alignment**: "left", "center", "right"
- **Chart data labels**: "outEnd", "inEnd", "center"

---

## Layout Code Patterns

Ready-to-use code blocks. All assume `T`, `F`, `SW`/`SH`/`M`/`CW`/`CY`/`CH` constants and helper functions from `references/themes.md` script template.

### Two-Column Comparison

```javascript
const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
sTitle(sl, "Comparison Title");
const c = cols2();
addCard(sl, c.left.x, CY, c.left.w, CH, T.ac.pos, "Approach A", "Description of approach A and its key properties.");
addCard(sl, c.right.x, CY, c.right.w, CH, T.ac.neg, "Approach B", "Description of approach B and its key properties.");
```

### Three-Card Grid

```javascript
const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
sTitle(sl, "Three Key Takeaways");
const gap = 0.25, cardW = (CW - 2 * gap) / 3;
const colors = [T.ac.pos, T.ac.neg, T.ac.emp];
["Point One", "Point Two", "Point Three"].forEach((title, i) => {
  const x = M + i * (cardW + gap);
  addNumCard(sl, i + 1, title, "Supporting detail text here.", x, CY + 0.15, cardW, CH - 0.15, colors[i]);
});
```

### Process Flow

```javascript
const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
sTitle(sl, "Pipeline Overview");
addFlow(sl, ["Input", "Encode", "Transform", "Decode", "Output"], CY + 0.3);
addBullets(sl, [
  "Step 1: Preprocess raw input data",
  "Step 2: Map to latent representation",
  "Step 3: Apply learned transformation"
], M, CY + 1.2, CW, CH - 1.2);
```

### Formula-Centered

```javascript
const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
sTitle(sl, "Core Equation");
sl.addText("The loss function minimizes:", { x: M, y: CY, w: CW, h: 0.35,
  fontFace: F.body.face, fontSize: F.body.size, color: T.tx.pri });
addFormula(sl, "f01-loss", CY + 0.5, 0.45);
sl.addText("where x is the input and y is the target label.", { x: M, y: CY + 1.2, w: CW, h: 0.3,
  fontFace: F.caption.face, fontSize: F.caption.size, color: T.tx.sec });
```

### Striped Table

```javascript
const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
sTitle(sl, "Performance Comparison");
addTable(sl,
  ["Method", "Accuracy", "Speed", "Memory"],
  [
    ["Baseline",  "78.2%", "1.0x", "4 GB"],
    ["Ours",      "92.1%", "1.3x", "6 GB"],
    ["Ours+Opt",  "91.8%", "2.1x", "4 GB"],
  ],
  M, CY + 0.1, CW, [3.5, 1.5, 1.5, 1.5]
);
```

### Table with OMML Math (Pattern 10)

```javascript
// Table cells containing math use OMML placeholders — inject_omml.py replaces them post-generation
const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
sTitle(sl, "Complexity Comparison");
addTable(sl,
  ["Algorithm", "Time", "Space", "Rounds"],
  [
    ["Naive",     {text: "{{MATH:f50}}", options: {}}, {text: "{{MATH:f51}}", options: {}}, "1"],
    ["Ours",      {text: "{{MATH:f52}}", options: {}}, {text: "{{MATH:f53}}", options: {}}, {text: "{{MATH:f54}}", options: {}}],
    ["Ours+Opt",  {text: "{{MATH:f55}}", options: {bold: true, color: T.ac.pos}}, {text: "{{MATH:f56}}", options: {}}, "1"],
  ],
  M, CY + 0.1, CW, [3, 2, 2, 2]
);
// formulas.json entries: {"id":"f50","latex":"O(n^2)","render":"omml"}, ...
sl.addText("Bold green = best in column", { x: M, y: CY + 1.65, w: CW, h: 0.3,
  fontFace: F.small.face, fontSize: F.small.size, color: T.tx.sec, shrinkText: true });
```

### Table + Insight (Pattern 11)

```javascript
// Upper 55%: table, lower 45%: insight card with key takeaway
const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
sTitle(sl, "Benchmark Results");
const tblY = CY + 0.1, tblRows = 4, rH = 0.35;
addTable(sl,
  ["Dataset", "Baseline", "Ours", "Δ"],
  [
    ["CIFAR-10",  "93.2%", "96.8%", {text: "+3.6%", options: {bold: true, color: T.ac.pos}}],
    ["CIFAR-100", "74.1%", "79.5%", {text: "+5.4%", options: {bold: true, color: T.ac.pos}}],
    ["ImageNet",  "76.3%", "78.9%", "+2.6%"],
    ["Places365", "55.1%", "57.2%", "+2.1%"],
  ],
  M, tblY, CW, [3, 2, 2, 2], rH
);
const insY = tblY + rH * (tblRows + 1) + 0.15;
addCard(sl, M, insY, CW, CH - (insY - CY) - 0.05, T.ac.pos,
  "Key Finding",
  "Largest gains on fine-grained tasks (CIFAR-100: +5.4%). Improvement correlates with number of categories — our method benefits most from richer label spaces."
);
```

### Continuation Table (Pattern 12)

```javascript
// Page 1 of a long table (rows 1-8)
const sl1 = pres.addSlide({ masterName: "CONTENT_SLIDE" });
sTitle(sl1, "Full Results");
addTable(sl1, headers, rowsPage1, M, CY + 0.1, CW, colWidths);
sl1.addText("Continued on next slide", { x: M, y: SH - M - 0.25, w: CW, h: 0.25,
  fontFace: F.caption.face, fontSize: F.caption.size, color: T.tx.mut, align: "right", shrinkText: true });

// Page 2 — uses addTableCont which auto-adds " (cont'd)" to title
const sl2 = pres.addSlide({ masterName: "CONTENT_SLIDE" });
addTableCont(sl2, "Full Results", headers, rowsPage2, M, CY + 0.1, CW, colWidths);
```

### Stat Callout

```javascript
const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
sTitle(sl, "Key Result");
const cx = SW / 2;
sl.addText("92.1%", { x: cx - 2, y: CY + 0.2, w: 4, h: 1.2,
  fontFace: F.stat.face, fontSize: F.stat.size, bold: true,
  color: T.ac.pos, align: "center", valign: "middle", margin: 0 });
sl.addText("Top-1 Accuracy on ImageNet", { x: cx - 2.5, y: CY + 1.5, w: 5, h: 0.4,
  fontFace: F.title.face, fontSize: F.title.size, bold: true,
  color: T.tx.pri, align: "center", margin: 0 });
sl.addText("Surpassing the previous SOTA by 3.9 percentage points\nwith 40% fewer parameters.", {
  x: cx - 3, y: CY + 2.1, w: 6, h: 0.6,
  fontFace: F.body.face, fontSize: F.body.size,
  color: T.tx.sec, align: "center", margin: 0 });
```

### Text-Left Diagram-Right

```javascript
const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
sTitle(sl, "System Architecture");
const c = cols2();
addBullets(sl, ["Component A handles input", "Component B transforms", "Component C routes output"],
  c.left.x, CY, c.left.w, CH);
addDiagramAt(sl, "d01-arch", c.right.x, CY, c.right.w, CH);
```

### Full-Width Diagram

```javascript
const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
sTitle(sl, "Data Flow Pipeline");
addDiagram(sl, "d02-flow", CY + 0.1, CH - 0.2);
```

### Chart-Centered

```javascript
const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
sTitle(sl, "Performance Breakdown");
addChart(sl, pres.charts.BAR, [
  { name: "Accuracy", labels: ["Model A", "Model B", "Ours"], values: [78, 85, 92] }
], M, CY + 0.1, CW, CH - 0.5, { barDir: "col", showValue: true, dataLabelPosition: "outEnd" });
sl.addText("All models evaluated on the same test set (n=10,000).", {
  x: M, y: SH - M - 0.25, w: CW, h: 0.25,
  fontFace: F.caption.face, fontSize: F.caption.size, color: T.tx.mut,
  align: "center", italic: true, margin: 0
});
```
