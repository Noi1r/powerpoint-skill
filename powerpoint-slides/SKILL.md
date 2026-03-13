---
name: powerpoint-slides
description: |
  Create visually rich PowerPoint (.pptx) presentations from academic papers, research notes,
  or any content the user wants in slide format. Uses PptxGenJS + LaTeX formula rendering.
  Always use this skill when the user wants PPT/PPTX output instead of Beamer/LaTeX slides.
  Trigger on: powerpoint, pptx, PPT, make a ppt, 做PPT, 做幻灯片, make slides (non-LaTeX),
  prepare a presentation (when context implies PPT), 做个报告, presentation slides,
  help me prepare a talk (when not Beamer), convert paper to slides (when PPT implied).
  Even if the user just says "make slides" or "做个演示" without specifying format, trigger this
  skill and ask whether they want PPT or Beamer — don't silently default to Beamer.
  Do NOT trigger on: beamer, latex slides, .tex files, tikz — those belong to the beamer skill.
argument-hint: "[action] [file] — actions: create, compile, review, audit, pedagogy, excellence, visual-check, validate, extract-figures"
allowed-tools: ["Read", "Write", "Edit", "Bash", "Grep", "Glob", "Agent", "AskUserQuestion", "TaskCreate", "TaskUpdate", "TaskList", "TaskGet"]
---

# PowerPoint Slides Workflow

Academic PowerPoint presentation skill. Full lifecycle:
create → compile → review → polish → verify.

Execution model: Claude writes a PptxGenJS (Node.js) script → executes → `.pptx`.
Math formulas: OMML (native math, default) or LaTeX → PNG/SVG pipeline (paragraph-mode fallback).

---

## 0. Quick Reference

| Task | Command | Description |
|------|---------|-------------|
| Create from paper | `create [topic]` | Full Phase 0-5 pipeline |
| Execute JS script | `compile [file.js]` | Run script, produce .pptx |
| Proofread | `review [file.pptx]` | Grammar, typos, consistency |
| Visual audit | `audit [file.pptx]` | Per-slide layout inspection |
| Teaching quality | `pedagogy [file.pptx]` | 13 pedagogical patterns |
| Full review | `excellence [file.pptx]` | 5 parallel agent review |
| Visual check | `visual-check [file.pptx]` | PDF→image systematic check |
| Validate metrics | `validate [file.pptx] [duration]` | Slide count, file size |
| Extract figures | `extract-figures [file.pdf] [pages]` | Extract paper figures for slides |

Parse `$ARGUMENTS` to determine which action to run. If no action specified, ask.

---

## 1. Hard Rules (Non-Negotiable)

1. **Motivation before formalism** — every concept starts with "Why?" before "What?". Never drop a definition without context.
2. **Worked example within 2 slides** — every definition must have a concrete worked example within 2 slides of its introduction.
3. **Telegraphic style** — keyword phrases, not full sentences. Slides are speaker prompts, not manuscripts. Exception: one framing sentence per slide to set context.
4. **Every slide earns its place** — must contain at least one substantive element (formula, diagram, table, theorem, algorithm, chart). A slide with only 3 short bullets and nothing else must be merged or enriched.
5. **Max 2 accent cards per slide** — more dilutes emphasis. Accent cards are colored boxes, callout shapes, or highlighted regions. Demote lower-priority items to plain text.
6. **Reference slide** — the second-to-last slide (before Thank You) must list key cited works. Include the primary paper and 3-5 most relevant references.
7. **Color-blind safe palette** — blue+orange preferred, never red+green for binary contrasts. WCAG AA contrast ratio ≥ 4.5:1 for text on backgrounds. Semantic colors defined in theme.
8. **Font size hierarchy** — `F` constant object defines default baselines: cover 48pt, section 36pt, title bar 22pt, body 15pt, small 12pt, caption 11pt, cardTitle 18pt, tblHead 13pt, tblCell 12pt. Never go below 10pt. When a card or region has sparse content (text fills <50% of available area), scale up by 2-4pt in the QA fix pass.
9. **Backup slides** — 3-5 after Thank You, separated by an appendix divider slide. For anticipated questions: detailed proofs, extended comparisons, additional data.
10. **Verify after every task** — execute JS → `soffice` → PDF → image inspection. No task is complete without visual verification.
11. **No AI-signature accent lines** — never place decorative lines directly under titles. This is a tell-tale of AI-generated slides and looks unprofessional.
12. **Layout diversity** — never use the same layout type for 3+ consecutive slides. Alternate among: full-width, two-column, formula-centered, stat-callout, icon-grid, timeline, table, chart, etc.
13. **Formula rendering** — OMML (native PowerPoint math) for display/inline formulas (default). PNG@600DPI or SVG for paragraph-mode formulas (fallback). Blurry image formulas unacceptable.
14. **16:9 layout, 0.5" margins** — `LAYOUT_16x9` (10" x 5.625"). Maintain 0.5" breathing room on all four sides. Content area: 9" x 4.625".
15. **Never reuse PptxGenJS option objects** — the library mutates objects in-place (e.g., converting shadow values to EMU). Always create fresh objects per call. Use factory functions: `const makeShadow = () => ({...})`.
16. **No # prefix on hex colors** — use `"FF0000"` not `"#FF0000"`. The `#` prefix corrupts the output file silently.
17. **Visual theme consistency** — the chosen palette applies globally: slide backgrounds, text colors, formula foreground, chart colors, accent card fills, divider lines. No off-theme colors.
18. **No Unicode math approximations** — never use Unicode subscripts (₀, ₁), superscripts (ᵖᵉ), or special math symbols (∑, ∏, ∈, ∀, ∝) as plain text inside cards or text boxes. All mathematical notation — including inline formulas, variable-with-subscript, Greek letters in equations — must be rendered via OMML (preferred) or the LaTeX image pipeline (for paragraph mode). For card bodies that mix text and math, use OMML placeholders for individual formulas or render the entire body as a single LaTeX image. The only exceptions: simple single Latin letters (x, y, n) and plain numbers without operators.
19. **Vertical content distribution** — content below the title bar must span ≥75% of `CH` vertically. If all elements end in the upper half, something is wrong: increase card heights, add spacing, scale up fonts, or add a summary/insight card at the bottom. Every content slide should feel "filled" — no large empty patches below the last element.
20. **Content density lower bounds** — a slide with ≤3 short text-only bullets (no formula, diagram, table, or card) is **too sparse** — merge with an adjacent slide or enrich with a visual element. Pure text-only slides should be ≤20% of the total deck. Each card must contain enough text to fill ≥40% of its interior area; if not, scale up font by 2-4pt or merge cards.
21. **Boundary validation** — every element must satisfy `y + h ≤ SH - M` (5.125"). During script generation, mentally verify each element's bottom edge before writing it. Common traps: stacking cards below a table without accounting for table rendered height; `addCardFormula` with a formula image taller than the card body area; two-row card layouts where the second row overshoots. When using `addCardFormula`, the formula height parameter `fH` must be ≤ `cardH - 0.75` (title area). If content cannot fit, reduce card height, shrink formula `targetH`, or split across two slides. The visual QA agent must specifically flag any element whose bottom edge is clipped by the slide boundary.
22. **Reference slide formatting** — strictly single-column layout. Divide `CH` equally among N references, each entry vertically centered in its row. Use `F.body.size` font (or `F.small.size` if many entries). No divider lines, no decorative elements — just evenly distributed text rows filling the page. If references exceed 8 per page, continue on a second reference slide. The reference slide must look evenly filled — no large blank patches at top or bottom.
23. **OMML post-processing & sizing** — after PptxGenJS generates the .pptx, run `inject_omml.py` to replace `{{MATH:id}}` placeholders with native OMML math. Verify zero residual `{{MATH:` placeholders in the output. OMML formulas display as blank or distorted in LibreOffice — this is expected and not a bug. Only check image formula clarity in visual QA. **CRITICAL OMML sizing rule**: OMML formulas render at the text box's font size and do NOT respect the text box height constraint — tall constructs like `\sqrt{}`, `\frac{}{}`, `\sum`, `\prod` will visually overflow the text box. To prevent overlap: (a) never place an OMML formula directly above or below a card/table with tight spacing — leave ≥0.15" extra vertical gap per level of nesting (fractions, roots, large operators); (b) for complex formulas with roots, fractions, or stacked operators, prefer image rendering (`"render": "image"` in formulas.json) which respects exact pixel dimensions; (c) when using `addMathText` with `F.body.size` (15pt), expect the rendered height to be ~0.4" for simple formulas but up to 0.7" for formulas with `\sqrt{}` or `\frac{}{}`; (d) when an OMML formula must appear between two elements, compute spacing as if the formula is 1.5–2× the `targetH` parameter.
24. **Diagram engine selection** — Graphviz for structural/flow diagrams (architecture, dependency, comparison, tree). Mermaid for behavioral/temporal diagrams (sequence, state, Gantt, ER). TikZ only for math-intensive or paper-faithful reproductions. Don't use Graphviz for sequence diagrams, don't use Mermaid for complex node layouts.
25. **Diagram theme consistency** — all rendered diagrams must use the current theme's colors. Node fill = `cardFill` equivalent, borders = `ac.pos`, text = `tx.pri`. Graphviz/Mermaid default colors (blue/black) must never appear in final output. Pass `--theme` to `render_diagrams.py`.
27. **Text overflow prevention & padding** — PptxGenJS text boxes do NOT clip overflow — text that exceeds the box renders outside, overlapping elements below. **Three mandatory protections**: (a) `shrinkText: true` on ALL text boxes (maps to `<a:normAutofit/>`). **NEVER use `autoFit: true`** — it maps to `<a:spAutoFit/>` which expands the shape. (b) Card text uses `const px = x + 0.25, pw = w - 0.4` — giving 0.19" clear of accent bar and 0.15" from right edge. Title: `(px, y+0.1, pw, 0.4)`, body: `(px, y+0.48, pw, h-0.62)` with `margin: [2, 0, 8, 0]`. (c) **Vertical gap discipline**: every element placed below another must start ≥0.1" after the previous element's bottom. Use `gap()` helper to verify at script-generation time. Standalone `addText` calls between cards/formulas are the #1 source of overlap bugs — always compute y from the preceding element's known bottom, never approximate. **Content limit**: card body should have ≤4 short lines per inch of body height. If more text is needed, increase card height or split content.
28. **Figure attribution** — figures extracted from papers must include source attribution (`Source: Author et al., Year`) via `addFigure()` caption. Width <800px requires warning to user about projection blur. Never extract tables — rebuild with PptxGenJS `addTable()`.
29. **Visual-first planning** — audiences grasp structure through visual anchors, not bullet lists. During Phase 2, ask for each slide: "Can the audience understand this in 10 seconds from text alone?" If not, the slide needs a visual element (diagram, table, chart, or figure). Common signals: multi-step processes, component relationships, data comparisons, hierarchies, numerical trends — these almost always need visuals rather than prose. Phase 2 self-check: every section has ≥1 non-formula visual element; slides with zero visuals (no diagram, table, chart, or figure) must stay ≤30% of the deck.

---

## 2. Color Themes & Visual System

5 themes: **Academic Light** (default), **Midnight**, **Ocean**, **Forest**, **Sandwich**.

**Read `references/themes.md`** for full theme JSON blocks, slide master definitions, typography table, semantic colors, script template, and scoring rubric. These are needed in Phase 4 (script generation) and Phase 5 (QA scoring).

Key points (always in context):
- 4 slide masters per theme: `TITLE_SLIDE`, `SECTION_SLIDE`, `CONTENT_SLIDE`, `THANK_YOU`
- Typography: Georgia (titles) + Calibri (body)
- Semantic colors: positive=`0173B2`, negative=`DE8F05`, emphasis=`029E73`, neutral=`8C8C8C`
- Sandwich theme has dual formula colors (light bg + dark bg variants)

### 2.1 Layout Type Pool

| Layout | Use Case |
|--------|----------|
| `title` | Cover slide |
| `section-divider` | Chapter transition |
| `text-left-image-right` | Concept + illustration (55/45 split) |
| `two-column` | Comparison / parallel content |
| `formula-centered` | Formula-dominant slide |
| `formula-with-annotation` | Formula + side annotations |
| `stat-callout` | Big number / key conclusion |
| `icon-grid` | Multiple points with colored shapes |
| `timeline` | Process / history flow |
| `table` | Data comparison |
| `chart` | Chart/graph display |
| `text-left-diagram-right` | Text + diagram side-by-side (55/45 split) |
| `full-diagram` | Full-width diagram |
| `chart-centered` | Chart + caption |
| `full-image` | Full-bleed image + overlay |
| `references` | Bibliography |
| `thank-you` | Closing slide |
| `backup` | Backup/appendix slides |

---

## 3. Actions

### 3.1 `create [topic]` — Full Pipeline (Phase 0-5)

Collaborative, iterative presentation creation. **Strict phase gates — never skip ahead.**

#### Phase 0: Material Analysis

**Read first, ask later.** Must understand the content before asking meaningful questions.

- Read the full paper/materials thoroughly
- Extract: core contributions, key techniques, main theorems, comparison with prior work
- Map notation conventions
- Identify logical structure and slide-worthy sections
- Note: prerequisite knowledge, natural section boundaries, what could be skipped or expanded

**Do NOT present results or ask questions yet — proceed directly to Phase 1.**

#### Phase 1: Requirements Interview (MANDATORY)

Conduct a content-driven interview via AskUserQuestion. Questions are informed by Phase 0 analysis.

**Required questions** (always ask):
1. **Duration** — How long is the presentation?
2. **Audience level** — Who are the listeners? (experts, graduate students, general academic)
3. **Prerequisites** — List concrete technical dependencies from Phase 0. Ask which the audience knows.
4. **Content scope** — Offer the paper's actual components as options. Which to emphasize, skip, or briefly mention?
5. **Depth vs breadth** — If the paper has both overview and detailed constructions, ask preference.
6. **Visual style** — Which theme?
   - A) Academic Light — white/light gray, clean and precise (default)
   - B) Midnight — dark throughout, tech/premium feel
   - C) Ocean — blue tones, professional academic
   - D) Forest — green natural tones
   - E) Sandwich — dark title/conclusion + light content

**Slide count heuristic** (~1 slide per 1.5-2 minutes):

| Duration | Total slides | Intro/Motivation | Methods/Background | Core content | Summary |
|----------|-------------|------------------|--------------------|-------------|---------|
| 5min (lightning) | 5-7 | 1-2 | 0-1 | 2-3 | 1 |
| 10min (short) | 8-12 | 2 | 1-2 | 4-5 | 1 |
| 15min (conference) | 10-15 | 2-3 | 2-3 | 5-7 | 1-2 |
| 20min (seminar) | 13-18 | 3 | 2-3 | 6-9 | 2 |
| 45min (keynote) | 22-30 | 4-5 | 5-7 | 10-14 | 2-3 |
| 90min (lecture) | 45-60 | 5-6 | 8-12 | 25-35 | 3-4 |

**Talk-type tips** — different formats demand different strategies. Misjudging the format is the #1 cause of bad presentations:

| Talk type | Key emphasis | Common mistake |
|-----------|-------------|----------------|
| Lightning (5min) | One core message, no background | Cramming a full talk into 5 minutes |
| Conference (10-20min) | 1-2 key results, fast methods overview | Too much technical detail, no big picture |
| Seminar (45min) | Deep dive OK, but need visual rhythm | Wall-to-wall formulas without examples |
| Defense/Thesis | Demonstrate mastery, systematic coverage | Skipping motivation, rushing results |
| Journal club | Critical analysis, facilitate discussion | Summarizing without evaluating |
| Grant pitch | Significance → feasibility → impact | Too technical, not enough "why it matters" |

**Pacing principle**: spend 40-50% of time on core content (results/techniques). Max 3-4 consecutive theory-heavy slides before a worked example or visual break — audiences disengage after ~5 minutes of uninterrupted formalism.

#### Phase 2: Structure Plan (GATE — user must approve)

Detailed outline per section:
- Section title
- Allocated slide count
- Key content points per slide (1-2 lines)
- **Layout type** for each slide (from layout pool in Section 2.4)
- Planned diagrams/charts (brief description)
- Notation to be introduced

**Phase 2 self-check before presenting:**
- Each section has ≥1 non-formula visual element (diagram, table, chart, or figure)
- Slides with zero visuals ≤30% of deck (R29)
- Every multi-step process, component relationship, or data comparison has a planned visual — not just bullets

**Present the plan. Ask: structure OK? User must approve before proceeding.**

#### Phase 3: Formula Preparation

1. Scan Phase 2 structure, collect all formulas into a list.
2. Write `formulas.json` with `render` field (`"omml"` | `"image"` | `"auto"`; default `"auto"` → paragraph→image, else→omml):
   ```json
   [
     {"id": "f01-entropy", "latex": "H(X) = -\\sum_{x} p(x) \\log p(x)", "mode": "display"},
     {"id": "p01-body", "latex": "\\begin{minipage}{8cm}...", "mode": "paragraph", "fg_color": "1A1A2E"}
   ]
   ```
3. Run renderer (only renders image entries; omml entries get manifest metadata only):
   ```bash
   python ~/.claude/skills/powerpoint-slides/scripts/render_latex.py formulas.json formulas/
   ```
4. Check `formulas/manifest.json` for errors. Fix LaTeX or fall back to plain text.

**Sandwich theme**: OMML formulas inherit text color automatically — no dual rendering needed. Only **image-rendered** formulas (paragraph mode) need two variants:
```json
[
  {"id": "p01-light", "latex": "...", "fg_color": "2D3436", "mode": "paragraph"},
  {"id": "p01-dark",  "latex": "...", "fg_color": "FFFFFF", "mode": "paragraph"}
]
```

**Math-heavy paragraphs** (MANDATORY — Hard Rule 18): scan every card body and text box planned in Phase 2. If ANY of these appear, the text MUST be rendered as a LaTeX image (not Unicode approximations):
- Subscripts or superscripts (x_i, e^{ax}, H_0)
- Greek letters in equations (α, β, Σ, ∏)
- Operators with operands (∈, ⊆, ∝, ≥)
- Fractions, summations, products, integrals

For isolated formulas (display/inline): use OMML (default `render:"auto"` handles this). For card bodies with mixed text+math: render the entire paragraph as a single LaTeX image using `\text{}` for non-math words — add to `formulas.json` with `"mode": "paragraph"` (auto-routes to image). In Phase 4, embed via `addFormula()` / `addCardFormula()` — helpers auto-route based on manifest.

**Error recovery**: if `render_latex.py` reports errors in manifest (`"error": "..."`), fix the LaTeX source, re-run. Fallback: render as plain text in PptxGenJS.

**Phase 3-4 iteration**: if Phase 4 discovers additional formulas needed, append to `formulas.json` and re-run incrementally.

#### Phase 3.5: Diagram Preparation

1. Scan Phase 2 structure, identify diagrams and visual elements that need rendering.
2. Classify each into the 5-layer pipeline:
   - **graphviz** — architecture, flow, dependency, neural network, tree, comparison diagrams (DOT language)
   - **mermaid** — sequence, Gantt, pie, state, ER diagrams
   - **tikz** — math-intensive, geometric constructions, paper-faithful reproductions
   - **shapes** — simple annotations, arrows, highlights (PptxGenJS native — handle in Phase 4, not here)
   - **extract** — figures from paper PDFs (extract-figures workflow)
3. Write `diagrams.json` (see `diagram-rendering.md` for schema).
4. Run the renderer:
   ```bash
   python ~/.claude/skills/powerpoint-slides/scripts/render_diagrams.py diagrams.json diagrams/ --theme <theme_name>
   ```
5. Check `diagrams/manifest.json` for errors. Fix DOT/Mermaid syntax or degrade to PptxGenJS shapes.

**DOT writing guide:** max ~15 nodes, `rankdir=LR` for horizontal flows / `rankdir=TB` for vertical hierarchies, `subgraph cluster_*` for grouping.

**Mermaid guide:** actor/participant ≤6, supported types: `sequenceDiagram`, `gantt`, `pie`, `stateDiagram-v2`, `erDiagram`.

**Read `diagram-rendering.md`** for full 5-layer architecture docs, `diagrams.json` schema, theme color mapping, and troubleshooting.

#### Phase 4: Generate JS Script + Execute

Write a complete PptxGenJS script. Install dependencies first:
```bash
npm ls pptxgenjs 2>/dev/null || npm install pptxgenjs
```

**Read `references/themes.md` → "Script Template" section** for the full starter template with theme constants (`T`), typography constants (`F`), layout constants (`SW`/`SH`/`M`/`CW`/`CY`/`CH`), slide masters, `addFormula()` helper, diagram helpers (`addDiagram`, `addDiagramAt`, `addFigure`), chart helper (`addChart`), and semantic helpers (`sTitle`, `sectionSlide`, `addBullets`, `addCard`, `addTable`, `addNumCard`, `addFlow`, `cols2`). Also read `pptxgenjs-reference.md` → "Layout Code Patterns" for 9 ready-to-use code blocks (including diagram and chart layouts), `formula-rendering.md` for the LaTeX pipeline, and `diagram-rendering.md` for the diagram pipeline.

**Use `F` object for all font sizes**, `cols2()` for two-column layouts, and helper functions for common patterns. Never hardcode font sizes — always reference `F.body.size`, `F.title.size`, etc.

**User-provided images** (logos, photos, screenshots): embed directly via `slide.addImage({ path: "image.png", ... })`. Calculate aspect ratio from original dimensions. Center with `x: (SW - w) / 2`. No special pipeline needed — just verify the file exists.

**Key rules for script generation:**
- Always use factory functions for shadows, fills, and other option objects
- Never use `#` prefix on hex colors
- Use `breakLine: true` between text array items
- Use `bullet: true` for list items, never unicode "•"
- Set `margin: 0` on text boxes that need precise alignment with shapes
- For `icon-grid` layout: use simple colored rectangles/circles as visual markers (shapes, not icon images)

##### 4a. Mathematical Slide Patterns

Math-heavy slides should follow one of these structural templates — they prevent the common trap of dumping a formula with no context:

**Definition slide:**
```
[Framing sentence: why this definition matters]
[Formal definition — display formula or card]
[Key properties / immediate consequences — 2-3 bullet items]
```

**Construction/Algorithm slide:**
```
[One-line goal statement]
[Core equation / algorithm steps]
[Complexity or performance: prover cost, verifier cost, soundness]
```

**Comparison slide:**
```
[Side-by-side table or two-column: prior work vs this work]
[1-2 lines highlighting the key difference]
```

**Theorem → Proof slide (two slides, never one):**
```
Slide A: [Informal statement] → [Formal theorem in card] → [Why it matters]
Slide B: [Proof sketch — key steps only, ≤5 lines. Full proof in backup.]
```

**Insight/Remark slide:**
```
[Observation the paper doesn't emphasize, or connection to related work]
[Why this matters / what it implies]
```

Every math slide must have a clear **takeaway** — the one thing the audience should remember from that slide.

##### 4b. Content Density Upper Bounds

Lower bounds are in R4/R20. Upper bounds prevent cognitive overload — a slide with too much is worse than one with too little, because the audience retains nothing:

- ≤ 7 bullet points or items per slide
- ≤ 2 displayed formulas per slide (more → split across slides)
- ≤ 5 new symbols introduced per slide
- ≤ 2 accent cards per slide (R5 already covers this)

**Density self-check after each batch:**
- Count slides with zero formulas/diagrams/tables → flag if >30% of batch
- Count slides with ≤3 short items and no math → candidates for merging
- Count slides exceeding upper bounds → split or redistribute

##### 4c. Table Best Practices

- Column alignment: numbers right-aligned, text left-aligned, short labels centered
- Max 6-7 columns, 8-10 rows per slide. More → split across slides or highlight a subset
- Highlight key cells with bold or `ac.pos` color — draw the eye to the result
- For comparison tables: bold the best result in each row/column
- Use striped rows (via `addTable` helper) for readability in tables with >5 rows
- Caption or insight below the table in `F.small.size` — never leave a table without a takeaway

##### 4d. Algorithm and Code Display

- Pseudocode ≤ 10 lines per slide. If longer, split into "high-level overview" and "key subroutine" slides
- Highlight the critical line(s) with `ac.pos` background color
- Input/output clearly stated at the top
- Variables in italic, functions in regular weight — consistent naming throughout
- For actual code (not pseudocode): use monospace font, show only the relevant fragment — never dump an entire source file

**Batching**: if >20 slides, write script in batches of 8-10 slides. Self-check density and layout diversity per batch. Final script is one file executed once.

**Opening strategies** (pick one):
- Provocative question the audience cannot immediately answer
- Surprising statistic that challenges assumptions
- Real-world failure/problem → "How do we solve this?"
- Visual demonstration — show the phenomenon before explaining it

**Closing strategies**:
- Callback to opening — revisit the opening question, now answered
- 3 key takeaways — numbered, telegraphic, one slide
- Future direction — open question that invites Q&A
- **Never end on a bare "Thank You"** — the second-to-last content slide is the lasting impression

##### 4e. Script Self-Review (before execution)

Re-read the completed JS script and verify before running `node`. Catching issues here avoids wasting a QA round:

*Structure:*
- [ ] Slide count matches Phase 2 plan (±2 tolerance)
- [ ] Logical flow: motivation → background → technique → results → summary
- [ ] No section has >4 consecutive formula-heavy slides without a worked example or visual break
- [ ] Section divider slides present between major sections

*Content density:*
- [ ] No slide has only ≤3 short bullets with no math/diagram/table
- [ ] No slide exceeds upper bounds (7 bullets, 2 formulas, 5 new symbols, 2 accent cards)
- [ ] Pure text-only slides ≤ 20% of deck

*Layout & positioning:*
- [ ] Every element satisfies `y + h ≤ SH - M` (R21)
- [ ] `shrinkText: true` on all multi-line text boxes (R27)
- [ ] No same layout type for 3+ consecutive slides (R12)
- [ ] `F` constants used for all font sizes — no hardcoded numbers

*Notation:*
- [ ] Same symbol used consistently throughout all slides
- [ ] Every symbol defined before first use

#### Phase 5: QA Loop (MANDATORY)

```
┌─→ 5a. Execute JS script → .pptx
│   5b. soffice → PDF → pdftoppm → slide images
│   5c. Subagent visual inspection (per-slide)
│   5d. markitdown text extraction → content verification
│   5e. Score (apply rubric)
│   5f. Fix (edit JS and/or re-render formulas → re-execute)
└── score < 90 AND round < 3: loop back to 5a
    score ≥ 90 OR round = 3: report to user
```

**5a. Execute + OMML injection**:
```bash
node generate_slides.js
# If manifest has any render:"omml" entries, inject OMML math:
python ~/.claude/skills/powerpoint-slides/scripts/inject_omml.py output.pptx formulas.json output.pptx
```

**5b. Convert to images**:
```bash
python ~/.claude/skills/powerpoint-slides/scripts/soffice.py --headless --convert-to pdf --outdir . output.pptx
pdftoppm -png -r 200 output.pdf /tmp/slide
```

**Quick thumbnail grid** (optional — for rapid overview before detailed inspection):
```bash
python ~/.claude/skills/powerpoint-slides/scripts/thumbnail.py output.pptx thumbnails --cols 4
```

**5c. Visual inspection** (via Agent subagents):

Slide images at 200 DPI are ~1-2MB each. Reading images accumulates in context — reading more than ~7 images in the main conversation or a single agent will exceed the 20MB limit. **NEVER read slide images directly in the main conversation.**

**Strategy**: dispatch parallel Agent subagents, each assigned 3-5 slides. Each agent reads its slides one at a time, inspects them, and returns a **text-only** report. The main conversation only receives text, keeping context clean.

```
Total slides: N
Agents: ceil(N/5) agents, each reading up to 5 consecutive slide images
All agents run in parallel
```

Each agent prompt:
> "Read slide images [start]-[end] from /tmp/slide-{NN}.png, ONE image per Read call.
> For each slide check: overflow, overlap, font legibility, formula clarity, contrast,
> layout clutter, content density, bottom-edge clipping.
>
> **CRITICAL — content density measurement**: judge blank space by the RENDERED PIXELS
> in the slide image, NOT by assumed text box boundaries. A text box can be 3" tall but
> render only 2 lines of text occupying 0.5" — that is 2.5" of visual blank space.
> Look at the actual slide image: where does visible content (text, shapes, images, formulas)
> END vertically? The gap from there to the slide bottom is the true blank space.
> Flag if visible content occupies <75% of the area below the title bar (R19).
>
> **CRITICAL — bottom-edge overflow (R21)**: check if any card, table, formula, or text
> is clipped at the slide bottom edge. Signs: text cut mid-line, card shadow missing
> at bottom, table rows disappearing. This is the #1 most common layout bug.
> Also check if a card title overlaps with its body content (title text running into
> body text below it — usually caused by long wrapped titles at large font sizes).
>
> **Diagram check**: labels readable, edges distinguishable, no node overlap,
> colors match slide theme (no Graphviz default blue/black), diagram fits
> within content area without cropping. Extracted figures have source caption.
>
> **OMML note**: OMML formulas display as blank in LibreOffice/PDF preview.
> This is a known limitation, not a bug. Only check image formula clarity.
>
> Return a TEXT report only — list issues by slide number + description + severity.
> If no issues found for a slide, report 'OK'."

Merge all agent reports into a single issues list.

**5a-post. XML validation** (after OMML injection): verify no residual `{{MATH:` placeholders:
```bash
python -c "
import zipfile, sys
with zipfile.ZipFile('output.pptx') as z:
    for n in z.namelist():
        if n.startswith('ppt/slides/') and n.endswith('.xml'):
            if b'{{MATH:' in z.read(n):
                print(f'RESIDUAL PLACEHOLDER in {n}'); sys.exit(1)
print('OK: no residual placeholders')
"
```

**5a-post2. Overlap & boundary check** (after OMML injection):
```bash
python ~/.claude/skills/powerpoint-slides/scripts/check_overlaps.py output.pptx
```
If exit code is non-zero (critical/major issues found), fix the JS script and re-run.
The checker is card-aware: it groups card internals (container rect + accent bar + title + body)
to avoid false positives. It detects: element overlaps, bottom-edge overflow, and tight gaps.

**5d. Content verification**:
```bash
pip install -q "markitdown[pptx]" 2>/dev/null
markitdown output.pptx > content.md
```
Read `content.md` to verify text content, notation consistency, spelling.

**5e. Quality Scoring**: Apply rubric from `references/themes.md` → "Quality Scoring Rubric". Thresholds: ≥90 deliver, 80-89 warnings, <80 must fix.

**5f. Fix**: edit JS script to fix issues. Re-render formulas if color/DPI problems. Re-execute and re-score. Max 3 rounds.

#### Post-Creation Checklist (final gate before reporting to user)

```
[ ] JS script executes without errors, .pptx generated
[ ] OMML injection successful, zero residual {{MATH:}} placeholders
[ ] QA score ≥ 90
[ ] Every definition has motivation + worked example within 2 slides
[ ] No slide exceeds density upper bounds (7 bullets, 2 formulas, 5 symbols, 2 cards)
[ ] No sparse slides (all slides have substantive content)
[ ] Diagrams use theme colors (no default blue/black)
[ ] Tables fit within content area, key cells highlighted
[ ] Notation consistent throughout — same symbol = same meaning
[ ] References slide present (second-to-last, before Thank You)
[ ] Slide images visually inspected (at least spot-check 3-5 slides via QA agents)
```

---

### 3.2 `compile [file.js]`

Execute an existing PptxGenJS script to generate .pptx.

```bash
npm ls pptxgenjs 2>/dev/null || npm install pptxgenjs
node FILE.js
```

Post-compile checks:
- Exit code 0?
- Output .pptx file exists?
- File size reasonable? (>10KB = has content, <50MB = reasonable)
- Report: output path, file size, success/failure

---

### 3.3 `review [file.pptx]`

Read-only proofreading report. No file edits.

Extract text:
```bash
pip install -q "markitdown[pptx]" 2>/dev/null
markitdown FILE.pptx > content.md
```

**4 check categories:**

| Category | Checks |
|----------|--------|
| Grammar | Subject-verb agreement, articles, prepositions, tense consistency |
| Typos | Spelling errors, search-replace artifacts, duplicate words, unreplaced placeholders (`[name]`, `[TODO]`) |
| Consistency | Citation format, notation, terminology, color usage across slides |
| Academic quality | Informal contractions, unsupported claims, ambiguous abbreviations |

**Report format per issue:**
```
### Issue N: [Brief description]
- **Location:** [slide number or title]
- **Current:** "[exact text]"
- **Proposed:** "[fix]"
- **Category / Severity:** [Category] / [High|Medium|Low]
```

---

### 3.4 `audit [file.pptx]`

Visual layout audit via image conversion. Read-only report.

```bash
python ~/.claude/skills/powerpoint-slides/scripts/soffice.py --headless --convert-to pdf --outdir /tmp FILE.pptx
pdftoppm -png -r 200 /tmp/FILE.pdf /tmp/audit-slide
```

**Per-slide visual checklist:**
- [ ] No text overflow at any edge
- [ ] No element overlap (text on text, text on image, shape on shape)
- [ ] All text legible (≥10pt equivalent)
- [ ] Tables and formulas fit within content area
- [ ] Consistent font sizes across similar slide types
- [ ] Adequate contrast between text and background
- [ ] No visual clutter (too many competing elements)
- [ ] Accent cards ≤ 2 per slide
- [ ] Layout diversity (no 3+ consecutive same type)

**Via Agent subagents**: dispatch parallel agents, each assigned 3-5 slides. Each agent reads its slides one at a time via Read tool and returns a text-only report. Never read slide images in the main conversation — images accumulate in context and exceed the 20MB limit after ~7 slides.

Report per issue with slide number, description, severity, and fix recommendation.

---

### 3.5 `pedagogy [file.pptx]`

Pedagogical review. Read-only report.

**13 teaching patterns to validate:**

| # | Pattern | Red Flag |
|---|---------|----------|
| 1 | Motivation before formalism | Definition without context |
| 2 | Incremental notation introduction | 5+ new symbols on one slide |
| 3 | Concrete examples after definitions | 2 consecutive definitions, no example |
| 4 | Progressive complexity | Advanced concept before prerequisite |
| 5 | Fragment reveal (problem → solution) | Dense theorem revealed all at once |
| 6 | Signpost slides at pivots | Abrupt topic jump, no transition |
| 7 | Two-slide strategy for dense theorems | Complex theorem crammed in 1 slide |
| 8 | Semantic color usage | Binary contrasts in same color |
| 9 | Card/box hierarchy | Wrong accent type for content |
| 10 | Card fatigue avoidance | 3+ accent cards on one slide |
| 11 | Socratic embedding | Zero questions in entire deck |
| 12 | Visual-first for complex concepts | Notation before visualization |
| 13 | Side-by-side comparison | Sequential slides for related definitions |

**Deck-level checks:**
- Narrative arc (motivation → build-up → climax → resolution)
- Pacing (max 3-4 theory-heavy slides before example or visual break)
- Visual rhythm (section dividers every 5-8 slides)
- Notation consistency throughout
- Student prerequisite assumptions reasonable

---

### 3.6 `excellence [file.pptx]`

Comprehensive multi-dimensional review. **Dispatch 5 parallel Agent calls.**

1. **Visual audit agent** — dispatch parallel Agent subagents, each assigned 3-5 slides. Each agent reads its slide images ONE at a time via Read tool, checks for overlap, font consistency, card fatigue, spacing issues, layout diversity, content density. Returns text-only report per slide with severity. Never read slide images in the main conversation.

2. **Pedagogy review agent** — "Extract text from the .pptx. Validate 13 pedagogical patterns and deck-level checks (narrative arc, pacing, visual rhythm, notation consistency). Report pattern-by-pattern."

3. **Proofreading agent** — "Extract text from the .pptx. Check grammar, spelling, citation consistency, notation consistency, academic quality. Report per issue with location and fix."

4. **Formula quality agent** — dispatch Agent subagents to check formula slide images (3-5 slides per agent, one image per Read call). Check: DPI adequate (not blurry), color matches slide theme, alignment correct, sizing consistent. Returns text-only report.

5. **Domain review agent** (optional — enabled with `--domain` flag or when paper is math/theory-heavy) — "Verify substantive correctness: assumptions stated, derivations valid, citation fidelity, claims supported. Report per issue with severity."

**After all agents return**, synthesize a combined report:

```markdown
# Excellence Review: [Filename]

## Overall Quality: [EXCELLENT / GOOD / NEEDS WORK / POOR]

| Dimension | Critical | Major | Minor |
|-----------|----------|-------|-------|
| Visual/Layout | | | |
| Pedagogical | | | |
| Proofreading | | | |
| Formula Quality | | | |
| Domain (if run) | | | |

### Critical Issues (Immediate Action Required)
### Major Issues (Next Revision)
### Recommended Next Steps
```

Quality score mapping: Excellent (0-2 critical, 0-5 major), Good (3-5 critical, 6-15 major), Needs Work (6-10 critical, 16-30 major), Poor (11+ critical, 31+ major).

---

### 3.7 `visual-check [file.pptx]`

PDF → image systematic visual review.

**Workflow:**

1. Convert:
   ```bash
   python ~/.claude/skills/powerpoint-slides/scripts/soffice.py --headless --convert-to pdf --outdir /tmp FILE.pptx
   pdftoppm -png -r 200 /tmp/FILE.pdf /tmp/vc-slide
   ```

2. Dispatch parallel Agent subagents for visual inspection, each assigned 3-5 slides. Each agent reads its slides one at a time via Read tool, runs the per-slide checklist, and returns a **text-only** report. **Never read slide images in the main conversation** — images accumulate in context and exceed 20MB after ~7 slides. Systematic per-slide checklist:
   - [ ] No text overflow at any edge (top, bottom, left, right)
   - [ ] No element overlap
   - [ ] All text legible at presentation distance
   - [ ] Formula images crisp (not blurry/pixelated)
   - [ ] Tables fit within slide width
   - [ ] Consistent font sizes across similar slides
   - [ ] Adequate text-background contrast
   - [ ] No visual clutter
   - [ ] Theme colors applied consistently
   - [ ] Content density adequate (text fills >50% of card/region area)

3. Report per issue:
   ```
   ### Slide N: [slide title]
   - **Issue:** [description]
   - **Severity:** Critical / Major / Minor
   - **Fix:** [specific recommendation]
   ```

---

### 3.8 `extract-figures [file.pdf] [pages]`

Extract figures from a paper PDF for use in slides.

1. Identify target figures — use `pdf_get_toc` / `pdf_read_pages` to locate. If unsure, ask user.
2. Write `diagrams.json` with `type: "extract"` entries, or extract directly:
   ```bash
   pdftoppm -png -r 300 -f PAGE -l PAGE paper.pdf diagrams/fig
   ```
3. If cropping needed, use Pillow to trim white margins and excess text regions.
4. Run renderer to produce manifest (or add entries to existing `diagrams/manifest.json`):
   ```bash
   python ~/.claude/skills/powerpoint-slides/scripts/render_diagrams.py diagrams.json diagrams/
   ```

**Rules:**
- Must attribute source (`Source: Author et al., Year`) unless it's the user's own paper
- Resolution check: <800px in any dimension → warn user about projection blur
- Never extract tables — rebuild with `addTable()` for crisp rendering
- Multi-figure pages: use `crop` field to isolate each figure separately

### 3.9 `validate [file.pptx] [duration]`

Automated quantitative validation. Checks measurable properties.

**Checks:**

1. **Slide count vs duration** (if duration provided):
   ```bash
   python ~/.claude/skills/powerpoint-slides/scripts/soffice.py --headless --convert-to pdf --outdir /tmp FILE.pptx
   pdfinfo /tmp/FILE.pdf | grep "Pages:"
   ```
   Compare against timing table. Flag if outside range.

2. **File size**:
   - \>50 MB: warning (slow to share)
   - \>100 MB: critical (likely uncompressed images)

3. **Content extraction**:
   ```bash
   markitdown FILE.pptx > content.md
   ```
   - Count slides with no text → flag empty slides
   - Check for placeholder text (`[TODO]`, `[XXX]`)

4. **Formula images** (if `formulas/` directory exists):
   - All manifest entries have files that exist
   - No manifest entries with `"error"` field
   - Image dimensions reasonable (width > 100px)

**Report format:**
```
# Validation Report: [Filename]

| Check | Result | Status |
|-------|--------|--------|
| Slide count | N slides / Xmin | OK / WARNING |
| File size | X.X MB | OK / WARNING |
| Empty slides | N found | OK / WARNING |
| Placeholder text | N found | OK / CRITICAL |
| Formula images | N/M valid | OK / WARNING |

Overall: PASS / PASS WITH WARNINGS / FAIL
```

---

## 4. Verification Protocol

**Every task ends with verification.** Non-negotiable.

```
[ ] JS script executes without errors (node exit code 0)
[ ] .pptx file generated and non-empty (>10KB)
[ ] soffice converts to PDF successfully
[ ] Slide images visually inspected (at least spot-check 3-5 slides)
[ ] Text content verified via markitdown (no garbled text, placeholders)
[ ] Score ≥ 90 (for create action) or issues documented (for review actions)
```

---

## 5. Troubleshooting

| Problem | Fix |
|---------|-----|
| `Cannot find module 'pptxgenjs'` | `npm install pptxgenjs` in working directory |
| Formula images missing / manifest errors | Check LaTeX syntax in `formulas.json` (double-escape `\\`), re-run `render_latex.py`. Fallback: plain text |
| soffice hangs or fails | `pkill -f soffice`, retry with `--headless` flag |
| Colors wrong in .pptx | Remove `#` prefix from hex. Use 6-char hex only. Transparency via `transparency` property |
| Shapes misaligned after multiple adds | Use factory functions — PptxGenJS mutates option objects in-place |
| `dot: command not found` | `brew install graphviz` |
| Diagram has default blue/black colors | Pass `--theme` flag to `render_diagrams.py` matching slide theme |
| Extracted figure blurry | Source PDF may be low-res; try higher DPI or find vector source |
| `pdfinfo: command not found` | `brew install poppler` (provides pdfinfo, pdftoppm) |

---

## 6. Dependencies

**Node.js**: `npm install pptxgenjs`
**Python**: `pip install Pillow lxml "markitdown[pptx]"`
**System**: `node`, `pandoc` (OMML conversion), `xelatex`, `pdfcrop`, `pdftoppm` (Poppler), `soffice` (LibreOffice), `pdfinfo` (Poppler), `dot` (Graphviz — `brew install graphviz`), `mmdc` (Mermaid CLI — `npm install -g @mermaid-js/mermaid-cli`, optional)
