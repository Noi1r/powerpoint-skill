# PowerPoint Slides Skill

An AI coding assistant skill for creating visually rich **PowerPoint (.pptx)** presentations from academic papers, research notes, or any structured content.

Built on **PptxGenJS** (Node.js) with native **OMML math** + LaTeX image rendering, **Graphviz/Mermaid/TikZ** diagram pipelines, and 5 built-in color themes.

Supports **Claude Code** natively. Adaptable to other AI coding assistants.

## Features

| Action | Description |
|--------|-------------|
| `create [topic/paper]` | Full lifecycle: material analysis → interview → structure plan → formula/diagram prep → script generation → QA loop |
| `compile [file.js]` | Execute PptxGenJS script to generate .pptx |
| `review [file.pptx]` | Read-only proofreading report (grammar, typos, consistency, academic quality) |
| `audit [file.pptx]` | Visual layout audit via PDF→image inspection |
| `visual-check [file.pptx]` | PDF-based systematic visual review with parallel subagents |
| `validate [file.pptx]` | Structural validation against skill constraints |
| `extract-figures [pdf]` | Extract figures from paper PDFs for slide inclusion |

### Highlights

- **5 color themes** — Academic Light (default), Midnight, Ocean, Forest, Sandwich (dark title + light content)
- **OMML math (native)** — display/inline formulas rendered as native PowerPoint math via pandoc, editable after generation
- **LaTeX image fallback** — paragraph-mode formulas rendered at 600 DPI via XeLaTeX pipeline
- **5-layer diagram pipeline** — Graphviz (structural) → Mermaid (behavioral) → TikZ (math) → PptxGenJS shapes (native) → PDF extraction
- **Theme color injection** — diagrams automatically use the slide theme's color palette
- **Quality scoring** — automated rubric (start at 100, deduct per issue) with visual QA via parallel subagents
- **29 hard rules** — covering content density, layout diversity, typography, math rendering, boundary validation, overflow prevention
- **Mathematical slide patterns** — 5 structural templates (Definition, Construction, Comparison, Theorem-Proof, Insight)
- **Content density guards** — both upper bounds (7 bullets, 2 formulas, 5 symbols/slide) and lower bounds (every slide earns its place)
- **Visual-first planning** — "10-second test" ensures diagrams/tables/charts where text alone fails
- **Pacing control** — talk-type tips, timing allocation, max 3-4 consecutive theory slides before a visual break
- **16:9 layout** with precise positioning via layout constants (`SW`/`SH`/`M`/`CW`/`CY`/`CH`)
- **17 helper functions** — `sTitle`, `sectionSlide`, `addBullets`, `addCard`, `addTable`, `addNumCard`, `addFlow`, `addChart`, `addFormula`, `addDiagram`, `cols2`, etc.
- **9 layout code patterns** — two-column, three-card, process flow, formula-centered, striped table, stat callout, text+diagram, full-diagram, chart-centered

## Prerequisites

### Required

```bash
# Node.js (PptxGenJS runtime)
brew install node

# pandoc (OMML formula conversion)
brew install pandoc

# TeX distribution (LaTeX image formulas + TikZ diagrams)
brew install --cask mactex

# Graphviz (diagram rendering)
brew install graphviz

# Python packages (image processing, OMML injection)
pip install Pillow lxml

# LibreOffice (PPT→PDF for visual QA)
brew install --cask libreoffice

# markitdown (text extraction for content verification)
pip install "markitdown[pptx]"
```

### Recommended

```bash
# Mermaid CLI (sequence/Gantt/pie/state/ER diagrams)
# Without it, Mermaid diagrams fall back to Graphviz or PptxGenJS shapes
npm install -g @mermaid-js/mermaid-cli
```

### Optional

```bash
# Ghostscript (SVG formula output via dvisvgm)
brew install ghostscript
```

### pdf-mcp (Recommended)

Install [pdf-mcp](https://github.com/Noi1r/pdf-mcp) (fork with file-based image extraction — avoids base64 token bloat) so the AI assistant can read papers and extract figures:

```bash
pip install git+https://github.com/Noi1r/pdf-mcp.git
claude mcp add pdf-mcp --scope user pdf-mcp
```

## Installation

Clone the repo:

```bash
git clone https://github.com/Noi1r/powerpoint-skill.git
```

### Claude Code

Copy the skill directory into your Claude Code skills folder:

```bash
mkdir -p ~/.claude/skills
cp -r powerpoint-skill/powerpoint-slides ~/.claude/skills/
```

Restart Claude Code. The skill triggers automatically when you mention powerpoint, pptx, PPT, slides, or related keywords.

## Usage

**Create slides from a paper:**
```
Help me make a PPT based on this paper: /path/to/paper.pdf
```

**Create slides with specific theme:**
```
Make a presentation about X using the Midnight theme
```

**Compile an existing script:**
```
Compile generate_slides.js
```

**Review a presentation:**
```
Review my slides: output.pptx
```

**Extract figures from a paper:**
```
Extract figures from /path/to/paper.pdf pages 3-5
```

## How It Works

### Creation Pipeline

```
Phase 0: Paper Analysis (extract key concepts, prerequisites, notation)
    ↓
Phase 1: Requirements Interview (duration, audience, scope, theme)
    ↓
Phase 2: Structure Plan (GATE — user approval required)
    ↓
Phase 3: Formula Preparation (LaTeX → OMML metadata + PNG@600DPI)
    ↓
Phase 3.5: Diagram Preparation (DOT/Mermaid/TikZ → SVG/PNG)
    ↓
Phase 4: Generate JS Script (PptxGenJS with helpers + self-review)
    ↓
Phase 5: QA Loop (execute → PDF → visual inspection → score → fix)
    ↓
Post-Creation Checklist (final gate)
```

### Formula Pipeline

```
formulas.json → render_latex.py → formulas/manifest.json
                                      ↓
                              image entries: PNG@600DPI
                              omml entries: metadata only
                                      ↓
              generate_slides.js → .pptx (with {{MATH:id}} placeholders)
                                      ↓
              inject_omml.py → .pptx (native OMML math injected)
```

### Diagram Pipeline

```
diagrams.json → render_diagrams.py → diagrams/manifest.json
                    ↓
            Graphviz → SVG (structural diagrams)
            Mermaid  → SVG (behavioral diagrams)
            TikZ     → SVG/PNG (math diagrams)
            Extract  → PNG@300DPI (paper figures)
```

## Color Themes

| Theme | Style | Best For |
|-------|-------|----------|
| **Academic Light** | White/light gray, clean | Research talks, seminars |
| **Midnight** | Dark throughout | Tech, premium feel |
| **Ocean** | Blue tones | Professional academic |
| **Forest** | Green natural tones | Biology, environmental |
| **Sandwich** | Dark title/conclusion + light content | Conference talks |

## File Structure

```
powerpoint-skill/
├── powerpoint-slides/
│   ├── SKILL.md                    # Main skill definition (29 rules, 9 actions)
│   ├── formula-rendering.md        # OMML + LaTeX image pipeline docs
│   ├── diagram-rendering.md        # 5-layer diagram pipeline docs
│   ├── pptxgenjs-reference.md      # PptxGenJS API + 9 layout code patterns
│   ├── references/
│   │   └── themes.md               # Theme JSON, script template, 17 helpers
│   └── scripts/
│       ├── render_latex.py          # Batch formula renderer (600 DPI + SVG)
│       ├── render_diagrams.py       # Batch diagram renderer (Graphviz/Mermaid/TikZ)
│       ├── inject_omml.py           # OMML post-processor (pandoc → OMML XML)
│       ├── check_overlaps.py         # Card-aware overlap & boundary checker
│       ├── soffice.py               # LibreOffice PPT→PDF wrapper
│       └── thumbnail.py             # Slide thumbnail grid generator
├── example/
│   ├── Bünz et al. - Linear-Time Permutation Check.pdf
│   └── linear_perm_slides/
│       ├── generate_slides.js       # Generated PptxGenJS script
│       ├── formulas.json            # 23 formulas (15 OMML + 8 image)
│       ├── diagrams.json            # 2 Graphviz diagrams
│       └── linear_perm_slides.pptx  # Output presentation (25 slides)
├── .gitignore
├── LICENSE
└── README.md
```

## Known Limitation: LibreOffice & OMML

OMML (Office MathML) is PowerPoint's native math format. LibreOffice does not fully support OMML rendering — when converting `.pptx` to PDF via LibreOffice (`soffice`), OMML formulas may appear blank or distorted. This is a LibreOffice limitation, not a bug in the generated file. The formulas display correctly in Microsoft PowerPoint and WPS Office.

The QA visual-check step uses LibreOffice for PDF conversion, so OMML formulas are expected to look wrong in the QA preview. Only image-rendered formulas (LaTeX → PNG) are checked for visual clarity during QA.

## Example

The `example/` directory contains a complete real-world example:

**Linear-Time Permutation Check** (Bünz, Chen, DeStefano, NYU 2025)
- 25-slide Academic Light theme presentation
- 23 formulas (15 native OMML + 8 LaTeX image)
- 2 Graphviz diagrams with theme color injection
- Source paper PDF included

## License

MIT
