# LZ-Protocol — a shared object protocol for PPTX ⇄ Google Slides

**Problem.** When a `.pptx` is imported into Google Slides, every shape's
`objectId` and OOXML `name` are **discarded and regenerated**. So a python-pptx
generator and the lizard Apps-Script module cannot address the same object by ID.
lizard's native recognition (objectId prefixes `tab_`, `progress_`, `page_num_`,
…) can therefore never match an imported shape.

**The one field that survives the import boundary is alt-text.** OOXML
`<p:cNvPr title="…" descr="…">` maps to Google Slides `PageElement.getTitle()` /
`getDescription()`, both readable/writable on either side. lizard already uses
`getTitle()` ad-hoc (`PROGRESS`, `PROGRESS_BG`, `MAIN_TITLE`, `PREVIOUS_TITLE`,
`GRID_N`). LZ-Protocol **formalizes alt-text as a namespaced identity channel**
so both tools share one vocabulary.

## The contract

Every object carries a **self-describing style instruction** in its alt-text:

| Field | Carries |
|---|---|
| `title` | the **ROLE** (uppercase enum) — quick exact-match, back-compatible |
| `description` | a **TOML instruction block** with role + every cosmetic setting |

Example (`description` of a BODY element, points + hex):

```toml
[lz]
role = "body"
font = "Source Sans Pro"
x = 24.48
y = 90.72
w = 671.04
h = 287.28
size = 15
bold = false
color = "#212121"
```

Because the full instruction travels WITH each element, **lizard is a dumb
interpreter** — it needs no synced copy of the spec. The single source of truth
is `blood_school_style/spec_source.py`; `lz_protocol.to_toml()` injects it into
each element at generation time. `lzParseInstr()` / `lzInstr()` read it back.

## Role vocabulary

Three lifecycles: **managed** chrome (lizard delete-rebuilds), **styled** content
(lizard catches & re-applies style, never deletes), and the **marker**.

| ROLE | Element | Lifecycle |
|---|---|---|
| `PROGRESS` / `PROGRESS_BG` | progress-bar fill / track | managed |
| `TAB` | breadcrumb nav tab | managed |
| `PAGE_NUM` | page number `i / N` | managed |
| `SECTION_BOX` / `SECTION_LABEL` / `OUTLINE` | section-page mini-TOC / chip / outline | managed |
| `MAIN_TITLE` | footer running title | managed |
| `SECTION` | marker on a section slide — makes it a section boundary | **authored** — never touched |
| `TITLE` / `TITLE_MAIN` / `SUBTITLE` | slide title bar / title-slide headline / subtitle | **styled** |
| `DATE` / `EMAIL` / `BRAND_CHIP` | title-slide bits | styled |
| `BODY` | bullet body | styled |
| `TABLE` | table frame | styled |
| `COL_LEFT` / `COL_RIGHT` / `COL_HEAD_L` / `COL_HEAD_R` | two-column bodies / heads | styled |
| `KEY_HEADLINE` / `KEY_SUB` / `KEY_POINTS` | keypoints headline / sub / list | styled |

- `LZ_MANAGED_ROLES` — chrome lizard owns: deletes any tagged instance (regardless
  of objectId) and rebuilds from slide order.
- `LZ_STYLED_ROLES` — content lizard **catches and re-applies the canonical font +
  role style to**, but never deletes. This is how a "foreign" PPTX-imported deck
  gets the house style online — no font embedding needed.
- `SECTION` — authored marker lizard reads but never removes.

## Workflow

```
collect info → build content locally in pptx + INJECT instructions (rough layout ok)
            → import to Google Slides → run one lizard command → cosmetics done
```

`src/protocol/lz_apply_style.js` + menu **⚙ 設定與批次 → 🦎 套用 PPTX 匯入樣式 (LZ)**:

1. `lzApplyStyleAll()` — walk every slide, read each element's `lzInstr()`, and
   apply it online: **font, size, bold, italic, color, geometry (setLeft/Top/
   Width/Height), fill, vertical anchor**, plus horizontal alignment in one
   Advanced-Slides batch. Tables (`getTables()`) get geometry + per-row header/
   cell styling. Managed chrome is skipped here (rebuilt in step 2). Foreign fonts
   from the import are overwritten — **no font embedding needed**.
2. `lzApplyAll()` — the above, then `runAllFunctionsUltraMegaBatch()` to rebuild
   all managed chrome from live slide order.

Because instructions are self-contained, changing the house style means editing
only `spec_source.py` and regenerating the deck — lizard needs no update.

### Citation block

`CITATION` is a styled role for reference/source small-print pinned to the slide
foot. python: `d.standard(..., cite="Tse E, Kwong YL. Blood 2013;121:4997.")`
(also on `table` / `two_column`). lizard styles it from its injected instruction
like any other content element.

## How the two sides use it

**python-pptx (producer).** `lz_protocol.tag(shape, role, **data)` writes
`title=ROLE` and `descr=JSON`. Generators stamp every chrome element and every
section-slide title.

**lizard (consumer).** `lzRoleOf(el)` returns the role from `getTitle()` (or JSON
in `getDescription()`). Two hook points:
1. `batchDeleteAllElements` also deletes any shape whose role ∈ `LZ_MANAGED_ROLES`
   — so imported, tag-carrying chrome is reclaimed even though its objectId is
   Google-generated. **This is what makes the round-trip work.**
2. `getSectionHeadersUltra` also treats a slide as a section when it holds a
   `SECTION`-tagged shape — decoupling section detection from layout naming
   (`SECTION_HEADER`), which is fragile across import.

## Why this is bidirectional

- **pptx → lizard:** python stamps roles; lizard recognizes, dedupes, rebuilds.
- **lizard → pptx:** lizard stamps the same roles (`setTitle(role)`) when it mints
  chrome; export back to `.pptx` and python-pptx reads the tags to know each
  object's role — no guessing by position or text.

Version: `v:1`. Bump `v` in the JSON payload on breaking vocabulary changes.
