---
name: html-to-pagx
description: >-
  Generates HTML in the PAGX-importable subset for a visual design, then converts it to a
  rendered PAGX file. Use when the user asks to turn a webpage, URL, screenshot, mockup, or
  design description into PAGX, to convert/import HTML into PAGX, or to "generate HTML and
  make it a PAGX". For authoring or editing PAGX XML directly, use the `pagx` skill instead.
---

# HTML to PAGX Skill

Generate HTML for a design, then convert it to PAGX through the HTML importer. HTML is a
fast, natural intermediate format for laying out a design; the importer maps each allowed
construct onto a single PAGX equivalent with near-lossless fidelity.

## Reference Lookup

| Reference | Content | Loading |
|-----------|---------|---------|
| `references/authoring.md` | Practical cheat-sheet for writing convertible HTML (layout, color, text, images, do/don't) | Read before generating HTML |
| `references/pipeline.md` | Conversion commands, flag selection (snapshot vs direct), fonts/images, diagnostics, troubleshooting | Read before converting |
| `spec/html_subset.md` (repo root) | Authoritative HTML/CSS subset contract — every allowed element, property, and normalization pass | As needed for exact rules |

After conversion, refine the resulting `.pagx` with the **`pagx` skill** (verify, layout
inspection, visual QA). This skill produces the PAGX; the `pagx` skill polishes it.

---

## CLI Setup

Both routes need the `pagx` CLI. Ensure it is installed and current:

```bash
PAGX_MIN="0.3.22"
if ! command -v pagx &>/dev/null; then
  npm install -g @libpag/pagx
elif [ "$(printf '%s\n' "$PAGX_MIN" "$(pagx -v | awk '{print $2}')" | sort -V | head -1)" != "$PAGX_MIN" ]; then
  npm update -g @libpag/pagx
fi
```

The **snapshot route** (rich/JS pages and URLs) additionally needs `node` and the
`tools/html-snapshot` dependencies installed (`npm install` in that directory, ~150 MB
Chromium download on first run). The **direct route** (subset HTML you authored) needs only
the `pagx` CLI. See `references/pipeline.md` for details.

---

## Choosing a Route

| Input | Route | Command |
|-------|-------|---------|
| HTML you author yourself from a design/description | **Direct import** (no browser) | `pagx import --input design.html --output design.pagx` |
| A live URL, or a JS/React/Tailwind page, or any HTML using non-subset CSS | **Snapshot import** (headless browser flattens it first) | `pagx import --input page.html --html-snapshot --html-infer-flex --output page.pagx` |

**Default to the direct route** when you are generating the HTML yourself: author it in the
subset (per `references/authoring.md`) and skip the browser entirely — it is faster,
deterministic, and needs no Chromium. Use the snapshot route only when the source is an
existing rich page or URL you do not control. `tools/html-snapshot/html2pagx <input>` is a
one-shot wrapper for the snapshot route (snapshot → import → resolve → render); see
`references/pipeline.md`.

---

## Generation Workflow (Direct Route)

**When**: The user describes a design, shows a reference image/mockup, or asks to build a
PAGX, and you are authoring the HTML.

At the start, create a task list: one task for Step 2 (HTML draft) and one task per major
iteration of Step 3–4. Mark each in-progress before starting and completed after its checks
pass. Do NOT start the next task until the current one passes.

### Step 1: Assess

1. Clarify requirements — ask the user if canvas size, visual style, text content, or color
   scheme is unclear or ambiguous.
2. Establish a style sheet — color palette, spacing scale, roundness, font hierarchy.
3. Decompose the visual into a **containment tree** — containers and their direct children.
   Elements described within the same block belong to that container, not as siblings. This
   tree maps directly to the nested `<div>` structure in Step 2.

**Forbidden**: Do NOT write any HTML in this step.

### Step 2: Author HTML

Read `references/authoring.md` first. Write a complete subset HTML document:

- Set the canvas size on `<body style="width: …px; height: …px;">`.
- Use **inline `style="…"`** on every element (highest precedence, most predictable).
- Lay out with **flexbox** (`display: flex`, `flex-direction`, `gap`, `padding`,
  `align-items`, `justify-content`, `flex: N`) — not margins or absolute positioning.
- Use **inline `<svg>`** for icons. NEVER use text characters (`+`, `×`, `→`) as icons.
- Give meaningful sections an `id` (propagated to the PAGX Layer `id` for scoped verify).

### Step 3: Convert

```bash
pagx import --input design.html --output design.pagx
```

**Checks**:
1. **Read every warning** the importer prints. Each `subset:<category>` warning means a
   property/element was dropped or downgraded — fix the HTML so the design does not depend
   on it (see `references/pipeline.md` §Diagnostics). Re-run until the import is clean (or
   every remaining warning is understood and harmless).
2. Resolve and render to inspect the result:
   ```bash
   pagx verify design.pagx     # resolves <svg>/import, runs checks, writes design.png + design.layout.xml
   ```
   **ALL `pagx verify` diagnostics MUST be fixed** — re-run until exit code 0 with no
   diagnostic output.

### Step 4: Visual QA and Iterate

1. Read `design.png` and compare against the design intent — check sizes, spacing, colors,
   font sizes, text content, icons, and alignment.
2. For any mismatch, decide **where to fix it**:
   - **Layout/content/style issue** → edit `design.html` and re-run Step 3.
   - **Fine PAGX-level polish** (effects, precise geometry the subset can't express) → hand
     the `.pagx` to the **`pagx` skill** and edit the PAGX directly.
3. Repeat until the render matches the intent.

**Final check**: run `pagx verify design.pagx` one last time; the task is complete only when
it exits with code 0 and no diagnostic output. Keep `design.png` for reference; delete
`design.layout.xml` and any scoped `{id}` artifacts.

---

## Snapshot Route (URL / rich page)

When the input is a URL or an existing JS/React/Tailwind page you do not author:

```bash
# One-shot wrapper: snapshot → import → resolve → render
tools/html-snapshot/html2pagx page.html                 # → page.subset.html, page.pagx, page.png
tools/html-snapshot/html2pagx https://example.com/demo --output-name demo --output-dir out/

# Or via the pagx CLI directly
pagx import --input page.html --html-snapshot --html-infer-flex --output page.pagx
pagx verify page.pagx
```

Then run Step 4 (Visual QA): inspect the PNG, and for residual issues edit the source HTML
(if you control it) or polish the `.pagx` with the `pagx` skill. For web fonts and external
images, see `references/pipeline.md` (`--download-fonts`, `--download-images`).
