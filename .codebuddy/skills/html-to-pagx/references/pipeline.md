# Conversion Pipeline

How to turn HTML into a rendered PAGX, which route to pick, and how to read the importer's
diagnostics. The two routes share the same importer; they differ only in whether a headless
browser flattens the HTML first.

## Contents

- [Prerequisites](#prerequisites)
- [Route A: Direct Import](#route-a-direct-import-subset-html)
- [Route B: Snapshot Import](#route-b-snapshot-import-rich--js--url)
- [Web Fonts](#web-fonts)
- [External Images](#external-images)
- [Diagnostics](#diagnostics)
- [Troubleshooting](#troubleshooting)

---

## Prerequisites

- **`pagx` CLI** (both routes). Install/update per the setup block in `SKILL.md`.
- **Snapshot route only**: `node`, plus the `tools/html-snapshot` dependencies:
  ```bash
  cd tools/html-snapshot && npm install   # ~150 MB Chromium on first run
  ```
  If Chromium is missing from the default cache after a sandboxed install:
  ```bash
  PUPPETEER_CACHE_DIR="$HOME/.cache/puppeteer" \
    npx --prefix tools/html-snapshot puppeteer browsers install chrome
  ```

---

## Route A: Direct Import (subset HTML)

Use when you authored the HTML yourself in the subset (see `references/authoring.md`). No
browser involved — fast and deterministic.

```bash
pagx import --input design.html --output design.pagx   # extension/content infers format
pagx import --input design.html --format html           # force HTML; → design.pagx
pagx verify design.pagx                                  # resolve + check + render → design.png
```

Useful import flags:

| Flag | Use |
|------|-----|
| `--format html` | Force HTML parsing when the extension is ambiguous |
| `--html-strict` | Treat any subset warning as a hard error (CI gate) |
| `--html-preserve-unknown` | Keep unknown tags as empty `data-html-unknown` Layers (debug) |
| `--html-no-normalize` | Skip the subset normalizer; importer sees the raw DOM (debug only) |

The importer auto-normalizes the input (resolves the `<style>` cascade, converts units,
drops/downgrades non-subset properties) before traversing. See `spec/html_subset.md` §11 for
the exact pass list.

---

## Route B: Snapshot Import (rich / JS / URL)

Use for a live URL, or an existing React/Tailwind/JS page, or any HTML that uses non-subset
CSS. A headless browser renders the page and emits a flat, absolute-positioned subset HTML,
which is then imported.

**One-shot wrapper** (`tools/html-snapshot/html2pagx`) — snapshot → import → resolve → render:

```bash
tools/html-snapshot/html2pagx page.html
# → page.subset.html   (flat subset HTML; the snapshot)
# → page.pagx          (after import + resolve)
# → page.png           (rendered at scale 1)

# URL input: --output-name is required (no filesystem basename)
tools/html-snapshot/html2pagx https://example.com/demo --output-name demo --output-dir out/
```

Common `html2pagx` flags:

| Flag | Use |
|------|-----|
| `-o, --output-dir <dir>` / `--output-name <name>` | Where/what to name outputs (`--output-name` required for URLs) |
| `--scale <N>` | PNG render scale (default 1; use 2 for detail) |
| `--no-render` / `--no-resolve` | Stop after resolve / after import |
| `--download-fonts` / `--embed-fonts` | Capture (and optionally embed) the page's web fonts — see below |
| `--download-images` | Reference external images by local path instead of inlining base64 |
| `--viewport-width/-height`, `--wait-ms`, `--selector` | Forwarded to the snapshot step |
| `--browser-engine puppeteer\|playwright` | Headless driver |

**Equivalent via the `pagx` CLI** (it shells out to the repo's `snapshot.js`):

```bash
pagx import --input page.html --html-snapshot --html-infer-flex --output page.pagx
pagx import --input https://example.com/demo --html-snapshot --output demo.pagx
pagx verify page.pagx
```

`--html-infer-flex` recovers `display: flex` semantics (gap/padding/alignment) from the
flat absolute-positioned snapshot. It is a lossy heuristic — opt in, then verify the result.

---

## Web Fonts

The snapshot bakes each text node's `font-family` **name** into the style, but PAGX resolves
fonts by name from the **render host's installed fonts**. A page using an uninstalled web
font renders with a wrong fallback. Options:

- `html2pagx --download-fonts` — download the faces Chromium fetched and register them as
  render fallbacks (not embedded into the `.pagx`).
- `html2pagx --embed-fonts` — additionally embed them so the `.pagx` is self-contained.
- Manual: `pagx font embed page.pagx --fallback page.fonts/*.ttf`, then
  `pagx render page.pagx -o page.png --fallback page.fonts/*.ttf`.

CJK families served as unicode-range subsets carry unreliable weight metadata; when weight
fidelity matters, install the real font on the render host instead.

---

## External Images

By default the snapshot inlines external images as base64 data URIs (self-contained but
large). For image-heavy pages, `html2pagx --download-images` writes them to disk
(`<output>.images/`) and references them by local path, keeping the `.pagx` small. Relative
`<img src>` paths in directly-authored HTML resolve against the HTML file's directory.

---

## Diagnostics

The importer reports issues on stderr under the `subset:<category>` namespace. Severity:

- **error** — structural; no document produced (no `<body>`, no canvas size). Fix before
  anything else.
- **warning** — an element or property was skipped or downgraded. The design may not look
  right where it depended on that property.
- **note** — informational (e.g. a defaulted font-size, a `subset:unit-coerced` px conversion).

**Always read the warnings after every import.** For each one, change the HTML so the design
no longer relies on the dropped construct (consult `references/authoring.md` for the
supported alternative). Common categories:

| Code | Meaning | Fix |
|------|---------|-----|
| `subset:unsupported-property` | A CSS property/function isn't in the subset | Replace with a supported property; remove `calc()`/`var()` etc. |
| `subset:unsupported-tag` | An element isn't allowed (`<table>`, `<input>`, …) | Restructure with `<div>`/`<span>`; the snapshot route rewrites some inputs automatically |
| `subset:unsupported-selector` / `-at-rule` | A `<style>` selector or `@media`/`@font-face` was ignored | Move styles inline; avoid descendant/pseudo selectors |
| `subset:unit-coerced` | `em/rem/pt/vw/vh` converted to px (note) | Usually fine; switch to px to silence |
| `subset:flex-inferred` / `-inference-skipped` | `--html-infer-flex` did (or couldn't) recover flex | Verify layout; tweak the source spacing if skipped |

`--html-strict` turns the first warning into a hard error — use it once the input is clean to
guard against regressions.

---

## Troubleshooting

| Symptom | Cause | Fix |
|---------|-------|-----|
| `failed to determine canvas size` | No `width`/`height` on `<body>` | Add `style="width:…px; height:…px;"` to `<body>` |
| Imported PAGX is empty | Source needs JS to render and you used the direct route | Use the snapshot route (`--html-snapshot` / `html2pagx`) |
| Icons render as `□` or are missing | Text-glyph icons, or an uninstalled icon font | Use inline `<svg>`; the snapshot route inlines icon-font glyphs automatically |
| Text in the wrong typeface | Web font not installed on the render host | `--download-fonts` / `--embed-fonts`, or install the font |
| Layout collapsed / overlapping after snapshot | Flex not recovered from absolute output | Add `--html-infer-flex`; if still wrong, simplify the source layout |
| Snapshot fails to launch | Chromium not installed | Run the puppeteer browser install command in §Prerequisites |
| Many `subset:*` warnings | HTML uses non-subset CSS | Rewrite per `references/authoring.md`, or route through the snapshot which flattens computed styles |
