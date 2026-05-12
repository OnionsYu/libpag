# HTML → PAGX Subset

The `pagx import` CLI ships an HTML→PAGX converter. It accepts a deliberately
narrow flex-based XHTML subset that maps cleanly onto PAGX's container layout
model (Layer + `layout`/`gap`/`padding`/`flex`/`alignment`/`arrangement`).

The conversion is **lossy**: anything outside the subset is dropped or
approximated and the converter appends a warning to the document. Use this
format only as a fast first-draft path; the canonical artifact for review and
publish is the `.pagx` file.

## CLI

```bash
pagx import --input draft.html --output ui.pagx
pagx import --format html --input fragment.txt --output ui.pagx
```

Options:

- `--html-preserve-unknown` — keep unsupported tags as empty Layers tagged with
  `customData["html-tag"]` instead of dropping them.

## Hard rules for AI authors

To make conversion deterministic:

1. **Well-formed XHTML only**. Lowercase tags, all attributes quoted, void
   elements self-closing (`<br/>`, `<img .../>`).
2. **One root `<html>` with explicit canvas size in the inline style.**
   Example: `<html style="width: 320px; height: 240px">`. Without this, the
   converter falls back to `800×600` and emits a warning.
3. **Inline `style="…"` only.** No `<style>` blocks, no class selectors, no
   media queries, no pseudo-elements.
4. **Use flex containers.** Block flow falls back to a vertical stack, but
   prefer `display: flex; flex-direction: row|column; gap: …` for explicit
   intent.
5. **No CSS Grid, float, position-relative reordering, animations, transitions.**
6. **Lengths in `px` or `%` only.** `em`/`rem`/`vh`/`vw`/`ch` are dropped with
   a warning.
7. **Keep text styling on the parent block element**, not on inline `<span>` /
   `<strong>` / `<em>` etc. Inline tags are flattened to their text content;
   their style (color, font-weight, font-size) is **lost**.
8. **`<br/>` is the only way to keep a hard line break inside a paragraph.**
   Whitespace and newlines in raw text are collapsed per CSS rules.

## Element mapping

| HTML element                               | PAGX result |
|--------------------------------------------|-------------|
| `<html>`                                   | Document canvas (size from `style.width/height`) |
| `<body>`                                   | Top-level Layer with `left=right=top=bottom=0` |
| `<div>`, `<section>`, `<article>`, `<header>`, `<footer>`, `<nav>`, `<main>`, `<aside>`, `<button>`, `<a>`, `<ul>`, `<ol>`, `<li>`, `<form>`, `<label>`, `<figure>`, `<figcaption>` | Layer (block container) |
| `<p>`, `<h1`–`h6>`, `<blockquote>`, `<pre>`, `<code>` | Layer with one TextBox + Text child |
| `<span>`, `<strong>`, `<b>`, `<em>`, `<i>`, `<u>`, `<small>`, `<mark>` | Flattened into surrounding text run |
| `<br/>`                                    | Newline character in surrounding text |
| `<img src="…">`                            | Layer with Rectangle + ImagePattern fill |
| `<svg>` (inline)                           | Layer with `importDirective.content` (resolve via `pagx resolve`) |
| `<head>`, `<title>`, `<meta>`, `<link>`, `<style>`, `<script>`, `<noscript>` | Dropped silently |
| Anything else                              | Dropped (warning), or kept as a tagged Layer with `--html-preserve-unknown` |

## CSS property mapping

### Container layout

| CSS                              | PAGX |
|----------------------------------|------|
| `display: flex`                  | Enables flex; default `layout = "horizontal"` |
| `display: flex; flex-direction: column` | `layout = "vertical"` |
| `display: none`                  | `visible = false` |
| `flex-direction: row\|column`    | `layout = "horizontal"\|"vertical"` |
| `gap: Npx`                       | `gap = N` |
| `padding: shorthand`             | `padding` (1, 2, 3, or 4 px values) |
| `padding-top/right/bottom/left`  | individual sides |
| `align-items: stretch\|center\|flex-start\|flex-end\|start\|end` | `alignment` |
| `justify-content: flex-start\|center\|flex-end\|space-between\|space-around\|space-evenly` | `arrangement` |
| `flex: N` or `flex-grow: N`      | `flex = N` |
| `width: Npx` / `width: N%`       | `width` / `percentWidth` |
| `height: Npx` / `height: N%`     | `height` / `percentHeight` |

`row-reverse` / `column-reverse` map to the forward direction with a warning.

### Background, border, shadow

| CSS                              | PAGX |
|----------------------------------|------|
| `background-color: <color>`      | Inserts a stretch-fill `Rectangle` + `Fill` |
| `border-radius: Npx`             | Sets `roundness` on the background Rectangle (or appends a `RoundCorner` modifier if no background) |
| `border: Npx solid <color>`      | Adds a `Stroke` painter to the background Rectangle |
| `box-shadow: H V B <color>`      | `DropShadowStyle{offsetX=H, offsetY=V, blurX=blurY=B, color}` (`inset` is currently unsupported) |
| `opacity: A`                     | `alpha = A` |
| `overflow: hidden`               | `clipToBounds = true` |
| `position: absolute\|fixed`      | `includeInLayout = false` |
| `left/right/top/bottom: Npx`     | LayoutNode constraint of the same name |

When background and padding coexist on the same element, the converter emits a
nested wrapper (outer Layer paints background + border + shadow, inner
`name="content"` Layer carries padding/gap/flex/children). This avoids the
PAGX `stretch-fill element affected by padding` diagnostic.

### Color

| CSS                              | PAGX |
|----------------------------------|------|
| `#RGB`, `#RRGGBB`, `#RRGGBBAA`   | `Color` (sRGB) |
| `#RGBA` (4-digit)                | `Color` (sRGB, with alpha) |
| `rgb(r, g, b)` / `rgba(r, g, b, a)` | `Color` (sRGB, components 0–255 or `%`) |
| `rgb(r g b / a)` (CSS 4 syntax)  | `Color` |
| Named CSS color (full CSS3 list + `rebeccapurple`) | `Color` |
| Anything else                    | Warning, no fill emitted |

### Text

| CSS                                            | PAGX |
|------------------------------------------------|------|
| `color: <color>`                               | `Fill` painted after the TextBox glyphs |
| `font-family: family[, fallback...]`           | First family → `Text.fontFamily` |
| `font-size: Npx`                               | `Text.fontSize` |
| `font-weight: bold\|≥600\|<600\|normal`        | `Text.fauxBold` |
| `font-style: italic\|oblique`                  | `Text.fauxItalic` |
| `letter-spacing: Npx`                          | `Text.letterSpacing` |
| `line-height: Npx` or unitless multiplier      | `TextBox.lineHeight` |
| `text-align: left\|center\|right\|justify\|start\|end` | `TextBox.textAlign` |

`<h1>`–`<h6>` get the browser-style font size + bold defaults baked in before
inline style is applied, so explicit `style="…"` overrides still win.

## Things deliberately not supported

- CSS Grid (`display: grid`, `grid-template-*`) — flatten to nested flex.
- Floats, multi-column, table layout.
- Pseudo-elements `::before`/`::after`.
- Media queries, container queries.
- Animations, transitions, `transform` (not yet — coming later).
- `z-index` reordering (PAGX uses document order).
- Font weight expressed as a custom font style name — only bold vs not is preserved
  via `fauxBold`. To use a real condensed/light/etc. style, edit the resulting
  PAGX directly.

## Quick template for AI to emit

```html
<html style="width: 360px; height: 540px">
  <body style="background-color: #FFFFFF; padding: 16px;
               font-family: Inter; color: #1F2937">
    <div style="display: flex; flex-direction: column; gap: 12px">
      <h2 style="font-size: 20px">Title</h2>
      <p style="font-size: 14px; line-height: 20px">
        Description text. Avoid styling spans inside this paragraph.
      </p>
      <div style="display: flex; flex-direction: row; gap: 8px;
                  justify-content: flex-end">
        <div style="background-color: #2563EB; padding: 8px 12px;
                    border-radius: 6px; color: #FFFFFF; font-size: 13px">
          Confirm
        </div>
      </div>
    </div>
  </body>
</html>
```

This template converts cleanly with **zero** `pagx verify` diagnostics.
