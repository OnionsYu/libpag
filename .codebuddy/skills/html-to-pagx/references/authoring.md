# Authoring HTML for PAGX Conversion

A practical cheat-sheet for writing HTML that converts cleanly into PAGX. This is the
prescriptive companion to `spec/html_subset.md` (the authoritative contract — consult it for
exact rules, edge cases, and the full normalization-pass behaviour).

## Contents

- [Golden Rules](#golden-rules)
- [Document Skeleton](#document-skeleton)
- [Layout (Flexbox)](#layout-flexbox)
- [Sizing](#sizing)
- [Backgrounds, Borders, Shadows](#backgrounds-borders-shadows)
- [Gradients](#gradients)
- [Text](#text)
- [Images and Icons](#images-and-icons)
- [Do / Don't](#do--dont)

---

## Golden Rules

1. **Inline styles only.** Put every style in `style="…"` on the element. It has the highest
   precedence and is the most predictable to convert. A single `<head><style>` block with
   class/element selectors is allowed, but inline is preferred.
2. **Flexbox for layout, not margins or absolute positioning.** `padding`, `gap`, and
   `flex: N` map directly onto PAGX; margins and absolute offsets are supported but lossy.
3. **Inline `<svg>` for icons.** Never use text glyphs (`+`, `×`, `→`, `★`) as icons.
4. **One element = one box.** An element that has both a painted background and padded
   children becomes the canonical PAGX "outer background + inner padded container" pair
   automatically — you do not need to nest manually for that.
5. **Pixels.** Use px everywhere. `em`/`rem`/`pt`/`vw`/`vh` are auto-converted to px (with a
   `subset:unit-coerced` note); `calc()`, `var()`, `min()`, `max()`, `clamp()` are rejected.

---

## Document Skeleton

```html
<!DOCTYPE html>
<html>
  <head>
    <title>Canvas Title</title>
  </head>
  <body style="width: 800px; height: 600px; font-family: Arial; background-color: #0F172A;">
    <!-- visual content -->
  </body>
</html>
```

- `<html>` and `<body>` are required; the canvas size comes from `<body>` `width`/`height`.
- Allowed container tags (all map to a `<Layer>`, semantics-neutral): `div`, `section`,
  `header`, `footer`, `main`, `aside`, `nav`, `article`.
- Text tags: `p`, `h1`–`h6`, `span`, `a`. Line break: `<br/>` (must be self-closed).
- All tags lower-case and properly closed (`<img/>`, `<br/>`).

---

## Layout (Flexbox)

| Goal | CSS | Notes |
|------|-----|-------|
| Enable flex | `display: flex` | Required to activate layout on the container |
| Direction | `flex-direction: row \| column` | Default is `row` |
| Spacing between children | `gap: 12px` | Single value only |
| Inner padding | `padding: 12px` / `12px 16px` / `T R B L` | |
| Cross-axis align | `align-items: stretch \| center \| flex-start \| flex-end` | |
| Main-axis distribute | `justify-content: flex-start \| center \| flex-end \| space-between \| space-evenly \| space-around` | |
| Grow to fill | `flex: 1` (or any integer) | Grow only; don't also set a main-axis size |

Block (`display: block`, the default) stacks children vertically — fine for simple columns,
but `display: flex; flex-direction: column; gap` gives explicit spacing control.

**Avoid**: `flex-wrap`, `flex-grow`, `flex-basis` (use the `flex: N` shorthand), `display:
grid`, `grid-*`, `float`, `order`, `align-content`, `align-self`, `*-reverse` directions —
all dropped with a warning.

---

## Sizing

- `width: 240px` / `height: 64px` → fixed px size.
- `width: 50%` / `height: 100%` → percentage of the parent's content box.
- `box-sizing: border-box` is the default and only behaviour.
- **Don't** set explicit main-axis width/height on a `flex: N` child — let flex distribute.
- `min-*`, `max-*`, `aspect-ratio` are ignored (warning).

---

## Backgrounds, Borders, Shadows

| CSS | PAGX result |
|-----|-------------|
| `background-color: #1E293B` | filled rectangle behind content |
| `border-radius: 12px` (px or `%` of `min(w,h)`; `50%` on a fixed square → circle) | rounded corners |
| `border: 1px solid #334155` (also `dashed` / `dotted`) | inside stroke |
| `box-shadow: 0 2px 6px #00000026` (repeatable; `inset` → inner shadow) | drop / inner shadow |
| `opacity: 0.8` | layer alpha |
| `overflow: hidden` | clip to bounds (rectangular only) |
| `filter: blur(4px)` / `backdrop-filter: blur(8px)` | blur filter / background blur |
| `transform: rotate(8deg)` (single function or `matrix()`) | layer matrix |

**Avoid**: per-side `border-top/right/bottom/left`, per-corner `border-*-radius`,
`background-image: url(...)` (use `<img>`), `background-size/repeat/position`, `outline`,
`clip-path`, compound/3D transforms — all dropped with a warning.

For a **rounded avatar**, wrap one `<img>` that exactly fills a `border-radius` +
`overflow: hidden` container; the importer folds it into a clipped rounded image (see
`spec/html_subset.md` §7.1).

---

## Gradients

Set as `background-image`:

```css
background-image: linear-gradient(135deg, #6366F1 0%, #8B5CF6 100%);
background-image: radial-gradient(circle at 50% 0%, #1E293B, #0F172A);
background-image: conic-gradient(from 0deg, #F00, #0F0, #00F);
```

To paint a gradient onto **text**, combine a gradient `background-image` with
`background-clip: text` (or `-webkit-background-clip: text`) on the text element.

---

## Text

| CSS | Effect |
|-----|--------|
| `color: #E2E8F0` | text color |
| `font-family: Arial` | font family (must be installed on the render host, or download it — see pipeline.md) |
| `font-size: 16px` | font size |
| `font-weight: bold` (or `700`+) | bold |
| `font-style: italic` | italic |
| `letter-spacing: 0.5px` | tracking |
| `line-height: 1.4` (unitless), `22px`, or `140%` | line height |
| `text-align: left \| center \| right \| justify` | alignment |
| `text-decoration: underline \| line-through` | decoration overlay |
| `white-space: nowrap` | disable wrapping |

Default font sizes: `h1` 32, `h2` 24, `h3` 20, `h4` 18, `h5` 16, `h6` 14, body 14px.

**Avoid**: `text-transform`, `text-indent`, `word-spacing`, `font` shorthand, `@font-face`
web fonts, `text-overflow: ellipsis` — dropped with a warning.

---

## Images and Icons

```html
<!-- raster image; src may be a data URI or a path relative to the HTML file -->
<img src="avatar.png" width="64" height="64" style="object-fit: cover;"/>

<!-- icon: inline SVG, NOT a text glyph -->
<span style="color: #38BDF8;">
  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
    <path d="M5 12h14M12 5l7 7-7 7"/>
  </svg>
</span>
```

- `object-fit`: `fill`→Stretch, `contain`→LetterBox, `cover`→Zoom (default Stretch).
- Inline `<svg>` is captured verbatim and expanded by `pagx resolve` / `pagx verify`.
- `currentColor` inside an SVG inherits the element's `color`, so color overrides work.
- `<img src="icon.svg"/>` becomes an external import directive.

---

## Do / Don't

**Do**
- Author one self-contained HTML file with inline styles and explicit px on `<body>`.
- Lay out with flex (`gap`, `padding`, `flex: N`).
- Use inline `<svg>` for every icon.
- Give meaningful sections an `id` for scoped `pagx verify --id`.
- Read importer warnings and fix the HTML until conversion is clean.

**Don't**
- Don't use `grid`, `float`, `position: absolute` for primary layout, or margins where
  `gap`/`padding` would do.
- Don't use `calc()` / `var()` / `min/max/clamp()`.
- Don't use text characters as icons.
- Don't rely on `<table>`, `<form>`, `<input>`, `<button>`, `<canvas>`, `<script>` — they
  are rejected (the snapshot route rewrites some of these; the direct route does not).
- Don't assume a web font is available at render time — install or download it.
