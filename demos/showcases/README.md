# Showcase decks

Two full decks, generated end to end by `@shbernal/ts-pptx`. No slide here was touched in
PowerPoint.

```bash
pnpm demos:build                    # both
pnpm demos:build field-notes        # one, by slug
```

Output goes to `output/` (git-ignored).

## The decks

### `quarterly-review/` — Kestrel Q3 FY26 Business Review

Eleven slides. The corporate flagship: a themed `<a:clrScheme>`, five slide masters, native
linear gradients on the cover and closing, KPI cards assembled as groups, a stacked column
chart, a doughnut with a text well in its hole, a line chart with a callout, a hand-styled
table with a totals row, chevron timeline, and speaker notes throughout.

It imports nothing from `node:` — every mark on every slide is drawn rather than loaded.
That is what lets the site's demos page (`www/demos/`) import this same module and build the
identical deck in a browser.

### `field-notes/` — Four Cities After Dark

Eight slides. The visual flagship: full-bleed photography, gradient scrims over images (the
standard editorial fix for putting white type on an unpredictable photo), a duotone picture
effect, a three-up image grid, an embedded video with a poster frame, and a radial-gradient
colophon carrying a real hyperlink relationship.

Node-only by nature — it loads photographs and a video from `demos/common` by path.

## Layout

```
lib/assets.mjs      absolute asset + output paths, and the one base64 helper addMedia needs
lib/layout.mjs      slide geometry: the 16:9 box, margins, column arithmetic
lib/showcases.mjs   the SHOWCASES registry — every deck, in build order
<deck>/design.mjs   palette, type scale, theme, masters — no slide names a raw hex
<deck>/data.mjs     content, kept apart from layout (quarterly review only; Field Notes
                    carries its handful of captions inline)
<deck>/index.mjs    the slides, plus an exported `showcase` descriptor
build.mjs           the runner; the only place that touches the filesystem
```

Each deck exports `{ slug, title, description, fileName, build }`. Adding a third deck means
writing that object and adding it to `SHOWCASES` in `lib/showcases.mjs`.

That registry is **not** in `build.mjs`, and the distinction is the whole reason it was
moved out: `scripts/byte-identity.mjs` enumerates the decks from it too. A deck registered
anywhere else still builds, and is silently absent from the gate that would have caught an
emitter regression in it.

## Two things worth knowing before editing a deck

**`addChart` mutates the series you hand it.** It normalizes `labels` from `string[]` to
`string[][]` **in place**, so an array passed to a chart is not the array you passed in.
Iterating it afterwards to build a legend yields one nested array instead of three strings.
`quarterly-review/data.mjs` keeps a plain `SEGMENTS` source of truth and derives the chart
shape from it; do the same rather than reusing a chart's arrays.

**`addMedia`'s `cover` takes base64, not a path** — unlike `addImage`, which takes either.
`lib/assets.mjs` exports `imageDataUri()` for exactly that one call site.
