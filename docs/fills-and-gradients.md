---
doc-schema-version: 1
title: "Fills and gradients"
summary: "Solid, gradient, pattern and picture fills on shapes, backgrounds, tables, charts and lines: where each kind applies, linear and radial gradients, and the difference between no fill, inherit and leaving the fill out."
read_when:
  - Filling a shape, slide background, table or chart area with a gradient, pattern or picture
  - Working out why a shape has no fill, or takes one from its style
  - Choosing between type 'none' and 'inherit'
doc_type: "guide"
---

# Fills and gradients

A fill paints an area: a shape's interior, a slide's background, a table cell, a chart area. The same
options work in every one of them:

```ts
slide.addShape("rect", {
  x: 1, y: 1, w: 4, h: 2,
  fill: {
    type: "gradient",
    gradient: {
      kind: "linear",
      angle: 90,
      stops: [
        { position: 0, color: "1F3A5F" },
        { position: 100, color: "4A8FD6" },
      ],
    },
  },
})
```

## Where each kind applies

| Target | Option | Solid | Gradient | Pattern | Picture |
| --- | --- | --- | --- | --- | --- |
| Shape, text box | `fill` | yes | yes | yes | yes, raster |
| Slide or master background | `background` | yes | yes | yes | through `path` or `data` |
| Table, table cell | `tableFill`, `fill`, a cell's `options.fill` | yes | yes | yes | yes, raster |
| Chart area, plot area | `chartArea.fill`, `plotArea.fill` | yes | yes | yes | no, see [invalid input](#invalid-input) |
| Line | `line` | yes | yes | yes | no, see [invalid input](#invalid-input) |

A colour string on its own is a solid fill, so `fill: "FF0000"` and `fill: { color: "FF0000" }` are the
same. A line is always an object: `line: { color: "FF0000" }`.

## Options at a glance

| Option | Type | Default | Effect |
| --- | --- | --- | --- |
| `type` | `'solid'`, `'gradient'`, `'pattern'`, `'image'`, `'none'` or `'inherit'` | the kind of the sub-object you set, otherwise `'solid'` | The kind of fill. When `type` and a sub-object disagree, `type` wins. |
| `color` | hex or theme colour | none | The colour of a solid fill. |
| `transparency` | 0 to 100 | `0` | How see-through the fill is. |
| `gradient` | `LinearGradientFillProps` or `RadialGradientFillProps` | none | A gradient. |
| `pattern` | `PatternFillProps` | none | A two-colour pattern. |
| `image` | `ImageFillProps` | none | A picture. |

## Linear gradients

```ts
fill: {
  type: "gradient",
  gradient: {
    kind: "linear",
    angle: 45,
    stops: [
      { position: 0, color: "FFFFFF" },
      { position: 60, color: "accent1", transparency: 20 },
      { position: 100, color: "1F3A5F" },
    ],
  },
}
```

`angle` is the direction the colours run, in degrees clockwise: `0` runs left to right and `90` runs top
to bottom. Any finite angle works, including negative ones and angles past 360.

<svg viewBox="0 0 220 150" width="220" height="150" role="img" aria-label="Gradient angles: 0 runs left to right, 90 top to bottom, 180 right to left, 270 bottom to top">
  <circle cx="110" cy="75" r="42" fill="none" stroke="currentColor" stroke-opacity="0.3" />
  <g stroke="currentColor" stroke-width="2">
    <line x1="110" y1="75" x2="146" y2="75" />
    <line x1="110" y1="75" x2="110" y2="111" />
    <line x1="110" y1="75" x2="74" y2="75" />
    <line x1="110" y1="75" x2="110" y2="39" />
  </g>
  <g fill="currentColor">
    <polygon points="152,75 144,71 144,79" />
    <polygon points="110,117 106,109 114,109" />
    <polygon points="68,75 76,71 76,79" />
    <polygon points="110,33 106,41 114,41" />
    <text x="160" y="79" font-size="12">0</text>
    <text x="102" y="134" font-size="12">90</text>
    <text x="36" y="79" font-size="12">180</text>
    <text x="98" y="24" font-size="12">270</text>
  </g>
</svg>

Each stop has a `position` from 0 to 100 along that direction, a `color`, and an optional
`transparency`. A gradient needs at least two stops; they are placed in position order whatever order
you list them in.

`rotateWithShape` (default `true`) turns the gradient with a rotated shape. `scaled` is PowerPoint's
"scale with shape" setting and is left out unless you set it.

## Radial gradients

```ts
fill: {
  type: "gradient",
  gradient: {
    kind: "radial",
    center: { x: 30, y: 30 },
    stops: [
      { position: 0, color: "FFFFFF" },
      { position: 100, color: "1F3A5F" },
    ],
  },
}
```

The stop at position 0 is the colour at the centre, and later stops spread outwards to the edges.
`center` places that centre as percentages of the filled area, `{ x: 50, y: 50 }` by default.

<svg viewBox="0 0 220 130" width="220" height="130" role="img" aria-label="A radial gradient with its centre at 30 percent across and 30 percent down the filled area">
  <rect x="20" y="15" width="180" height="100" fill="none" stroke="currentColor" />
  <g fill="none" stroke="currentColor" stroke-opacity="0.35">
    <circle cx="74" cy="45" r="22" />
    <circle cx="74" cy="45" r="46" />
    <circle cx="74" cy="45" r="74" />
  </g>
  <circle cx="74" cy="45" r="4" fill="currentColor" />
  <text x="82" y="42" font-size="11" fill="currentColor">center: 30, 30</text>
</svg>

## No fill, inherit, or left out

| You write | On a shape or text box | On a table cell or slide background |
| --- | --- | --- |
| `type: 'none'` | transparent | transparent |
| `type: 'inherit'` | the shape's style or placeholder decides | the table style or master decides |
| no `fill` at all | transparent | the table style or master decides |

Leaving `fill` off a shape makes it transparent, so a shape that should take its colour from its style or
placeholder needs `type: 'inherit'` spelled out.

A chart area or plot area with no fill is transparent too, and there a fill counts only when it has a
colour or a `type`. `{ gradient: { ... } }` on its own leaves the chart area transparent; write
`{ type: 'gradient', gradient: { ... } }`.

## Patterns

```ts
fill: { type: "pattern", pattern: { preset: "diagCross", fgColor: "1F3A5F", bgColor: "FFFFFF" } }
```

`preset` names one of PowerPoint's patterns (see
[`PatternFillProps`](reference/api/index/interfaces/PatternFillProps.md)). `fgColor` defaults to black
and `bgColor` to white.

## Pictures

```ts
fill: { type: "image", image: { path: "texture.png", crop: { l: 10, r: 10 } } }
```

The picture is stretched to the filled area. `crop` trims a percentage from each edge of the source
first; `l` plus `r`, and `t` plus `b`, each stay under 100. `transparency` fades the picture. The
source has to be a raster image: an SVG leaves the area unfilled. To put a picture on a slide as its
own object, clipped to a shape, see [Images in shapes](image-in-shape.md).

## Invalid input

| Condition | Result | Code |
| --- | --- | --- |
| a gradient `kind` other than `linear` or `radial` | throws `UnsupportedFeatureError` | `gradient/type-unsupported` |
| fewer than two stops | throws `InvalidOptionError` | `gradient/too-few-stops` |
| a stop `position` that is not a finite number | throws `InvalidOptionError` | `gradient/stop-position-non-finite` |
| a stop `position` below 0 or above 100 | throws `InvalidOptionError` | `gradient/stop-position-out-of-range` |
| an `angle` that is not a finite number | throws `InvalidOptionError` | `gradient/angle-non-finite` |
| `rotateWithShape` or `scaled` that is not a boolean | throws `InvalidOptionError` | `gradient/rotate-with-shape-not-boolean`, `gradient/scaled-not-boolean` |
| a `center` coordinate below 0 or above 100 | warns and clamps it | `gradient/center-out-of-range` |
| `type: 'pattern'` with no `pattern` | throws `InvalidOptionError` | `pattern-fill/missing-pattern` |
| a picture fill with no `path` or `data` | warns; no fill | `image-fill/missing-source` |
| picture `data` without its base64 header | warns; no fill | `image-fill/missing-base64-header` |
| an SVG picture fill | warns; no fill | `image-fill/svg-unsupported` |
| a picture fill on a chart area or plot area | warns; no fill | `image-fill/unresolved-media` |
| a picture on a `line` | throws `UnsupportedFeatureError` | `line/image-fill-unsupported` |

## Limits

- Radial gradients are circular. PowerPoint's rectangular and shape-following gradients can be read, not authored.
- No gradient tiling or flipping.
- Picture fills are raster only.
- A chart's data point `fill` takes a colour only.

## Reading it back

A deck opened through `pptx-ts/read` reports fills on shapes, tables and table cells:

| Member | Returns |
| --- | --- |
| `resolvedFill` | a solid colour, resolved against the theme, or `null` for any other kind |
| `gradientFill` | the stops (positions from 0 to 1), and `angleDeg` for a linear gradient in the same convention as `angle`, or `path` for a radial one |
| `patternFill` | the preset and both colours |
| `pictureFill` | the embedded picture and how it is stretched or tiled |
| `lineGradient` | a shape's gradient stroke |

`slide.background` is the background a slide shows, including one it takes from its layout or master.
See [PPTX read and round-trip](reference/pptx-read.md).

## See also

- [Tables](tables.md), for table and cell fills in context
- [Images in shapes](image-in-shape.md)
- API reference: [`ShapeFillProps`](reference/api/index/interfaces/ShapeFillProps.md),
  [`LinearGradientFillProps`](reference/api/index/interfaces/LinearGradientFillProps.md),
  [`RadialGradientFillProps`](reference/api/index/interfaces/RadialGradientFillProps.md),
  [`GradientStopProps`](reference/api/index/interfaces/GradientStopProps.md),
  [`PatternFillProps`](reference/api/index/interfaces/PatternFillProps.md),
  [`ImageFillProps`](reference/api/index/interfaces/ImageFillProps.md),
  [`BackgroundProps`](reference/api/index/interfaces/BackgroundProps.md),
  [`ShapeLineProps`](reference/api/index/interfaces/ShapeLineProps.md)
