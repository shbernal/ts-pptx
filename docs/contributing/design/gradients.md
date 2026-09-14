---
doc-schema-version: 1
title: "Fills and gradients: design"
summary: "How fills are serialized and validated: the dispatch every fill site shares, the gradient, pattern and picture builders, where each fill element sits, and the three states a fill can be in."
read_when:
  - Changing how a shape, background, table, chart or line fill is emitted
  - Adding a gradient kind or a fill target
  - Working out why a fill option produces no fill
doc_type: "architecture"
---

# Fills and gradients: design

The user guide is [Fills and gradients](../../fills-and-gradients.md). This page is how the code
behind it is shaped.

## One dispatch

Every fill site goes through `genXmlColorSelection` in `src/gen/drawingml/fill.ts`, which picks the
kind with `resolveFillKind`:

| Input | Kind |
| --- | --- |
| a bare colour string, or nothing | solid |
| an explicit `type` | that type, whatever sub-objects are also set |
| `gradient` | gradient |
| `pattern` | pattern |
| `image` | picture |
| anything else | solid |

A stroke resolves through `resolveLineKind`, which throws `line/image-fill-unsupported` for a picture:
`a:ln` takes `EG_LineFillProperties`, which has no `a:blipFill`.

## Where the fill element goes

| Target | Container | Position and notes |
| --- | --- | --- |
| shape, text box | `p:spPr` | after the geometry, before `a:ln` |
| slide, layout, master background | `p:bg/p:bgPr` | the fill, then `a:effectLst`; an image background (`path`, `data`) wins over a fill, and only the built-in layout gets the white `p:bgRef` default |
| table cell | `a:tcPr` | per cell |
| table | `a:tblPr` | `tableFill`, before the table style id |
| chart area, plot area | `c:spPr` | written only when the fill is stated: a colour string, `color` or `type` (`isStatedFill` in `src/gen/chart/chart-xml.ts`) |
| line | `a:ln` | solid, gradient or pattern |

## Gradients

`genXmlGradientFill` writes `a:gradFill` with `rotWithShape` (`1` unless `rotateWithShape` is `false`),
then `a:gsLst`, then the shade element. It writes no `a:tileRect` and no `flip`.

| Part | Rule |
| --- | --- |
| stops | at least two; each position a finite number from 0 to 100; sorted ascending; `pos` is the position times 1000, rounded |
| stop colour | `createColorElement`, with `transparency` as an alpha modifier |
| linear | `a:lin`; `ang` is the angle brought into 0 to 360 then converted to 60000ths of a degree, so 360 and negative angles stay legal; `scaled` only when given as a boolean |
| radial | `a:path path="circle"` with one `a:fillToRect`; `center` is clamped to 0 to 100 and maps to insets `l = x`, `t = y`, `r = 100 - x`, `b = 100 - y`, each times 1000 |

The stop at position 0 is the start of a linear sweep and the centre of a radial one; the public order
does not change between the two.

| Condition | Result | Code |
| --- | --- | --- |
| `kind` other than `linear` or `radial` | throws `UnsupportedFeatureError` | `gradient/type-unsupported` |
| fewer than two stops | throws `InvalidOptionError` | `gradient/too-few-stops` |
| a stop position that is not a finite number | throws `InvalidOptionError` | `gradient/stop-position-non-finite` |
| a stop position outside 0 to 100 | throws `InvalidOptionError` | `gradient/stop-position-out-of-range` |
| an angle that is not a finite number | throws `InvalidOptionError` | `gradient/angle-non-finite` |
| `rotateWithShape` or `scaled` that is not a boolean | throws `InvalidOptionError` | `gradient/rotate-with-shape-not-boolean`, `gradient/scaled-not-boolean` |
| a `center` coordinate outside 0 to 100 | clamps and warns | `gradient/center-out-of-range` |

## Patterns and pictures

`genXmlPatternFill` writes `a:pattFill` with the preset, `a:fgClr` (default `000000`) and `a:bgClr`
(default `FFFFFF`); a pattern fill with no `pattern` throws `pattern-fill/missing-pattern`.

A picture fill registers its media when the object is added (`registerImageFillMedia` in
`src/gen/define/image.ts`). A missing source, `data` without its base64 header, or an SVG source warns
(`image-fill/missing-source`, `image-fill/missing-base64-header`, `image-fill/svg-unsupported`) and
turns the fill into none. At write time `genXmlImageFill` needs that relationship; without it, which
is always the case for a chart element since nothing registers chart media, it warns
`image-fill/unresolved-media` and writes `a:noFill`. `transparency` becomes `a:alphaModFix`, and `crop`
becomes `a:srcRect`.

## Three states

| Spelling | Shape or text box | Table cell, slide background |
| --- | --- | --- |
| `type: 'none'` | `a:noFill` | `a:noFill` |
| `type: 'inherit'` | no fill element, so `p:style` or the placeholder decides | no fill element |
| fill omitted | `a:noFill` | no fill element, so the table style or master decides |

Omitting `fill` on a shape means no fill, not inherit, because the shape emitter defaults a missing
fill to `a:noFill`. That is why `inherit` needs its own spelling there.
