---
doc-schema-version: 1
title: "Image Embedded In A Shape"
summary: "Clip a picture to a preset or freeform shape and crop it to fill the box (blipFill + custGeom)."
read_when:
  - Clipping an image to a circle, rounded rectangle, hexagon, or freeform path
  - Filling a clip shape with a center-cropped photo (cover/contain)
  - Reproducing a picture-placeholder look (e.g. a half-disc "D" cover image)
  - Reaching for a named clip silhouette instead of hand-authoring a points path
  - Working out why an SVG icon is letterboxed, stretched, or sized at 1 inch
doc_type: "guide"
---

# Image Embedded In A Shape

`slide.addImage()` can clip a picture to a shape and, independently, crop the
source bitmap so it fills that shape at the right aspect ratio. This is the
idiomatic OOXML form (a `<p:pic>` whose `<p:spPr>` carries the clip geometry and
whose `<p:blipFill>` carries the source crop) exactly what a PowerPoint *picture
placeholder* produces.

## Choosing the clip geometry

Three mutually exclusive ways to set the clip, in precedence order:

| Option | Emits | Use when |
|--------|-------|----------|
| `points` | `<a:custGeom>` (freeform path) | Arbitrary outline: a half-disc, a speech bubble, any custom silhouette. **Wins over `shape`/`rounding`.** |
| `shape` | `<a:prstGeom prst="…">` | A named PowerPoint preset: `'roundRect'`, `'hexagon'`, `'ellipse'`, etc. |
| `rounding: true` | `<a:prstGeom prst="ellipse">` | Shorthand for a circular/elliptical crop. Lowest precedence. |

With none of these the picture stays a plain rectangle (`prst="rect"`).

```js
// Preset clip
slide.addImage({ path: 'avatar.png', x: 1, y: 1, w: 2, h: 2, shape: 'roundRect', rectRadius: 0.25 })

// Circle (shorthand)
slide.addImage({ path: 'avatar.png', x: 1, y: 1, w: 2, h: 2, rounding: true })

// Freeform clip (triangle)
slide.addImage({
  path: 'photo.png', x: 1, y: 1, w: 2, h: 2,
  points: [{ x: 1, y: 0 }, { x: 2, y: 2 }, { x: 0, y: 2 }, { close: true }],
})
```

`points` are authored in the image's **own** inch/EMU space (`0..w`, `0..h`), not
slide-relative and not normalized. The DSL matches freeform shapes:
`moveTo` / `lnTo` / `cubicBezTo` / `quadBezTo` / `arcTo` / `close`.

## Filling the shape: pair with `sizing`

A clip changes the *outline*; it does not change how the source pixels map into
the box. A **raster** is stretched to the box extent by default, which distorts a
photo whose aspect ratio differs from the clip box. Pair the clip with
`sizing: { type: 'cover' }` to center-crop the source so it fills the box at its
natural aspect ratio:

```js
slide.addImage({
  path: 'photo.png', x: 1, y: 1, w: 2, h: 3,
  points: [/* clip path */],
  sizing: { type: 'cover' },   // w/h default to the picture's own box
})
```

- `cover`: scales the source to **cover** the box, cropping the overflow (no
  distortion, no gaps). This is what you want for a photo behind a clip shape.
- `contain`: scales the source to **fit** inside the box (letterbox; negative
  `srcRect` inset).
- `crop`: cuts an explicit window using `x`/`y`/`w`/`h` offsets.
- `stretch`: fills the box regardless of aspect. A raster's default; name it to
  opt a vector out of the one below.

`sizing.w` / `sizing.h` default to the picture's own `w` / `h`, so supply them only
when the fit box is genuinely something other than the picture.

`cover`/`contain` read the image's natural dimensions from the embedded bytes (a
PNG/JPEG/GIF/BMP/WebP header, or an SVG's `width`/`height` or `viewBox`), so the
crop is aspect-correct. For an unrecognized format, or an SVG carrying no
intrinsic size at all, the displayed `w`/`h` ratio is used as a fallback and a
warning is logged.

### Vectors place aspect-correct without being asked

An SVG states its own aspect ratio, and a glyph squashed into a box that disagrees
with it is a defect rather than a layout choice. So a **vector** source with no
`sizing` is letterboxed to its intrinsic ratio inside its box (exactly as
`sizing: { type: 'contain' }` would) instead of being stretched:

```js
// A square icon in a non-square box: centered at its true aspect, not squashed.
slide.addImage({ svg: iconMarkup, x: 1, y: 1, w: 3, h: 1 })

// Opt out where the distortion is the point (a stretched vector backdrop, say).
slide.addImage({ svg: bandMarkup, x: 0, y: 0, w: 13.33, h: 0.4, sizing: { type: 'stretch' } })
```

Nothing is emitted when the ratios already agree: a square glyph in a square box
produces the same plain `<a:stretch>` it always did. An SVG with neither a
`viewBox` nor `width`/`height` cannot be measured, so it stretches, silently: no
sizing was requested, so there is nothing to warn about.

The same intrinsic ratio fills in a missing dimension: `{ svg, w: 4 }` on a 2:1
`viewBox` is 4in × 2in. What it will *not* do is treat user units as pixels: an
SVG given neither `w` nor `h` falls back to 1 inch rather than becoming a
quarter-inch object because its icon set was authored on a 24-unit grid.

`points` (clip) lives in `<p:spPr>` and `sizing` (crop) lives in `<p:blipFill>`,
so the two compose freely. The emitted blip fill uses the canonical
`<a:srcRect/><a:stretch><a:fillRect/></a:stretch>` form (ECMA-376 §L.4.8.4.3),
which PowerPoint and LibreOffice both render with a clean clip edge.

## Worked example: the half-disc ("D") cover

A right-flush half-disc photo, the curved edge expressed with a single `arcTo`,
center-cropped to fill the portrait box. The flat side sits at `0.3179·w` from the
left (the placeholder geometry from the source deck).

```js
const w = 5.22, h = 7.5            // box size (inches)
const fx = 0.3179 * w              // x of the flat edge

slide.addImage({
  path: 'cover-photo.jpg',
  x: 0, y: 0, w, h,
  points: [
    { x: fx, y: 0 },               // top, at the flat edge
    { x: w,  y: 0 },               // top-right
    { x: w,  y: h },               // bottom-right
    { x: fx, y: h },               // bottom, back to the flat edge
    // curved left edge: an ellipse arc bulging left, from the bottom sweeping 180°
    { curve: { type: 'arc', hR: h / 2, wR: fx, stAng: 90, swAng: 180 } },
    { close: true },
  ],
  sizing: { type: 'cover', w, h }, // center-crop the (wide) photo into the tall "D"
})
```

An `arc` node takes no `x`/`y`. An `<a:arcTo>` carries no explicit end point
(PowerPoint derives it from the current pen position, the radii and the swept
angle), so the arc above ends where the 180° sweep lands, back at the flat edge.
Supplying an end point emits a warning and is otherwise ignored. Unlike a shape
rotation, arc angles are not wrapped into `0..360`: `swAng: 400` draws a 400°
sweep, not a 40° one.

The wide source photo is cropped to the box aspect (not squashed) and the curved
edge is a smooth ellipse arc. See `test/regression/image/image-shape.test.js` for the
composition tests.

## Named silhouettes: `clipPath()`

The half-disc above recurs often enough (it is what a cover-slide picture
placeholder cuts) that it is available as data rather than as arithmetic.
`clipPath(shape, w, h)` resolves a named `ClipShape` to the same `points` array:

```js
import { clipPath } from '@shbernal/ts-pptx'

const w = 5.22, h = 7.5

slide.addImage({
  path: 'cover-photo.jpg', x: 0, y: 0, w, h,
  points: clipPath({ kind: 'half-disc', flat: 'right' }, w, h),
  sizing: { type: 'cover' },
})
```

`flat` names the edge the straight side sits on (`'right'` = flat right edge, arc
bulging left); `preset` picks the proportion: `'deep'` (the default; the arc spans
about 32% of the box width, symmetric about mid-height) or `'shallow'` (about 13%,
with its apex just below mid-height). Both are traced as two cubic Béziers rather
than an `arcTo`, so unlike the hand-authored example above they are not perfect
half-ellipses: they are the placeholder proportions PowerPoint decks actually use.

**`w` and `h` must be the size the picture is drawn at.** The returned path is in
the image box's *own* inch space (`0..w`, `0..h`), because a `custGeom` point given
as `%` resolves against the **slide**, not the box. That is why the box size is an
argument at all: the silhouette's fractions are multiplied out at build time, so
one shape scales to any region. Hand a `clipPath` result to a picture of a
different size and the clip lands in the wrong place.

## Borders, shadows, recolor

A clipped picture still supports `line` (a `<a:ln>` outline that follows the clip
geometry), `shadow`, `transparency`, and the recolor modes: the same
picture-formatting vocabulary as an unclipped image. The recolor options are
`duotone` (`{ shadow, highlight }`), `grayscale` (`true`), `biLevel`
(`{ threshold }`, a `0.0–1.0` luminance split into black/white), and `clrChange`
(`{ from, to }`, repaint one source color as another); they mirror the five
effects the read model's `Picture.recolor` decodes.
