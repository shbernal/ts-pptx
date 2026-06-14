---
doc-schema-version: 1
title: "Image Embedded In A Shape"
summary: "Clip a picture to a preset or freeform shape and crop it to fill the box (blipFill + custGeom)."
read_when:
  - Clipping an image to a circle, rounded rectangle, hexagon, or freeform path
  - Filling a clip shape with a center-cropped photo (cover/contain)
  - Reproducing a picture-placeholder look (e.g. a half-disc "D" cover image)
doc_type: "guide"
---

# Image Embedded In A Shape

`slide.addImage()` can clip a picture to a shape and, independently, crop the
source bitmap so it fills that shape at the right aspect ratio. This is the
idiomatic OOXML form — a `<p:pic>` whose `<p:spPr>` carries the clip geometry and
whose `<p:blipFill>` carries the source crop — exactly what a PowerPoint *picture
placeholder* produces.

## Choosing the clip geometry

Three mutually exclusive ways to set the clip, in precedence order:

| Option | Emits | Use when |
|--------|-------|----------|
| `points` | `<a:custGeom>` (freeform path) | Arbitrary outline — a half-disc, a speech bubble, any custom silhouette. **Wins over `shape`/`rounding`.** |
| `shape` | `<a:prstGeom prst="…">` | A named PowerPoint preset — `'roundRect'`, `'hexagon'`, `'ellipse'`, etc. |
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
the box. By default the image is stretched to the box extent, which distorts a
photo whose aspect ratio differs from the clip box. Pair the clip with
`sizing: { type: 'cover' }` to center-crop the source so it fills the box at its
natural aspect ratio:

```js
slide.addImage({
  path: 'photo.png', x: 1, y: 1, w: 2, h: 3,
  points: [/* clip path */],
  sizing: { type: 'cover', w: 2, h: 3 },
})
```

- `cover` — scales the source to **cover** the box, cropping the overflow (no
  distortion, no gaps). This is what you want for a photo behind a clip shape.
- `contain` — scales the source to **fit** inside the box (letterbox; negative
  `srcRect` inset).
- `crop` — cuts an explicit window using `x`/`y`/`w`/`h` offsets.

`cover`/`contain` read the image's natural pixel dimensions from the embedded
bytes (PNG/JPEG/GIF/BMP/WebP header), so the crop is aspect-correct. For SVG or an
unrecognized format the displayed `w`/`h` ratio is used as a fallback and a
warning is logged.

`points` (clip) lives in `<p:spPr>` and `sizing` (crop) lives in `<p:blipFill>`,
so the two compose freely. The emitted blip fill uses the canonical
`<a:srcRect/><a:stretch><a:fillRect/></a:stretch>` form (ECMA-376 §L.4.8.4.3),
which PowerPoint and LibreOffice both render with a clean clip edge.

## Worked example — the half-disc ("D") cover

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
    { x: 0, y: h / 2, curve: { type: 'arc', hR: h / 2, wR: fx, stAng: 90, swAng: 180 } },
    { close: true },
  ],
  sizing: { type: 'cover', w, h }, // center-crop the (wide) photo into the tall "D"
})
```

The wide source photo is cropped to the box aspect — not squashed — and the curved
edge is a smooth ellipse arc. See `demos/common/image-in-shape.js` for a runnable
version, and `test/regression/image-shape.test.js` for the composition tests.

## Borders, shadows, recolor

A clipped picture still supports `line` (a `<a:ln>` outline that follows the clip
geometry), `shadow`, `transparency`, and `duotone` — the same picture-formatting
vocabulary as an unclipped image.
