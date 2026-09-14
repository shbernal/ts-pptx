---
doc-schema-version: 1
title: "Layout units"
summary: "Public slide-layout constants and unit helpers for PowerPoint geometry."
read_when:
  - Defining custom presentation layouts
  - Converting inches, points, pixels, or EMUs
  - Avoiding PowerPoint widescreen size drift
doc_type: "reference"
---

# Layout units

ts-pptx exposes PowerPoint geometry constants and small conversion helpers
from the root package (`pptx-ts`).

## Standard layouts

`STANDARD_LAYOUTS` contains the built-in presentation sizes used by
`pptx.layout`:

- `LAYOUT_4x3`: `10 x 7.5 in`
- `LAYOUT_16x9`: `10 x 5.625 in`
- `LAYOUT_16x10`: `10 x 6.25 in`
- `LAYOUT_WIDE`: PowerPoint widescreen, `13.333 x 7.5 in`

PowerPoint widescreen is stored as `40 / 3 x 7.5 in`, which converts exactly to
`12192000 x 6858000` EMUs. Prefer the constant over writing rounded decimal
widths by hand.

```ts
import TsPptx, { STANDARD_LAYOUTS } from "pptx-ts"

const pptx = new TsPptx()
const wide = STANDARD_LAYOUTS.LAYOUT_WIDE

pptx.defineLayout({ name: "POWERPOINT_WIDESCREEN", width: wide.widthIn, height: wide.heightIn })
pptx.layout = "POWERPOINT_WIDESCREEN"
```

The built-in `pptx.layout = "LAYOUT_WIDE"` also uses the same exact EMU
dimensions.

## Unit helpers

The public helpers are:

- `inchesToEmu(inches)`
- `pointsToEmu(points)`
- `pixelsToEmu(pixels, dpi)`
- `emuToInches(emu)`
- `emuToPoints(emu)`
- `emuToPixels(emu, dpi)`

The public unit constants are:

- `EMU_PER_INCH`: `914400`
- `EMU_PER_POINT`: `12700`
- `POINTS_PER_INCH`: `72`

## Positions and sizes

Every object is placed by `x`, `y`, `w` and `h`, in any of the units listed in
[Core concepts](../getting-started/concepts.md#positions-and-sizes). A negative
`w` or `h` measures the box the other way, left of `x` or above `y`. ts-pptx
writes that box with a positive size and a flip, because PowerPoint refuses to
open a file that stores a negative size.

```ts
// Runs upward, from (1, 3) to (2.5, 1)
slide.addShape("line", { x: 1, y: 3, w: 1.5, h: -2 })
```

| You pass | Box written | Flip written |
| --- | --- | --- |
| `x: 1, y: 3, w: 1.5, h: -2` | `x: 1, y: 1, w: 1.5, h: 2` | vertical |
| `x: 4, y: 1, w: -3, h: 1` | `x: 1, y: 1, w: 3, h: 1` | horizontal |
| `x: 4, y: 1, w: -3, h: 1, flipH: true` | `x: 1, y: 1, w: 3, h: 1` | none |

- A flip that comes from a negative size combines with the `flipH` or `flipV`
  you set. Either one alone mirrors the object, and the two together cancel.
- Unit strings behave the same way: `w: "-2in"` and `w: "-25%"` are handled
  like a negative number.
- Shapes, text boxes, images, media and group frames get the flip. Charts,
  tables, OLE objects, 3D models and zooms get the corrected box with no flip.
- `addConnector()` takes two endpoints instead of a size, and derives its box
  and flips from them the same way. See
  [Place the endpoints](../connectors.md#place-the-endpoints).
