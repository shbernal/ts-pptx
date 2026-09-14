---
doc-schema-version: 1
title: "Layout units"
summary: "The standard slide sizes in inches and EMU, the unit conversion helpers and constants exported from pptx-ts, and how a negative w or h becomes a positive size plus a flip."
read_when:
  - Setting a standard slide size, or checking its exact EMU dimensions
  - Converting between inches, points, pixels, percentages and EMU
  - Working out which constant a conversion uses, or why a helper threw coord/out-of-range
  - Placing an object with a negative w or h
doc_type: "reference"
---

# Layout units

PowerPoint stores geometry in EMU (English Metric Units), 914400 to the inch. The slide sizes, helpers and constants on this page are exported from `pptx-ts`.

## Standard layouts

`STANDARD_LAYOUTS` holds the four sizes `pptx.layout` accepts without `defineLayout()`. A new presentation uses `LAYOUT_16x9`.

| Name | Inches | EMU |
| --- | --- | --- |
| `LAYOUT_16x9` | 10 × 5.625 | 9144000 × 5143500 |
| `LAYOUT_16x10` | 10 × 6.25 | 9144000 × 5715000 |
| `LAYOUT_4x3` | 10 × 7.5 | 9144000 × 6858000 |
| `LAYOUT_WIDE` | 13.333 × 7.5 | 12192000 × 6858000 |

Each entry carries `widthIn`, `heightIn`, `widthEmu` and `heightEmu`. `LAYOUT_WIDE` is 40/3 inches wide, which converts to exactly 12192000 EMU. A width typed as `13.333` converts to 12191695 EMU, so use the preset.

```ts
import TsPptx, { STANDARD_LAYOUTS, inchesToEmu } from "pptx-ts"

const pptx = new TsPptx()
pptx.layout = STANDARD_LAYOUTS.LAYOUT_WIDE // the same as pptx.layout = "LAYOUT_WIDE"

const wide = STANDARD_LAYOUTS.LAYOUT_WIDE
inchesToEmu(wide.widthIn) // 12192000, the same as wide.widthEmu
```

## Unit helpers

| Helper | From | To | Constant |
| --- | --- | --- | --- |
| `inchesToEmu(inches)` | inches | EMU | `EMU_PER_INCH` |
| `pointsToEmu(points)` | points | EMU | `EMU_PER_POINT` |
| `pixelsToEmu(pixels, dpi)` | pixels at `dpi` | EMU | `EMU_PER_INCH` |
| `percentToEmu(percent, axisEmu)` | a percentage of `axisEmu`, `50` for half | EMU | none |
| `coordToEmu(value, axisEmu)` | a coordinate in any form `x`, `y`, `w` and `h` accept | EMU | the ones above, and `DEFAULT_PX_PER_INCH` for `"px"` |
| `emuToInches(emu)` | EMU | inches | `EMU_PER_INCH` |
| `emuToPoints(emu)` | EMU | points | `EMU_PER_POINT` |
| `emuToPixels(emu, dpi)` | EMU | whole pixels at `dpi` | `EMU_PER_INCH` |
| `ptToHundredths(pt)` | points | hundredths of a point, the unit of a font size | `HUNDREDTHS_PER_POINT` |

- The helpers that return EMU, pixels or hundredths round to a whole number. `emuToInches` and `emuToPoints` do not round.
- `NaN` or an infinite input throws `InvalidOptionError` with the code `coord/non-finite`.
- An EMU result outside `MIN_COORDINATE_EMU` to `MAX_COORDINATE_EMU` throws `coord/out-of-range`.
- A `dpi` of 0 or less throws `coord/not-positive`.
- `coordToEmu` throws `coord/invalid-format` for a string it cannot parse, such as `"2cm"`. It reads a bare number above 1000 as inches and warns `coord/bare-number-is-inches`.

## Constants

| Constant | Value | Meaning |
| --- | --- | --- |
| `EMU_PER_INCH` | 914400 | EMU in an inch |
| `EMU_PER_POINT` | 12700 | EMU in a point |
| `POINTS_PER_INCH` | 72 | points in an inch |
| `DEFAULT_PX_PER_INCH` | 96 | pixels in an inch for the `"px"` coordinate unit |
| `HUNDREDTHS_PER_POINT` | 100 | stored units in a point of font size, character spacing or line spacing |
| `ANGLE_UNITS_PER_DEGREE` | 60000 | stored angle units in a degree |
| `PERCENT_SCALE` | 100000 | the stored value of 100% |
| `FIXED_PCT_PER_PERCENT` | 1000 | the stored value of 1% |
| `MIN_COORDINATE_EMU` | -27273042329600 | the smallest coordinate DrawingML allows |
| `MAX_COORDINATE_EMU` | 27273042316900 | the largest coordinate DrawingML allows |

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
