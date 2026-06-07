---
doc-schema-version: 1
title: "Layout Units"
summary: "Public slide-layout constants and unit helpers for PowerPoint geometry."
read_when:
  - Defining custom presentation layouts
  - Converting inches, points, pixels, or EMUs
  - Avoiding PowerPoint widescreen size drift
doc_type: "reference"
---

# Layout Units

PptxGenJS exposes PowerPoint geometry constants and small conversion helpers
from the root package and from `@shbernal/pptxgenjs/core`.

## Standard Layouts

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
import pptxgen, { STANDARD_LAYOUTS } from "@shbernal/pptxgenjs"

const pptx = new pptxgen()
const wide = STANDARD_LAYOUTS.LAYOUT_WIDE

pptx.defineLayout({ name: "POWERPOINT_WIDESCREEN", width: wide.widthIn, height: wide.heightIn })
pptx.layout = "POWERPOINT_WIDESCREEN"
```

The built-in `pptx.layout = "LAYOUT_WIDE"` also uses the same exact EMU
dimensions.

## Unit Helpers

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
