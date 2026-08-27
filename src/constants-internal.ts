/**
 * Generator-internal constants — the non-published half of the enums module.
 *
 * `enums.ts` holds the genuinely public surface (the `ChartType`/`ShapeType`/
 * `SchemeColor`/`AlignH`… enums, the shape-preset sets, and the `*_NAME` unions) and is
 * `export *`-ed by every entrypoint. This module holds the constants only the generators
 * need — default styling, fixed axis/field ids, colour palettes, layout bases, and the
 * base64 placeholder images — none of which are part of the public API. Splitting them out
 * keeps `export *` from leaking them onto the package surface, mirroring the
 * `units-internal.ts` / `types/internal.ts` split. Nothing here is exported from an entrypoint.
 */

import type { BorderProps, OptsChartGridLine } from './types/index.js'
import type { ShadowPropsInternal } from './types/internal.js'

// CONST
export const CRLF = '\r\n' // AKA: Chr(13) & Chr(10)
// Base for `<p:sldLayoutId id>` (layout N gets BASE + N). 2147483649 = 0x80000001:
// ECMA-376 ST_SlideLayoutId requires id >= 0x80000000, and PowerPoint conventionally
// starts layout IDs one past that. See `makeXmlMaster` in `gen/slide/master.ts`.
export const LAYOUT_IDX_SERIES_BASE = 2147483649
export const REGEX_HEX_COLOR = /^[0-9a-fA-F]{6}$/
export const LINEH_MODIFIER = 1.67 // AKA: Golden Ratio Typography

export const DEF_BULLET_MARGIN = 27
export const DEF_CELL_BORDER: BorderProps = { type: 'solid', color: '666666', width: 1 }
export const DEF_CELL_MARGIN_IN: [number, number, number, number] = [0.05, 0.1, 0.05, 0.1] // "Normal" margins in PPT-2021 ("Narrow" is `0.05` for all 4)
export const DEF_CHART_BORDER: BorderProps = { type: 'solid', color: '363636', width: 1 }
export const DEF_CHART_GRIDLINE: OptsChartGridLine = { color: '888888', style: 'solid', size: 1, cap: 'flat' }
export const DEF_FONT_COLOR = '000000'
export const DEF_FONT_SIZE = 12
export const DEF_FONT_TITLE_SIZE = 18
export const DEF_PRES_LAYOUT = 'LAYOUT_16x9'
export const DEF_PRES_LAYOUT_NAME = 'DEFAULT'
export const DEF_SHAPE_LINE_COLOR = '333333'
export const DEF_SHAPE_SHADOW: ShadowPropsInternal = {
	type: 'outer',
	blur: 3,
	offset: 23000 / 12700,
	angle: 90,
	color: '000000',
	_alpha: 0.35,
	rotateWithShape: true,
}
export const DEF_SLIDE_MARGIN_IN: [number, number, number, number] = [0.5, 0.5, 0.5, 0.5] // TRBL-style
export const DEF_TEXT_SHADOW: ShadowPropsInternal = {
	type: 'outer',
	blur: 8,
	offset: 4,
	angle: 270,
	color: '000000',
	_alpha: 0.75,
}
export const DEF_TEXT_GLOW = { size: 8, color: 'FFFFFF', opacity: 0.75 }

// Fixed `<c:axId val>` identifiers that wire each chart's axes together (e.g. a value
// axis's `<c:crossAx>` points at its category axis by id). ECMA-376 only requires the id
// be a unique unsigned int within the chart; these stable constants keep the emitted
// cross-references deterministic across decks. Consumed throughout `gen/chart/`.
export const AXIS_ID_VALUE_PRIMARY = '2094734552'
export const AXIS_ID_VALUE_SECONDARY = '2094734553'
export const AXIS_ID_CATEGORY_PRIMARY = '2094734554'
export const AXIS_ID_CATEGORY_SECONDARY = '2094734555'
export const AXIS_ID_SERIES_PRIMARY = '2094734556'

export const LETTERS: string[] = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('')
// The default series palettes hold their distinct entries once. Every consumer reaches them
// through `paletteColor()` (`gen/chart/chart-parts.ts`), which wraps back to the start, so a
// chart with more series or points than the palette has entries repeats it rather than
// running off the end.
export const BARCHART_COLORS: string[] = [
	'C0504D',
	'4F81BD',
	'9BBB59',
	'8064A2',
	'4BACC6',
	'F79646',
	'628FC6',
	'C86360',
]
export const PIECHART_COLORS: string[] = [
	'5DA5DA',
	'FAA43A',
	'60BD68',
	'F17CB0',
	'B2912F',
	'B276B2',
	'DECF3F',
	'F15854',
	'A7A7A7',
]

// Fixed GUID PowerPoint stamps on a slide-number field: `<a:fld id="{...}" type="slidenum">`.
// Microsoft-specific (not defined by ECMA-376); reused verbatim so PowerPoint recognizes the
// field and auto-updates the page number. Consumed in `gen/slide/object.ts` (`slideObjectToXml`).
export const SLDNUMFLDID = '{F7021451-1387-4CA6-816F-3879F97B5CBC}'

// Cached text of a slide-number field on a part that has no slide number of its own — a master
// or a layout. Guillemets around a hash, which is verbatim what PowerPoint caches there: every
// en-US master and layout in `test/read/fixtures/` writes `<a:t>‹#›</a:t>`. (A German-locale
// deck writes `‹Nr.›`; the cache is recomputed on open, so the en-US spelling is the one to
// emit rather than a per-locale table.) Consumed in `gen/slide/object.ts`.
export const SLDNUM_PLACEHOLDER_TEXT = '‹#›'

// The XML prolog every emitted OOXML part begins with. Kept as one constant so a
// stray edit can't desync one part's declaration from the rest.
export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
