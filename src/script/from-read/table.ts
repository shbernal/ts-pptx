/**
 * `Table` → `addTable(TableRow[], TableProps)`.
 *
 * Tables are the one place the exact-EMU rule stops. `colW` and `rowH` are number-typed
 * (inches) rather than `Coord`-typed, so they cannot take a raw-EMU string the way every
 * position and size can. They print at six decimal places, which is the measured minimum
 * at which `Math.round(inches × 914400)` returns the original EMU — the drift is bounded
 * at 0.4572 EMU, roughly half a millionth of an inch, and rounding shorter to make the
 * numbers look tidy would turn a cosmetic problem into a real geometry loss.
 *
 * Table *style* is the opposite story, and better than it looks. A style is a GUID into
 * the deck's `ppt/tableStyles.xml`, and a template-anchored output keeps that part intact,
 * so passing the source GUID through reproduces the banding, header shading, and borders
 * exactly rather than approximately. That is why per-cell fills read off the style graph
 * are deliberately *not* baked into each cell here: doing so would freeze the appearance
 * and stop the table responding to its own style.
 */
import type { CellBorder, Table, TableCell } from '../../read/api/table.js'
import type { GraphicFrame } from '../../read/api/shapes.js'
import type { NoteScope } from '../fidelity.js'
import type { CallIr, IrValue } from '../ir.js'
import type { AssetResolver } from './shape.js'
import { compact, emu, inches, literalColor, orUndefined } from './values.js'
import { pictureFillOption, type PictureFillSubject } from './picture-fill.js'
import { runOptions, textRuns } from './text.js'

/** How {@link pictureFillOption}'s notes name a cell. */
const CELL_PICTURE_FILL: PictureFillSubject = {
	construct: 'table.cell.fill.picture',
	subject: 'this table cell',
	element: 'a:tcPr/a:blipFill',
}

/** `a:tcPr/@anchor` → the write API's `valign`. */
const ANCHOR: Record<string, string> = { t: 'top', ctr: 'middle', b: 'bottom' }

export function tableCall(frame: GraphicFrame, table: Table, notes: NoteScope, assets: AssetResolver): CallIr {
	const styleId = table.styleId
	const hasStyle = table.resolvedStyle !== null
	const rows: IrValue[] = []
	const rowHeights: number[] = []

	for (const row of table.rows) {
		const cells: IrValue[] = []
		for (const cell of row.cells) {
			// A merge continuation is the covered half of a span: it carries no content and
			// exists only so the grid stays rectangular. The write path derives those from
			// colspan/rowspan, so emitting them would double-count the span.
			if (cell.isMergeContinuation) continue
			cells.push(cellIr(cell, hasStyle, notes, assets))
		}
		rows.push(cells)
		rowHeights.push(row.heightEmu ?? 0)
	}

	// A row of height 0 is PowerPoint's "auto — as tall as its content needs", not a
	// zero-height row. The write path takes it literally and then lets the row grow anyway,
	// so the output reports whatever height the content forced. Nothing is visibly wrong;
	// the row simply stops being auto and is pinned to that measurement.
	if (rowHeights.some((height) => height === 0) && !rowHeights.every((height) => height === 0)) {
		notes.note(
			'table.rowAuto',
			'approximated',
			'unsupported',
			'at least one row is auto-height (a:tr/@h of 0) while others are not; rowH must be given for every row or none, so the auto rows are emitted as 0 and come back pinned to the height their content produced'
		)
	}

	if (styleId === null) {
		notes.note(
			'table.style',
			'dropped',
			'unsupported',
			"this table names no a:tableStyleId, so it takes the generated deck's default table style"
		)
	}

	const columnWidths = table.columnWidths
	const options = compact({
		...positionOfFrame(frame),
		objectName: frame.name || undefined,
		// The source GUID resolves against the destination's own tableStyles.xml, which a
		// template-anchored output carries over intact.
		tableStyle: orUndefined(styleId),
		hasHeader: table.firstRowHeader ? true : undefined,
		hasBandedRows: table.bandedRows ? true : undefined,
		colW: columnWidths.every((w) => w === null) ? undefined : columnWidths.map((w) => inches(w ?? 0)),
		rowH: rowHeights.every((h) => h === 0) ? undefined : rowHeights.map(inches),
	})

	return { method: 'addTable', args: [rows, options ?? {}], ...(frame.name ? { sourceName: frame.name } : {}) }
}

/** A graphic frame's position; identical to a shape's, but frames never rotate or flip. */
function positionOfFrame(frame: GraphicFrame): Record<string, IrValue> {
	const box = frame.absoluteFrame
	if (!box) return {}
	return { x: emu(box.left), y: emu(box.top), w: emu(box.width), h: emu(box.height) }
}

/**
 * One cell as a `TableCell`.
 *
 * Cell text reuses the shape text mapper, so a per-run format inside a cell behaves
 * identically to one on a shape — the read model already shares `TextFrame` between the
 * two, and diverging here would be a difference with no cause.
 */
function cellIr(cell: TableCell, hasStyle: boolean, notes: NoteScope, assets: AssetResolver): IrValue {
	const frame = cell.textFrame
	const anchor = cell.anchor
	const margins = cell.marginsEmu

	const options = compact({
		fill: cellFill(cell, hasStyle, notes, assets),
		border: cellBorders(cell, notes),
		valign: anchor === null ? undefined : ANCHOR[anchor],
		// `textDirection` reaches `a:tcPr/@vert` through the same table-cell inheritance list every
		// other cell option uses, so a vertical label survives the round trip. Anything outside the
		// option's own union would be written back out verbatim as an invalid attribute.
		textDirection: cellTextDirection(cell, notes),
		// Only the two values the write option accepts survive; anything else in the source
		// deck would be written straight back out as an invalid attribute.
		horzOverflow: cell.horzOverflow === 'clip' || cell.horzOverflow === 'overflow' ? cell.horzOverflow : undefined,
		colspan: cell.gridSpan > 1 ? cell.gridSpan : undefined,
		rowspan: cell.rowSpan > 1 ? cell.rowSpan : undefined,
		// `margin` takes inches, and the read model reports the insets in EMU.
		margin: margins
			? [inches(margins.top ?? 0), inches(margins.right ?? 0), inches(margins.bottom ?? 0), inches(margins.left ?? 0)]
			: undefined,
		// A cell's own runs may each carry formatting; the first run's options double as the
		// cell default, which is how the write path applies cell-level character formatting.
		...(frame?.paragraphs[0]?.runs[0] ? (runOptions(frame.paragraphs[0].runs[0], notes) ?? {}) : {}),
	})

	return compact({ text: frame ? textRuns(frame, notes) : cell.text, options }) ?? { text: '' }
}

/** The `ST_TextVerticalType` values `TextBaseProps.textDirection` accepts. */
const WRITABLE_TEXT_DIRECTIONS = new Set(['horz', 'vert', 'vert270', 'wordArtVert'])

/**
 * A cell's `a:tcPr/@vert` as the write API's `textDirection`.
 *
 * `ST_TextVerticalType` has nine values and the write option covers four. The rest
 * (`eaVert`, `mongolianVert`, the WordArt right-to-left variants) are East-Asian layout
 * modes with no `TextBaseProps` spelling, so they are noted rather than written — passing
 * one straight through would put a value the option does not admit into the attribute.
 */
function cellTextDirection(cell: TableCell, notes: NoteScope): IrValue | undefined {
	const vert = cell.verticalText
	if (vert === null || vert === 'horz') return undefined
	if (WRITABLE_TEXT_DIRECTIONS.has(vert)) return vert
	notes.note(
		'table.cell.vert',
		'dropped',
		'unwritable',
		`this cell's a:tcPr/@vert is \`${vert}\`, which textDirection does not spell (it covers horz/vert/vert270/wordArtVert), so its text lands horizontal`
	)
	return undefined
}

/**
 * A cell's fill, and the one place this mapper deliberately emits *less* than it can read.
 *
 * `resolvedFill` folds two different things together: the cell's own `a:solidFill`, and the
 * colour it merely inherits from the table style's header/banding rules. There is no
 * accessor that returns the first without the second. So when the table has a resolvable
 * style, writing `resolvedFill` back would turn every banded cell into an explicitly-filled
 * one — the table would look right until someone changed its style, and then nothing would
 * move. The style GUID reproduces those colours anyway, so only a scheme token (which can
 * only come from the cell itself) is emitted, and the ambiguity is noted.
 *
 * With no style in play, `resolvedFill` can only be the cell's own, so it is safe to use.
 */
function cellFill(cell: TableCell, hasStyle: boolean, notes: NoteScope, assets: AssetResolver): IrValue | undefined {
	// A cell's fill is `ShapeFillProps`, the same type a shape's is, so an image-filled cell
	// re-embeds its bytes exactly as an image-filled shape does. First, because a cell whose
	// `a:tcPr` holds a `a:blipFill` holds no `a:solidFill` for the colour legs to find — and
	// `resolvedFill` would answer with the table style's banding colour, painting over the
	// image with something the source never showed.
	const picture = cell.pictureFill
	if (picture) return pictureFillOption(picture, assets, notes, CELL_PICTURE_FILL)

	const scheme = cell.fillSchemeColor
	if (scheme !== null) return { color: scheme }

	const resolved = cell.resolvedFill
	if (!resolved) return undefined
	if (hasStyle) {
		notes.note(
			'table.cell.fill',
			'dropped',
			'unread',
			"a cell's own literal fill is not separable from the one it inherits from the table style (only resolvedFill exists, and it folds in banding), so it is left to the style rather than baked in"
		)
		return undefined
	}
	return { color: literalColor(resolved.effectiveHex) }
}

/**
 * The four edge borders as the write API's `[top, right, bottom, left]` tuple.
 *
 * A suppressed edge (`a:noFill`) becomes `type: 'none'`, which is distinct from an absent
 * one — the first means "deliberately no rule here", the second means "inherit". The two
 * diagonals the read model also decodes are carried separately, by {@link cellDiagonals}.
 */
function cellBorders(cell: TableCell, notes: NoteScope): IrValue | undefined {
	const borders = cell.borders
	if (!borders) return undefined

	const edges = [borders.top, borders.right, borders.bottom, borders.left]
	if (edges.every((edge) => edge === null)) return undefined

	return edges.map((edge) => borderIr(edge, notes))
}

/** The `ST_PresetLineDashVal` tokens `BorderProps.dashType` accepts. */
const WRITABLE_DASHES = new Set([
	'solid',
	'dot',
	'dash',
	'lgDash',
	'dashDot',
	'lgDashDot',
	'lgDashDotDot',
	'sysDash',
	'sysDot',
	'sysDashDot',
	'sysDashDotDot',
])

/**
 * One decoded edge (or diagonal) as `BorderProps`.
 *
 * `type` stays the coarse on/off switch and `dashType` carries the exact preset, so a
 * `lgDashDot` rule comes back as itself rather than as the generic `sysDash` every dash
 * used to collapse onto. A dash outside `ST_PresetLineDashVal` cannot have come from a
 * conformant deck, so it is dropped to a plain dashed rule and noted.
 */
function borderIr(edge: CellBorder | null, notes: NoteScope): IrValue {
	if (!edge) return {}
	if (edge.noFill) return { type: 'none' }
	const dash = edge.dash
	const known = dash === null || WRITABLE_DASHES.has(dash)
	if (!known) {
		notes.note(
			'table.cell.borders.dash',
			'approximated',
			'unsupported',
			`a cell border's a:prstDash/@val is \`${dash}\`, which is outside ST_PresetLineDashVal; it is written as a plain dashed rule`
		)
	}
	return (
		compact({
			type: dash === null || dash === 'solid' ? 'solid' : 'dash',
			// `solid` is what an absent/solid dash already implies, so emitting it would be noise.
			dashType: known && dash !== null && dash !== 'solid' ? dash : undefined,
			color: edge.schemeColor ?? (edge.color === null ? undefined : literalColor(edge.color)),
			width: orUndefined(edge.widthPt),
		}) ?? {}
	)
}
