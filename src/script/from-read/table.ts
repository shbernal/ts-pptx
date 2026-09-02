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
import {
	ANCHOR_TO_VALIGN,
	compact,
	inches,
	literalColor,
	nameOf,
	orUndefined,
	positionOptions,
	schemeColorOption,
	WRITABLE_DASHES,
} from './values.js'
import { pictureFillOption, type PictureFillSubject } from './picture-fill.js'
import { gradientStops, patternOption } from './surface-fill.js'
import { runOptions, textRuns } from './text.js'

/** How {@link pictureFillOption}'s notes name a cell. */
const CELL_PICTURE_FILL: PictureFillSubject = {
	construct: 'table.cell.fill.picture',
	subject: 'this table cell',
	element: 'a:tcPr/a:blipFill',
}

/** How {@link pictureFillOption}'s notes name the table's own background. */
const TABLE_PICTURE_FILL: PictureFillSubject = {
	construct: 'table.fill.picture',
	subject: 'this table',
	element: 'a:tblPr/a:blipFill',
}

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
	// zero-height row, and neither the mixed nor the all-auto case survives the write path.
	// Nothing is visibly wrong either way; what is lost is the *implicitness*, which matters
	// the moment someone edits a cell and expects the row to grow.
	//
	// The all-auto case is the more common of the two — a table authored without explicit row
	// heights has `a:tr/@h="0"` on every row — and it went undeclared for longer precisely
	// because it looks like the harmless one. It is the case where `rowH` is omitted entirely
	// below, and `addTable` then divides the frame height evenly rather than leaving the rows
	// auto, so three auto rows come back pinned to a third of the frame each.
	const autoRows = rowHeights.filter((height) => height === 0).length
	if (autoRows === rowHeights.length && autoRows > 0) {
		notes.note(
			'table.rowAuto',
			'approximated',
			'unsupported',
			'every row is auto-height (a:tr/@h of 0), so no rowH is emitted and addTable divides the frame height evenly among the rows; the rows come back pinned to that even split rather than sized to their content'
		)
	} else if (autoRows > 0) {
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
		...positionOptions(frame, notes),
		objectName: frame.name || undefined,
		// The source GUID resolves against the destination's own tableStyles.xml, which a
		// template-anchored output carries over intact.
		tableStyle: orUndefined(styleId),
		// `tableFill`, not `fill`: the source has one fill on `a:tblPr`, and `fill` would
		// flatten it into a copy on every cell, which is a different package for the same
		// picture. See `TableProps.tableFill`.
		tableFill: tableFill(table, notes, assets),
		hasHeader: table.firstRowHeader ? true : undefined,
		hasBandedRows: table.bandedRows ? true : undefined,
		colW: columnWidths.every((w) => w === null) ? undefined : columnWidths.map((w) => inches(w ?? 0)),
		rowH: rowHeights.every((h) => h === 0) ? undefined : rowHeights.map(inches),
	})

	return { method: 'addTable', args: [rows, options ?? {}], ...nameOf(frame) }
}

/**
 * The table's own background (`a:tblPr` fill).
 *
 * Simpler than {@link cellFill}, and for one reason: a `a:tblPr` fill can only be the
 * table's own. There is no style graph feeding into it, so there is no
 * own-versus-inherited ambiguity to be careful about and the resolved colour can be used
 * directly. A scheme token still wins over the literal when there is one, so the replica
 * keeps tracking its theme.
 */
function tableFill(table: Table, notes: NoteScope, assets: AssetResolver): IrValue | undefined {
	const gradient = table.gradientFill
	if (gradient) {
		const stops = gradientStops(gradient, notes, 'table.fill')
		if (stops) return { type: 'gradient', gradient: stops }
	}

	const pattern = patternOption(table.patternFill)
	if (pattern) return pattern

	const picture = table.pictureFill
	if (picture) return pictureFillOption(picture, assets, notes, TABLE_PICTURE_FILL)

	const resolved = table.resolvedFill
	const scheme = schemeColorOption(
		table.fillSchemeColor,
		resolved?.effectiveHex ?? null,
		notes,
		'table.fill.schemeToken',
		'table fill'
	)
	if (scheme !== undefined) return { color: scheme }

	return resolved ? { color: literalColor(resolved.effectiveHex) } : undefined
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

	// `a:tc/@id` and `a:tcPr/a:headers` are the one pair here that is *deliberately* not
	// mapped. PowerPoint strips both on save (probe:
	// `test/read/fixtures/authoring/probe-table-cell-a11y-and-3d.ps1`), so there is no write
	// option to map them to and adding one would ship a feature that dies on the first save.
	// A source deck from another producer can still carry them, so the loss is recorded.
	if (cell.id !== null || cell.headerIds.length > 0) {
		notes.note(
			'table.cell.headers',
			'dropped',
			'unwritable',
			'this cell carries a screen-reader header association (a:tc/@id / a:tcPr/a:headers); PowerPoint strips both on save, so there is no write option and the association is lost — hasHeader (a:tblPr/@firstRow) is the marker that survives'
		)
	}

	const options = compact({
		fill: cellFill(cell, hasStyle, notes, assets),
		border: cellBorders(cell, notes),
		diagonal: cellDiagonals(cell, notes),
		anchorCtr: cell.anchorCtr ? true : undefined,
		cell3D: cellThreeD(cell),
		valign: anchor === null ? undefined : ANCHOR_TO_VALIGN[anchor],
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

/**
 * The two corner-to-corner rules as `TableCellProps.diagonal`.
 *
 * Kept out of {@link cellBorders}' tuple for the same reason the write option is: they are
 * not edges. Every dash and colour a diagonal can carry is one an edge can, so they share
 * {@link borderIr} — including its out-of-enum dash note.
 */
function cellDiagonals(cell: TableCell, notes: NoteScope): IrValue | undefined {
	const borders = cell.borders
	if (!borders || (!borders.tlToBr && !borders.blToTr)) return undefined
	return compact({
		tlToBr: borders.tlToBr ? borderIr(borders.tlToBr, notes) : undefined,
		blToTr: borders.blToTr ? borderIr(borders.blToTr, notes) : undefined,
	})
}

/**
 * A cell's `a:cell3D` as `TableCellProps.cell3D`.
 *
 * Every field is passed through as read: the write path vets each against its `ST_` union
 * before emission, so a value from a source deck that somehow falls outside one is caught
 * there rather than being screened twice with two chances to disagree. Sizes are already in
 * points on both sides.
 */
function cellThreeD(cell: TableCell): IrValue | undefined {
	const cell3D = cell.cell3D
	if (!cell3D) return undefined
	const rig = cell3D.lightRig
	return (
		compact({
			preset: orUndefined(cell3D.bevel.preset),
			width: orUndefined(cell3D.bevel.widthPt),
			height: orUndefined(cell3D.bevel.heightPt),
			material: orUndefined(cell3D.material),
			// CT_LightRig requires both, so a half-specified one from a source deck is dropped
			// rather than passed on for the emitter to reject.
			lightRig: rig?.rig && rig.dir ? { rig: rig.rig, dir: rig.dir } : undefined,
		}) ?? {}
	)
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
 * A cell's fill.
 *
 * The question here is not "what colour is this cell" but "what did the *source* say", and
 * the two differ for a styled table: `resolvedFill` answers the first, folding the cell's
 * own `a:solidFill` together with the colour it merely inherits from the style's
 * header/banding rules. Writing an inherited colour back would turn every banded cell into
 * an explicitly-filled one — the copy looks right until someone changes its table style, and
 * then nothing moves.
 *
 * {@link TableCell.hasOwnFill} separates them: an `a:tcPr` either carries an
 * `EG_FillProperties` child or it does not. So a cell's own fill is emitted (in whatever
 * form it took), and a cell with none is left to the style GUID, which reproduces the
 * banding exactly rather than approximately. Neither case loses anything, which is why
 * neither records a note.
 */
function cellFill(cell: TableCell, hasStyle: boolean, notes: NoteScope, assets: AssetResolver): IrValue | undefined {
	// An explicit `a:noFill` is a statement, and the one `EG_FillProperties` member whose
	// loss is invisible further down: a suppressed cell reports `null` from every colour
	// accessor, so without this it falls out of the bottom as "no fill option" and the copy
	// takes the style's banding — the opposite of what the source shows. Same distinction
	// `lineOption` makes for `lineNoFill` on the shape side.
	if (cell.fillNoFill) return { type: 'none' }

	// Every non-solid choice comes next, and for one shared reason: a cell whose `a:tcPr`
	// holds a `a:gradFill`/`a:pattFill`/`a:blipFill` holds no `a:solidFill` for the colour
	// legs below to find, so `resolvedFill` would answer with the table style's banding
	// colour and paint over the cell with something the source never showed.
	const gradient = cell.gradientFill
	if (gradient) {
		const stops = gradientStops(gradient, notes, 'table.cell.fill')
		if (stops) return { type: 'gradient', gradient: stops }
	}

	const pattern = patternOption(cell.patternFill)
	if (pattern) return pattern

	const picture = cell.pictureFill
	if (picture) return pictureFillOption(picture, assets, notes, CELL_PICTURE_FILL)

	const scheme = schemeColorOption(
		cell.fillSchemeColor,
		cell.resolvedFill?.effectiveHex ?? null,
		notes,
		'table.cell.fill.schemeToken',
		'cell fill'
	)
	if (scheme !== undefined) return { color: scheme }

	// A styled cell with no fill of its own takes the style's banding, and the style GUID
	// travels with the table — so emitting nothing here is not a loss, it is what keeps the
	// copy responsive to its own style.
	if (hasStyle && !cell.hasOwnFill) return undefined

	const resolved = cell.resolvedFill
	if (!resolved) return undefined
	// Reached only when the cell's fill IS its own: either there is no style to inherit from,
	// or `hasOwnFill` said so. The literal is exact for a `srgbClr`, and for the rarer colour
	// models (`a:sysClr`, `a:prstClr`, `a:hslClr`, …) it is the colour they resolve to, which
	// is the closest the write API's hex option can come.
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
			color:
				schemeColorOption(
					edge.schemeColor,
					edge.resolvedColor?.effectiveHex ?? null,
					notes,
					'table.cell.borders.schemeToken',
					'cell border'
				) ?? (edge.color === null ? undefined : literalColor(edge.color)),
			width: orUndefined(edge.widthPt),
		}) ?? {}
	)
}
