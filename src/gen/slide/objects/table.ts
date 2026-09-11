/**
 * ts-pptx: table slide-object serialization
 *
 * Emits a `table` slide object as a `<p:graphicFrame>` wrapping `<a:tbl>`: the merge grid
 * (`_hmerge`/`_vmerge` dummy cells for col/rowspan), the column widths, and per-cell borders,
 * fill, margins and text body.
 */

import { SlideObjectType } from '../../../enums.js'
import type { TableCellInternal } from '../../../types/internal.js'
import type { BorderProps, ObjectOptions, TableCellProps } from '../../../types/index.js'
import { checkEnumOrWarn } from '../../../ooxml/check-enum.js'
import { TEXT_HORZ_OVERFLOW } from '../../../ooxml/st-enums.js'
import { rejectEmptyColor } from '../../drawingml/color.js'
import { genXmlColorSelection } from '../../drawingml/fill.js'
import { genXmlObjectLock, GRAPHIC_FRAME_LOCK_ATTRS } from '../../drawingml/locks.js'
import { genTableCellBorderXml } from '../../drawingml/table-border.js'
import { resolveSpan, withCheckedSpans } from '../../table/spans.js'
import { genTableCell3DXml } from '../../drawingml/table-cell3d.js'
import { genXmlPlaceholder, genXmlTextBody, resolveTextAnchor } from '../../drawingml/text-body.js'
import { el, raw, voidEl, type XmlAttrs } from '../../oxml/el.js'
import {
	marginToEmu,
	resolveCellMarginsInches,
	resolveTableColWidthsEmu,
	resolveTableRowHeightEmu,
} from '../../../units-internal.js'
import { EMU_PER_INCH } from '../../../units.js'
import { type RenderContext, cNvPrOpen, graphicFrameEl } from './shared.js'
import { OOXML_NS, TABLE_GRAPHIC_DATA_URI } from '../../../ooxml/namespaces.js'
import { xsdBoolIfTrue } from '../../../ooxml/xsd-boolean.js'
import { type GridPlacement, tableColCount, walkTableGrid } from '../../table/grid.js'
import { CELL_INHERITED_KEYS } from '../../table/cell-inherit.js'
import { warn } from '../../../diagnostics.js'

/**
 * The table-level options a cell inherits when it states none of its own.
 * @see http://officeopenxml.com/drwTableCellProperties-alignment.php
 */
type TableInheritableOption = (typeof CELL_INHERITED_KEYS)[number]
type TableInheritableValue = ObjectOptions[TableInheritableOption]

/**
 * Validate a cell's `horzOverflow` before it reaches the XML, reporting and dropping anything
 * outside `ST_TextHorzOverflowType`. `undefined` (the common case) emits nothing.
 */
function resolveHorzOverflow(value: TableCellProps['horzOverflow']): string | null {
	return checkEnumOrWarn(value, TEXT_HORZ_OVERFLOW, 'table/invalid-horz-overflow', 'table cell: horzOverflow')
}

/** A cell's grid position, for deciding which of the table's outer edges it sits on. */
interface GridEdge {
	rIdx: number
	cIdx: number
	lastRow: number
	lastCol: number
}

/**
 * Overlay `TableProps.outerBorder` onto one cell's own 4-side border tuple.
 *
 * The perimeter is decided per **grid position**, not per authored cell, which is what makes
 * merges work: a colspan's origin sits at the left of its span and its covered `_hmerge`
 * cells sit to the right, so a span straddling the last column gets that column's rule on the
 * covered cell — exactly where PowerPoint defines a merged region's outer edge. A cell in the
 * table's interior touches no edge and comes back untouched.
 *
 * Returns the input array (not a copy) when nothing applies, so the common no-perimeter path
 * allocates nothing and stays byte-identical. Never mutates: `arrTabRows` holds the caller's
 * own cell objects and repeated `write()` calls must not accumulate borders.
 *
 * @param base - the cell's resolved `[top, right, bottom, left]` borders, or `null` when it has none
 * @param outer - the normalized perimeter tuple, `undefined` for a side the caller left out
 * @param at - where this cell sits in the merge grid
 * @returns the borders to emit, or `null` when there are none and the perimeter adds none
 */
function applyOuterBorder(
	base: ReadonlyArray<BorderProps | undefined> | null,
	outer: ReadonlyArray<BorderProps | undefined> | undefined,
	at: GridEdge
): ReadonlyArray<BorderProps | undefined> | null {
	if (!outer) return base
	// TRBL, matching the public tuple order.
	const onEdge = [at.rIdx === 0, at.cIdx === at.lastCol, at.rIdx === at.lastRow, at.cIdx === 0]
	const applies = ([0, 1, 2, 3] as const).filter((idx) => onEdge[idx] && outer[idx])
	if (applies.length === 0) return base
	// A missing base is a cell the definition step left with no border tuple at all, which is
	// a styled table nobody authored a border on. Its other three sides stay holes so the
	// chosen table style keeps painting them: inventing `{type:'none'}` here would erase the
	// style's interior grid the moment an `outerBorder` was added (#23).
	const merged: (BorderProps | undefined)[] = base ? [...base] : [undefined, undefined, undefined, undefined]
	for (const idx of applies) {
		const side = outer[idx]
		if (side) merged[idx] = side
	}
	return merged
}

/**
 * The options a table cell is emitted with: its own values, with the table's filled in wherever
 * the cell stated nothing.
 *
 * The guard is `=== undefined`, not a falsy test: the question is whether the CELL said
 * anything, and `false`/`0`/`''` are things it said. An older guard rescued `0` with an
 * explicit `!== 0` arm and nothing else, so a cell's `bold: false` was still overwritten by the
 * table's `bold: true`.
 *
 * `src/measure/table-fit.ts` resolves the same inheritance for the measured-fit pass, over the
 * text half of the same list -- see `gen/table/cell-inherit.ts`, which owns both halves.
 * @param cellOpts - the cell's own options, if any
 * @param tableOpts - the table's options
 * @returns a fresh bag; neither input is written to
 */
function inheritTableOptions(cellOpts: TableCellProps | undefined, tableOpts: ObjectOptions): TableCellProps {
	const merged: TableCellProps = { ...cellOpts }
	const inheritedCell = merged as Partial<Record<TableInheritableOption, TableInheritableValue>>
	const inheritedTable = tableOpts as Partial<Record<TableInheritableOption, TableInheritableValue>>
	for (const name of CELL_INHERITED_KEYS) {
		if (inheritedCell[name] === undefined && inheritedTable[name] !== undefined)
			inheritedCell[name] = inheritedTable[name]
	}
	return merged
}

/**
 * The origin cell a placement writes: the authored cell, carrying the spans the grid gave it.
 *
 * `walkTableGrid` clamps a colspan to the columns left and a rowspan to the rows left. The emitter
 * used to write the authored spans unclamped, so `[[{ text: 'A', rowspan: 3 }, 'B'], ['C']]` wrote
 * `rowSpan="3"` into a two-row table while the auto-pager and `tableLayout()` described the clamped
 * one. A span the grid had to shorten is reported and written shortened. A cell whose spans stand
 * is returned by identity, which is what its covered cells' `_spanOrigin` points at.
 * @param placement - the cell's placement on the grid
 */
function placedOrigin({ cell, rowSpan, colSpan }: GridPlacement): TableCellInternal {
	const opts = cell.options
	const statedCol = resolveSpan(opts?.colspan, 'colspan')
	const statedRow = resolveSpan(opts?.rowspan, 'rowspan')
	if (statedCol === colSpan && statedRow === rowSpan) return cell
	if (statedCol !== colSpan)
		warn(
			'table/span-out-of-range',
			`table cell: colspan ${statedCol} reaches past the table's last column; using ${colSpan}.`
		)
	if (statedRow !== rowSpan)
		warn(
			'table/span-out-of-range',
			`table cell: rowspan ${statedRow} reaches past the table's last row; using ${rowSpan}.`
		)
	return { ...cell, options: { ...opts, colspan: colSpan, rowspan: rowSpan } }
}

/**
 * The table as the grid it is written on: one cell per row and grid column.
 *
 * Every position is filled. An origin sits at its `walkTableGrid` placement, and each position its
 * span covers gets a covered cell pointing back at it: `_hmerge` right of the origin, `_vmerge`
 * below it, both where the span runs both ways. A covered cell in the origin's own row repeats the
 * origin's `rowspan`, and one in its column repeats its `colspan`, which is how PowerPoint writes a
 * merged region. A position no cell reaches (a row shorter than the grid) is a blank cell, and a cell
 * starting past the last column is dropped and reported, because a `<a:tr>` whose `<a:tc>` count
 * differs from `<a:tblGrid>` is a table PowerPoint repairs.
 *
 * Two splice passes used to insert the covered cells into the authored rows by index. They read the
 * spans unclamped, so a span past the table edge wrote cells past it:
 * `[['A', 'B'], [{ text: 'D', colspan: 3 }]]` wrote three `<a:tc>` against two `<a:gridCol>`.
 * @param rows - the table's rows, spans already range-checked
 * @param numCols - the grid's column count
 */
function buildMergeGrid(rows: TableCellInternal[][], numCols: number): TableCellInternal[][] {
	const grid: Array<Array<TableCellInternal | undefined>> = rows.map(() =>
		Array.from({ length: numCols }, (): TableCellInternal | undefined => undefined)
	)
	const placedPerRow = rows.map(() => 0)
	for (const placement of walkTableGrid(rows, numCols)) {
		placedPerRow[placement.row] = (placedPerRow[placement.row] ?? 0) + 1
		const origin = placedOrigin(placement)
		const { row, col, rowSpan, colSpan } = placement
		for (let dr = 0; dr < rowSpan; dr++) {
			const target = grid[row + dr]
			if (!target) continue
			for (let dc = 0; dc < colSpan; dc++) {
				if (dr === 0 && dc === 0) {
					target[col] = origin
					continue
				}
				const options: TableCellProps = {}
				if (dr === 0 && rowSpan > 1) options.rowspan = rowSpan
				if (dc === 0 && colSpan > 1) options.colspan = colSpan
				const covered: TableCellInternal = { _type: SlideObjectType.tablecell, options, _spanOrigin: origin }
				if (dr > 0) covered._vmerge = true
				if (dc > 0) covered._hmerge = true
				target[col + dc] = covered
			}
		}
	}
	const dropped = rows.reduce((count, row, r) => count + row.length - (placedPerRow[r] ?? 0), 0)
	if (dropped > 0)
		warn(
			'table/cell-past-grid',
			`table: ${dropped} cell(s) start past the last of the table's ${numCols} grid columns and are not written. The first row decides how many columns a table has.`
		)
	return grid.map((cells) =>
		cells.map((cell): TableCellInternal => cell ?? { _type: SlideObjectType.tablecell, text: '' })
	)
}

/**
 * Render a `table` slide object to its `<p:graphicFrame>` XML (merge-grid, row/col spans, per-cell styling).
 */
export function renderTableObject(ctx: RenderContext): string {
	const {
		obj: slideItemObj,
		shapeId,
		frame: { x, y, cx, cy },
		placeholder: placeholderObj,
		itemOpts,
	} = ctx
	// `itemOpts` is the caller's already-normalized `itemOpts` (see the dispatch in
	// `slideObjectToXml`). Read it rather than re-narrowing the field: this function has exactly
	// one call site, and a contract stated there beats a defensive re-assignment here.
	let tblInner = ''
	// Checked again here, not only in `addTableDefinition`: the merge grid below is laid out from the
	// spans, and an emitter must not size anything from a number it has not seen. Rows that came
	// through the definer are already correct, so this warns about nothing. Nothing below writes to
	// the rows, so the stored `arrTabRows` survive repeated `write()` calls without a copy.
	const arrTabRows: TableCellInternal[][] = withCheckedSpans(slideItemObj.arrTabRows ?? [])
	const objTabOpts: ObjectOptions = itemOpts
	const intColCnt = tableColCount(arrTabRows)

	// STEP 1: Start Table XML
	// NOTE: The cNvPr id must be unique among ALL shapes on the slide. A table is an
	// ordinary top-level slide object, so it uses the same `idx + 2` scheme as every other
	// object type below. The legacy `intTableNum * slide._slideNum + 1` formula could collide
	// with another shape's `idx + 2` on the same slide (e.g. a table plus enough sibling
	// shapes on slide 7), producing a duplicate id that makes PowerPoint report the file as
	// corrupt/unreadable (0x80070570) while LibreOffice silently tolerates it.
	const nvGraphicFramePr = el('p:nvGraphicFramePr', null, [
		raw(cNvPrOpen(shapeId, itemOpts.objectName, itemOpts.altText || '') + '/>'),
		raw(
			el(
				'p:cNvGraphicFramePr',
				null,
				raw(
					genXmlObjectLock(
						'a:graphicFrameLocks',
						GRAPHIC_FRAME_LOCK_ATTRS,
						{ noGrp: true, ...itemOpts.objectLock },
						itemOpts.objectName
					)
				)
			)
		),
		// A table bound to a layout placeholder emits that placeholder's <p:ph> (idx/type) so
		// PowerPoint treats the graphicFrame as filling the placeholder. The <p:ph>
		// precedes <p:extLst> per CT_ApplicationNonVisualDrawingProps document order.
		raw(
			el(
				'p:nvPr',
				null,
				[
					raw(genXmlPlaceholder(placeholderObj)),
					raw(
						el(
							'p:extLst',
							null,
							raw(
								el(
									'p:ext',
									{ uri: '{D42A27DB-BD31-4B8C-83A1-F6EECF244321}' },
									raw(voidEl('p14:modId', { 'xmlns:p14': OOXML_NS.p14, val: '1579011935' }))
								)
							)
						)
					),
				],
				{ openPrefix: '  ' }
			)
		),
	])
	// A table's box falls back to one inch on any axis the caller left undefined, rather than to
	// the zero a missing extent would otherwise emit.
	const frame = {
		x: x || (x === 0 ? 0 : EMU_PER_INCH),
		y: y || (y === 0 ? 0 : EMU_PER_INCH),
		cx: cx || (cx === 0 ? 0 : EMU_PER_INCH),
		cy: cy || EMU_PER_INCH,
	}
	{
		// NOTE: attribute ORDER is byte-significant. None of these flags appears in the byte-gate
		// baseline (zero parts each), so their emission is pinned by test/regression instead.
		const tblPrAttrs: XmlAttrs = {
			rtl: xsdBoolIfTrue(objTabOpts.rtl),
			firstRow: xsdBoolIfTrue(objTabOpts.hasHeader),
			lastRow: xsdBoolIfTrue(objTabOpts.hasFooter),
			bandRow: xsdBoolIfTrue(objTabOpts.hasBandedRows),
			bandCol: xsdBoolIfTrue(objTabOpts.hasBandedColumns),
			firstCol: xsdBoolIfTrue(objTabOpts.hasFirstColumn),
			lastCol: xsdBoolIfTrue(objTabOpts.hasLastColumn),
		}
		// `CT_TableProperties` sequences its children as EG_FillProperties, EG_EffectProperties,
		// then the tableStyle/tableStyleId choice — so a table background precedes the style id.
		// (No effects surface: PowerPoint's UI exposes no table-level effect, so a source deck
		// will not contain one and there would be nothing to reproduce.)
		rejectEmptyColor(objTabOpts.tableFill, 'tableFill')
		const tableFillXml = objTabOpts.tableFill ? genXmlColorSelection(objTabOpts.tableFill) : ''
		const tblPrChildren =
			tableFillXml + (objTabOpts.tableStyle ? el('a:tableStyleId', null, objTabOpts.tableStyle) : '')
		// Paired when it carries a fill or a style id, else self-closing — an arity difference.
		const tblPr = tblPrChildren ? el('a:tblPr', tblPrAttrs, raw(tblPrChildren)) : voidEl('a:tblPr', tblPrAttrs)
		// The `<a:tbl>` children accumulate here and are wrapped once at STEP 5, so the byte-significant
		// (and non-depth-regular) indentation on the closing tags is described in one place.
		tblInner = tblPr
	}

	// `addTableDefinition` normalizes `outerBorder` to a TRBL tuple, but a table object built by
	// hand (or replayed from an older serialized form) can still carry the single-`BorderProps`
	// form, so broadcast that here rather than silently dropping the perimeter.
	const outerBorder: ReadonlyArray<BorderProps | undefined> | undefined = !objTabOpts.outerBorder
		? undefined
		: Array.isArray(objTabOpts.outerBorder)
			? objTabOpts.outerBorder
			: [objTabOpts.outerBorder, objTabOpts.outerBorder, objTabOpts.outerBorder, objTabOpts.outerBorder]

	// STEP 2: Set column widths
	// Per-column inches from an explicit `colW` array, else split the table's
	// resolved EMU width (`cx`) evenly. `resolveTableColWidthsEmu` is the single
	// source of truth shared with the measured-fit pass. NOTE: divide the EMU
	// width, not the raw inches `options.w` — the latter collapsed auto-width
	// tables to ~0-EMU columns (e.g. `w=9` → `gridCol w="3"`).
	{
		const gridColsEmu = resolveTableColWidthsEmu(objTabOpts.colW, cx, intColCnt)
		tblInner += el(
			'a:tblGrid',
			null,
			gridColsEmu.map((w) => raw(voidEl('a:gridCol', { w })))
		)
	}

	// STEP 3: Build our row arrays into an actual grid to match the XML we will be building next
	// Note row arrays can arrive "lopsided" as in row1:[1,2,3] row2:[3] when first two cols rowspan!,
	// so a simple loop below in XML building wont suffice to build table correctly.
	// We have to build an actual grid now
	/*
					EX: (A0:rowspan=3, B1:rowspan=2, C1:colspan=2)

					/------|------|------|------\
					|  A0  |  B0  |  C0  |  D0  |
					|      |  B1  |  C1  |      |
					|      |      |  C2  |  D2  |
					\------|------|------|------/
				*/
	// Placement comes from `walkTableGrid`, the traversal the auto-pager and the measured-fit pass
	// read, so the table written is the one they paged and measured.
	const grid = buildMergeGrid(arrTabRows, intColCnt)

	// STEP 4: Build table rows/cells
	grid.forEach((cells, rIdx) => {
		// A: `rowH` pins the row; a table height provided without one is split evenly.
		// The height comes from the PLACED frame, not from `options.h`, which is why this does not
		// go through `resolveTableGridEmu` the way `pptx.tableLayout()` and the measured-fit pass
		// do: only the emitter knows the frame, and a layout placeholder overrides `h` there. The
		// row split itself is the same `resolveTableRowHeightEmu` all three read.
		// IMPORTANT: `null` (auto-height) must reach the attribute as zero for auto-sizing to work.
		const tableHeightEmu = itemOpts.h ? cy : typeof itemOpts.cy === 'number' ? itemOpts.cy : 0
		const intRowH = resolveTableRowHeightEmu(objTabOpts.rowH, rIdx, tableHeightEmu, arrTabRows.length) ?? 0

		// B: Start row — cells accumulate here and the row wraps them once, below.
		const rowCells: string[] = []

		// C: Loop over each CELL
		cells.forEach((cellObj, cIdx) => {
			const cell: TableCellInternal = cellObj
			// The grid is rectangular by now (STEP 3 filled every span with a dummy cell), so a
			// cell's index in its row *is* its grid column and the perimeter can be decided here.
			const at: GridEdge = { rIdx, cIdx, lastRow: arrTabRows.length - 1, lastCol: cells.length - 1 }

			// NOTE: attribute ORDER is byte-significant; `undefined` omits the attribute entirely,
			// which is what the old `.filter(([, v]) => !!v)` did.
			const cellSpanAttrs: XmlAttrs = {
				rowSpan: cell.options?.rowspan && cell.options.rowspan > 1 ? cell.options.rowspan : undefined,
				gridSpan: cell.options?.colspan && cell.options.colspan > 1 ? cell.options.colspan : undefined,
				vMerge: cell._vmerge ? 1 : undefined,
				hMerge: cell._hmerge ? 1 : undefined,
			}

			// 1: COLSPAN/ROWSPAN: Emit the dummy covered cell for any active span. PowerPoint defines a
			// merged region's outer edges (e.g. the right border of a colspan, the bottom border of a
			// rowspan) on the *covered* cells, so inherit the origin cell's border + fill here instead of
			// emitting an empty `<a:tcPr/>` that drops those edges.
			if (cell._hmerge || cell._vmerge) {
				const origin = cell._spanOrigin
				const originOpts = origin?.options || {}
				let spanPrXml = ''
				// Outside the `origin` guard below: a covered cell on the table's edge carries that
				// edge's rule whether or not its origin was resolvable.
				// A null tuple side means "this edge is omitted" (inherits) — normalize it to a
				// hole here, which `applyOuterBorder` already treats as "leave the edge alone".
				const originBorder = applyOuterBorder(
					Array.isArray(originOpts.border) ? originOpts.border.map((side) => side ?? undefined) : null,
					outerBorder,
					at
				)
				if (originBorder) spanPrXml += genTableCellBorderXml(originBorder)
				if (origin) {
					// Resolve the origin's fill with the same precedence the origin cell itself uses below,
					// so the whole merged region fills uniformly. This is the origin's fill *object*, so
					// an image fill arrives with its `_imgRid` already stashed and the covered cell emits
					// a `blipFill` against the same relationship.
					// NOTE: PowerPoint itself writes a bare `<a:tcPr/>` on a covered cell and repeats no
					// fill there (verified against `test/read/fixtures/table-cell-image-fill.pptx`). We
					// diverge deliberately: a covered cell is never rendered — the origin spans over it —
					// so copying the fill is invisible either way, and keeping it uniform with the solid
					// case avoids a branch that would change nothing on screen.
					rejectEmptyColor(originOpts.fill, 'cell fill')
					const spanFill = originOpts.fill || ''
					if (spanFill) spanPrXml += genXmlColorSelection(spanFill)
				}
				// NOTE: the covered cell is FLAT, unlike the real cell below, which carries indentation
				// before its `</a:tcPr>` and `</a:tc>`.
				rowCells.push(el('a:tc', cellSpanAttrs, raw(el('a:tcPr', null, raw(spanPrXml)))))
				return
			}

			// 2: OPTIONS: the cell's own values, with the table's filled in where the cell said
			// nothing. Resolved into a LOCAL bag: this used to be written back onto `cell.options`,
			// which is the stored model, so emitting the same deck twice inherited into it again and
			// the reference handed to `genXmlTextBody` below carried keys the cell never stated.
			// The invariant at the head of this file has said "never mutates" throughout.
			const cellOpts = inheritTableOptions(cell.options, objTabOpts)

			const cellValign = resolveTextAnchor(cellOpts.valign)
			const cellTextDir = cellOpts.textDirection && cellOpts.textDirection !== 'horz' ? cellOpts.textDirection : null

			rejectEmptyColor(cellOpts.fill, 'cell fill')
			const fillColor = cellOpts.fill || ''
			const cellFill = fillColor ? genXmlColorSelection(fillColor) : ''

			const cellMargin = resolveCellMarginsInches(cellOpts.margin)
			// Cell text ALWAYS wraps — PowerPoint has no per-cell no-wrap, so there is nothing to emit
			// for one. `wrap="none"` on a cell's `<a:bodyPr>` renders inert and is stripped on the next
			// save, and `TextFrame.WordWrap` is read-only on a cell over COM (it reports msoTrue whatever
			// the XML says). `horzOverflow` below is NOT that switch: per ECMA-376 §20.1.10.68 it decides
			// whether a single glyph too wide for the line is clipped or draws past the cell edge.
			// Probe for both: `test/read/fixtures/authoring/probe-table-cell-wrap.ps1`
			const cellHorzOverflow = resolveHorzOverflow(cellOpts.horzOverflow)

			// Cell margins are inches (see `marginToEmu`); a `>= 1` value warns once as a likely legacy points value.
			// NOTE: attribute ORDER is byte-significant (margins, then anchor, then vert). `horzOverflow`
			// goes last, which is both where CT_TableCellProperties declares it and where PowerPoint
			// writes it, so an existing cell's bytes are unchanged while it stays unset. `anchorCtr`
			// follows `anchor`, matching the schema's own adjacency; that the emitter's `anchor`/`vert`
			// pair is inverted relative to CT_TableCellProperties is pre-existing and harmless — XML
			// attributes are unordered, so the order is only about keeping emitted bytes stable.
			const tcPrAttrs: XmlAttrs = {
				marL: marginToEmu(cellMargin[3]),
				marR: marginToEmu(cellMargin[1]),
				marT: marginToEmu(cellMargin[0]),
				marB: marginToEmu(cellMargin[2]),
				anchor: cellValign,
				// `false` is the schema default, so only `true` is worth writing.
				anchorCtr: xsdBoolIfTrue(cellOpts.anchorCtr),
				vert: cellTextDir,
				horzOverflow: cellHorzOverflow,
			}

			// 4: Set CELL content and properties; 5: borders; 6: fill ==============
			// The trailing indentation before `</a:tcPr>` and `</a:tc>` is byte-significant.
			// Child order is the CT_TableCellProperties sequence: the four edges, the two diagonals,
			// `cell3D`, then the fill. Unlike the edges, the diagonals are NOT copied onto a merged
			// region's covered cells (see `genTableCellBorderXml`).
			// A null tuple side is *omitted* (inherits) — normalize it to a hole, which
			// `applyOuterBorder` already treats as "leave the edge alone".
			const cellBorder = applyOuterBorder(
				Array.isArray(cellOpts.border) ? cellOpts.border.map((side) => side ?? undefined) : null,
				outerBorder,
				at
			)
			const cellDiagonal = cellOpts.diagonal
			const cellBorderXml = cellBorder || cellDiagonal ? genTableCellBorderXml(cellBorder ?? [], cellDiagonal) : ''
			rowCells.push(
				el(
					'a:tc',
					cellSpanAttrs,
					[
						raw(genXmlTextBody({ ...cell, options: cellOpts })),
						raw(
							el('a:tcPr', tcPrAttrs, [raw(cellBorderXml), raw(genTableCell3DXml(cellOpts.cell3D)), raw(cellFill)], {
								closePrefix: '  ',
							})
						),
					],
					{ closePrefix: ' ' }
				)
			)
		})

		// D: Complete row
		tblInner += el('a:tr', { h: intRowH }, rowCells.map(raw))
	})

	// STEP 5: Complete table. NOTE: the closing tags carry indentation the opening tags do not,
	// so each `closePrefix` is stated explicitly rather than derived from depth.
	return graphicFrameEl({
		nvGraphicFramePr,
		frame,
		uri: TABLE_GRAPHIC_DATA_URI,
		payload: el('a:tbl', null, raw(tblInner), { closePrefix: '      ' }),
		fmt: { graphic: { closePrefix: '  ' }, graphicData: { closePrefix: '    ' } },
	})
}
