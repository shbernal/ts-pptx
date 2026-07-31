/**
 * ts-pptx: table slide-object serialization
 *
 * Emits a `table` slide object as a `<p:graphicFrame>` wrapping `<a:tbl>`: the merge grid
 * (`_hmerge`/`_vmerge` dummy cells for col/rowspan), the column widths, and per-cell borders,
 * fill, margins and text body.
 */

import { SlideObjectType } from '../../../enums.js'
import { DEF_CELL_MARGIN_IN } from '../../../constants-internal.js'
import type { BorderProps, ObjectOptions, TableCell, TableCellProps } from '../../../types/index.js'
import type { SlideObject } from '../../../types/internal.js'
import { warnOnce } from '../../../diagnostics.js'
import { genXmlColorSelection } from '../../drawingml/fill.js'
import { genXmlObjectLock, GRAPHIC_FRAME_LOCK_ATTRS } from '../../drawingml/locks.js'
import { genTableCellBorderXml } from '../../drawingml/table-border.js'
import { genXmlPlaceholder, genXmlTextBody } from '../../drawingml/text-body.js'
import { el, raw, voidEl, type XmlAttrs } from '../../oxml/el.js'
import { inch2Emu, marginToEmu, resolveTableColWidthsEmu } from '../../../units-internal.js'
import { EMU_PER_INCH } from '../../../units.js'
import { cNvPrOpen, P14_NS } from './shared.js'

type TableInheritableOption =
	| 'align'
	| 'bold'
	| 'border'
	| 'color'
	| 'fill'
	| 'fontFace'
	| 'fontSize'
	| 'margin'
	| 'textDirection'
	| 'underline'
	| 'valign'
type TableInheritableValue = ObjectOptions[TableInheritableOption]

/** The two values `ST_TextHorzOverflowType` allows on `a:tcPr/@horzOverflow`. */
const HORZ_OVERFLOW_VALUES: readonly string[] = ['clip', 'overflow']

/**
 * Validate a cell's `horzOverflow` before it reaches the XML. Anything outside
 * `ST_TextHorzOverflowType` would make the slide part schema-invalid — and PowerPoint
 * reports that as a corrupt file, not as a mis-set option — so an unrecognized value is
 * reported and dropped instead of written. `undefined` (the common case) emits nothing.
 */
function resolveHorzOverflow(value: TableCellProps['horzOverflow']): string | null {
	if (value === undefined || value === null) return null
	if (HORZ_OVERFLOW_VALUES.includes(value)) return value
	warnOnce(
		'table/invalid-horz-overflow',
		`table cell: horzOverflow \`${String(value)}\` is not a valid value and is ignored — ` +
			`use ${HORZ_OVERFLOW_VALUES.join(' or ')}.`,
		{ received: value, valid: HORZ_OVERFLOW_VALUES }
	)
	return null
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
	base: BorderProps[] | null,
	outer: ReadonlyArray<BorderProps | undefined> | undefined,
	at: GridEdge
): BorderProps[] | null {
	if (!outer) return base
	// TRBL, matching the public tuple order.
	const onEdge = [at.rIdx === 0, at.cIdx === at.lastCol, at.rIdx === at.lastRow, at.cIdx === 0]
	const applies = ([0, 1, 2, 3] as const).filter((idx) => onEdge[idx] && outer[idx])
	if (applies.length === 0) return base
	// A cell with no borders of its own still needs the other three sides spelled out, since
	// `genTableCellBorderXml` reads a dense tuple; `{type:'none'}` is what the definition step
	// already puts on an unstyled cell, so this matches rather than invents.
	const merged: BorderProps[] = base
		? [...base]
		: [{ type: 'none' }, { type: 'none' }, { type: 'none' }, { type: 'none' }]
	for (const idx of applies) {
		const side = outer[idx]
		if (side) merged[idx] = side
	}
	return merged
}

/**
 * Render a `table` slide object to its `<p:graphicFrame>` XML (merge-grid, row/col spans, per-cell styling).
 */
export function renderTableObject(
	slideItemObj: SlideObject,
	idx: number,
	x: number,
	y: number,
	cx: number,
	cy: number,
	placeholderObj: SlideObject | null,
	itemOpts: ObjectOptions
): string {
	// Caller guarantees options is set (see slideObjectToXml); re-narrow for this scope.
	slideItemObj.options = slideItemObj.options || {}
	let strXml: string
	let arrTabRows: TableCell[][] = []
	let objTabOpts: ObjectOptions = {}
	let intColCnt = 0
	let tblInner = ''
	let cellOpts: TableCellProps | null = null
	// Shallow-clone each row so splice() in the merge-grid builder does not mutate the stored
	// arrTabRows, which would corrupt output on repeated write()/writeFile() calls.
	arrTabRows = (slideItemObj.arrTabRows ?? []).map((row) => [...row])
	objTabOpts = slideItemObj.options
	intColCnt = 0

	// Calc number of columns
	// NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
	// ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
	;(arrTabRows[0] ?? []).forEach((cell) => {
		cellOpts = cell.options || null
		intColCnt += cellOpts?.colspan ? Number(cellOpts.colspan) : 1
	})

	// STEP 1: Start Table XML
	// NOTE: The cNvPr id must be unique among ALL shapes on the slide. A table is an
	// ordinary top-level slide object, so it uses the same `idx + 2` scheme as every other
	// object type below. The legacy `intTableNum * slide._slideNum + 1` formula could collide
	// with another shape's `idx + 2` on the same slide (e.g. a table plus enough sibling
	// shapes on slide 7), producing a duplicate id that makes PowerPoint report the file as
	// corrupt/unreadable (0x80070570) while LibreOffice silently tolerates it.
	strXml =
		'<p:graphicFrame><p:nvGraphicFramePr>' +
		cNvPrOpen(idx + 2, slideItemObj.options.objectName, slideItemObj.options.altText || '') +
		'/>'
	strXml +=
		el(
			'p:cNvGraphicFramePr',
			null,
			raw(
				genXmlObjectLock(
					'a:graphicFrameLocks',
					GRAPHIC_FRAME_LOCK_ATTRS,
					{ noGrp: true, ...slideItemObj.options.objectLock },
					slideItemObj.options.objectName
				)
			)
		) +
		// A table bound to a layout placeholder emits that placeholder's <p:ph> (idx/type) so
		// PowerPoint treats the graphicFrame as filling the placeholder. The <p:ph>
		// precedes <p:extLst> per CT_ApplicationNonVisualDrawingProps document order.
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
								raw(voidEl('p14:modId', { 'xmlns:p14': P14_NS, val: '1579011935' }))
							)
						)
					)
				),
			],
			{ openPrefix: '  ' }
		) +
		'</p:nvGraphicFramePr>'
	strXml += el('p:xfrm', null, [
		raw(voidEl('a:off', { x: x || (x === 0 ? 0 : EMU_PER_INCH), y: y || (y === 0 ? 0 : EMU_PER_INCH) })),
		raw(voidEl('a:ext', { cx: cx || (cx === 0 ? 0 : EMU_PER_INCH), cy: cy || EMU_PER_INCH })),
	])
	{
		// NOTE: attribute ORDER is byte-significant. None of these flags appears in the byte-gate
		// baseline (zero parts each), so their emission is pinned by test/regression instead.
		const tblPrAttrs: XmlAttrs = {
			rtl: objTabOpts.rtl ? '1' : null,
			firstRow: objTabOpts.hasHeader ? '1' : null,
			lastRow: objTabOpts.hasFooter ? '1' : null,
			bandRow: objTabOpts.hasBandedRows ? '1' : null,
			bandCol: objTabOpts.hasBandedColumns ? '1' : null,
			firstCol: objTabOpts.hasFirstColumn ? '1' : null,
			lastCol: objTabOpts.hasLastColumn ? '1' : null,
		}
		// Paired when a style id is carried, else self-closing — an arity difference.
		const tblPr = objTabOpts.tableStyle
			? el('a:tblPr', tblPrAttrs, raw(el('a:tableStyleId', null, objTabOpts.tableStyle)))
			: voidEl('a:tblPr', tblPrAttrs)
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
	// A: add _hmerge cell for colspan. should reserve rowspan
	arrTabRows.forEach((cells) => {
		for (let cIdx = 0; cIdx < cells.length;) {
			const cell = cells[cIdx]
			if (!cell) break
			const colspan = cell.options?.colspan
			const rowspan = cell.options?.rowspan
			if (colspan && colspan > 1) {
				const vMergeCells = new Array(colspan - 1).fill(undefined).map(() => {
					return {
						_type: SlideObjectType.tablecell,
						options: { rowspan },
						_hmerge: true,
						_spanOrigin: cell,
					} as const
				})
				cells.splice(cIdx + 1, 0, ...vMergeCells)
				cIdx += colspan
			} else {
				cIdx += 1
			}
		}
	})
	// B: add _vmerge cell for rowspan. should reserve colspan/_hmerge
	arrTabRows.forEach((cells, rIdx) => {
		const nextRow = arrTabRows[rIdx + 1]
		if (!nextRow) return
		cells.forEach((cell, cIdx) => {
			const rowspan = cell._rowContinue || cell.options?.rowspan
			const colspan = cell.options?.colspan
			const _hmerge = cell._hmerge
			if (rowspan && rowspan > 1) {
				// Point back to the true origin cell: when `cell` is itself an `_hmerge` dummy
				// (combined colspan+rowspan), use its origin rather than the dummy.
				const _spanOrigin = cell._spanOrigin || cell
				const hMergeCell = {
					_type: SlideObjectType.tablecell,
					options: { colspan },
					_rowContinue: rowspan - 1,
					_vmerge: true,
					_hmerge,
					_spanOrigin,
				} as const
				nextRow.splice(cIdx, 0, hMergeCell)
			}
		})
	})

	// STEP 4: Build table rows/cells
	arrTabRows.forEach((cells, rIdx) => {
		// A: Table Height provided without rowH? Then distribute rows
		let intRowH = 0 // IMPORTANT: Default must be zero for auto-sizing to work
		if (Array.isArray(objTabOpts.rowH) && objTabOpts.rowH[rIdx]) intRowH = inch2Emu(Number(objTabOpts.rowH[rIdx]))
		else if (objTabOpts.rowH && !isNaN(Number(objTabOpts.rowH))) intRowH = inch2Emu(Number(objTabOpts.rowH))
		else if (itemOpts.cy || itemOpts.h) {
			// `cy` already holds the table height resolved to EMU (line ~276), correctly handling
			// inches/percent/unit-string inputs — reuse it rather than re-parsing options.h.
			intRowH = Math.round((itemOpts.h ? cy : typeof itemOpts.cy === 'number' ? itemOpts.cy : 1) / arrTabRows.length)
		}

		// B: Start row — cells accumulate here and the row wraps them once, below.
		const rowCells: string[] = []

		// C: Loop over each CELL
		cells.forEach((cellObj, cIdx) => {
			const cell: TableCell = cellObj
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
				const originBorder = applyOuterBorder(
					Array.isArray(originOpts.border) ? originOpts.border : null,
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
					const spanFill = originOpts.fill || ''
					if (spanFill) spanPrXml += genXmlColorSelection(spanFill)
				}
				// NOTE: the covered cell is FLAT, unlike the real cell below, which carries indentation
				// before its `</a:tcPr>` and `</a:tc>`.
				rowCells.push(el('a:tc', cellSpanAttrs, raw(el('a:tcPr', null, raw(spanPrXml)))))
				return
			}

			// 2: OPTIONS: Build/set cell options
			const cellOpts = cell.options || {}
			cell.options = cellOpts

			// B: Inherit some options from table when cell options dont exist
			// @see: http://officeopenxml.com/drwTableCellProperties-alignment.php
			const inheritedCellOpts = cellOpts as Partial<Record<TableInheritableOption, TableInheritableValue>>
			const inheritedTableOpts = objTabOpts as Partial<Record<TableInheritableOption, TableInheritableValue>>
			;(
				[
					'align',
					'bold',
					'border',
					'color',
					'fill',
					'fontFace',
					'fontSize',
					'margin',
					'textDirection',
					'underline',
					'valign',
				] as const
			).forEach((name) => {
				if (inheritedTableOpts[name] && !inheritedCellOpts[name] && inheritedCellOpts[name] !== 0)
					inheritedCellOpts[name] = inheritedTableOpts[name]
			})

			const cellValign = cellOpts.valign
				? cellOpts.valign
						.replace(/^c$/i, 'ctr')
						.replace(/^m$/i, 'ctr')
						.replace('center', 'ctr')
						.replace('middle', 'ctr')
						.replace('top', 't')
						.replace('btm', 'b')
						.replace('bottom', 'b')
				: null
			const cellTextDir = cellOpts.textDirection && cellOpts.textDirection !== 'horz' ? cellOpts.textDirection : null

			const fillColor = cellOpts.fill || ''
			const cellFill = fillColor ? genXmlColorSelection(fillColor) : ''

			let cellMargin = cellOpts.margin === 0 || cellOpts.margin ? cellOpts.margin : DEF_CELL_MARGIN_IN
			if (!Array.isArray(cellMargin) && typeof cellMargin === 'number')
				cellMargin = [cellMargin, cellMargin, cellMargin, cellMargin]
			// defensive fallback - if `cellMargin` is not a 4-element array of finite numbers, use defaults (prevents NaN in marL/R/T/B)
			if (
				!Array.isArray(cellMargin) ||
				cellMargin.length !== 4 ||
				cellMargin.some((v) => typeof v !== 'number' || !isFinite(v))
			) {
				cellMargin = DEF_CELL_MARGIN_IN
			}
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
			// writes it, so an existing cell's bytes are unchanged while it stays unset.
			const tcPrAttrs: XmlAttrs = {
				marL: marginToEmu(cellMargin[3]),
				marR: marginToEmu(cellMargin[1]),
				marT: marginToEmu(cellMargin[0]),
				marB: marginToEmu(cellMargin[2]),
				anchor: cellValign,
				vert: cellTextDir,
				horzOverflow: cellHorzOverflow,
			}

			// 4: Set CELL content and properties; 5: borders; 6: fill ==============
			// The trailing indentation before `</a:tcPr>` and `</a:tc>` is byte-significant.
			const cellBorder = applyOuterBorder(Array.isArray(cellOpts.border) ? cellOpts.border : null, outerBorder, at)
			rowCells.push(
				el(
					'a:tc',
					cellSpanAttrs,
					[
						raw(genXmlTextBody(cell)),
						raw(
							el('a:tcPr', tcPrAttrs, [cellBorder ? raw(genTableCellBorderXml(cellBorder)) : null, raw(cellFill)], {
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
	strXml += el(
		'a:graphic',
		null,
		raw(
			el(
				'a:graphicData',
				{ uri: 'http://schemas.openxmlformats.org/drawingml/2006/table' },
				raw(el('a:tbl', null, raw(tblInner), { closePrefix: '      ' })),
				{ closePrefix: '    ' }
			)
		),
		{ closePrefix: '  ' }
	)
	strXml += '</p:graphicFrame>'

	return strXml
}
