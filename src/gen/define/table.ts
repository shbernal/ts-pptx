/**
 * ts-pptx: Table Definition
 *
 * `addTableDefinition` applies the `headerRow` / `columns` sugar, normalizes rows into
 * fully-resolved `TableCell`s (incl. 4-side borders), computes width, and — when `autoPage` is
 * set — shreds the table across overflow slides via `getSlidesForTableRows`.
 */
import { SlideObjectType } from '../../enums.js'
import {
	DEF_CELL_BORDER,
	DEF_CELL_MARGIN_IN,
	DEF_FONT_COLOR,
	DEF_FONT_SIZE,
	DEF_SLIDE_MARGIN_IN,
} from '../../constants-internal.js'
import { warn } from '../../diagnostics.js'
import type {
	AddSlideProps,
	BorderProps,
	PresLayout,
	ShapeFillProps,
	TableCell,
	TableProps,
	TableRow,
} from '../../types/index.js'
import type { PresSlideInternal, SlideLayoutInternal } from '../../types/internal.js'
import { getSlidesForTableRows } from '../table/autopage.js'
import { encodeXmlAttrValue, validateObjectName } from '../utils.js'
import { getSmartParseNumber } from '../../units-internal.js'
import { EMU_PER_INCH } from '../../units.js'
import { createHyperlinkRels } from './hyperlinks.js'
import { registerImageFillMedia } from './image.js'
import { InvalidOptionError } from '../../errors.js'

/** A per-cell TRBL border tuple; a null side is *omitted* (inherits), not erased. */
type BorderTuple = [BorderProps | null, BorderProps | null, BorderProps | null, BorderProps | null]
type OuterBorderTuple = [BorderProps?, BorderProps?, BorderProps?, BorderProps?]

/**
 * Expand a cell/table border into the 4-side tuple the rest of the pipeline expects.
 * A single `BorderProps` is broadcast to all four sides `[top, right, bottom, left]`;
 * an already-4-element `BorderTuple` passes through unchanged.
 * @param border - one border style (applied to every side) or a per-side tuple
 * @return {BorderTuple} a 4-element `[top, right, bottom, left]` tuple
 */
function normalizeBorderTuple(border: BorderProps | BorderTuple): BorderTuple {
	return Array.isArray(border) ? border : [border, border, border, border]
}

/** Fill a border's defaulted keys, preserving anything else the caller set (`cap`, `dashType`, …). */
function withBorderDefaults(side: BorderProps): BorderProps {
	return {
		...side,
		type: side.type || DEF_CELL_BORDER.type,
		color: side.color || DEF_CELL_BORDER.color,
		width: typeof side.width === 'number' ? side.width : DEF_CELL_BORDER.width,
	}
}

/**
 * Resolve `TableProps.outerBorder` into the `[top, right, bottom, left]` tuple the emitter
 * applies to the table's edge cells, with `undefined` for a side the caller left out.
 *
 * A sparse side is *not* the same as `{ type: 'none' }`: it means "leave this edge to
 * whatever `border` (or the cell's own `options.border`) already drew", whereas `'none'` is
 * an explicit instruction to erase the rule there. Keeping the hole is what lets
 * `outerBorder: [rule, undefined, rule, undefined]` add rules above and below a table
 * without also clearing its left and right edges.
 */
function normalizeOuterBorder(outer: TableProps['outerBorder']): OuterBorderTuple | undefined {
	if (!outer || typeof outer !== 'object') return undefined
	const sides: OuterBorderTuple = Array.isArray(outer) ? outer : [outer, outer, outer, outer]
	const resolved = ([0, 1, 2, 3] as const).map((idx) => {
		const side = sides[idx]
		return side && typeof side === 'object' ? withBorderDefaults(side) : undefined
	}) as OuterBorderTuple
	return resolved.some((side) => side !== undefined) ? resolved : undefined
}

/**
 * Adds a table object to a slide definition.
 * @param {PresSlideInternal} target - slide object that the table should be added to
 * @param {TableRow[]} tableRows - table data
 * @param {TableProps} options - table options
 * @param {SlideLayoutInternal} slideLayout - Slide layout
 * @param {PresLayout} presLayout - Presentation layout
 * @param {Function} addSlide - method
 * @param {Function} getSlide - method
 */
/**
 * Apply the `headerRow` / `columns` inline-styling sugar: bake blanket header/column
 * formatting into per-cell options so it flows through the normal cell pipeline. Sets
 * `opt.hasHeader` when `headerRow` implies it. Returns the (possibly shallow-copied) rows.
 */
function applyTableHeaderColumnSugar(tableRows: TableRow[], opt: TableProps): TableRow[] {
	const hdr = opt.headerRow && typeof opt.headerRow === 'object' ? opt.headerRow : undefined
	const cols = Array.isArray(opt.columns) && opt.columns.length ? opt.columns : undefined
	let srcRows: TableRow[] = tableRows
	if ((hdr || cols) && Array.isArray(tableRows[0])) {
		if (hdr && opt.hasHeader === undefined) opt.hasHeader = true
		srcRows = tableRows.map((row, rowIdx) => {
			if (!Array.isArray(row)) return row
			// Column-scoped defaults only apply when we actually have `columns`; the header row
			// alone is a cheaper positional map. Skip untouched body rows to avoid needless copies.
			if (!cols && rowIdx !== 0) return row
			let colCursor = 0
			return row.map((cell: number | string | TableCell): TableCell => {
				const cellObj: TableCell =
					typeof cell === 'string' || typeof cell === 'number'
						? { text: String(cell), options: {} }
						: { ...cell, options: { ...cell.options } }
				const colDef = cols ? cols[colCursor] : undefined
				colCursor += cellObj.options?.colspan || 1
				cellObj.options = {
					...(colDef && typeof colDef === 'object' ? colDef : {}),
					...(rowIdx === 0 && hdr ? hdr : {}),
					...cellObj.options,
				}
				return cellObj
			})
		})
	}
	return srcRows
}

/**
 * Transform loosely-typed table rows (strings / numbers / TableCell) into a grid of
 * well-formed TableCell objects with fully-resolved 4-side cell borders.
 *
 * @param srcRows - the rows as authored, after the `headerRow`/`columns` sugar
 * @param opt - the table's options
 */
function normalizeTableRows(srcRows: TableRow[], opt: TableProps): TableCell[][] {
	const arrRows: TableCell[][] = []
	srcRows.forEach((row, idx) => {
		const newRow: TableCell[] = []

		if (Array.isArray(row)) {
			row.forEach((cell: number | string | TableCell) => {
				// A:
				const newCellOptions = typeof cell === 'object' && cell.options ? cell.options : {}
				const newCell: TableCell = {
					_type: SlideObjectType.tablecell,
					text: '',
					options: newCellOptions,
				}

				// B:
				if (typeof cell === 'string' || typeof cell === 'number') newCell.text = cell.toString()
				else if (cell.text) {
					// Cell can contain complex text type, or string, or number
					if (typeof cell.text === 'string' || typeof cell.text === 'number') newCell.text = cell.text.toString()
					else if (cell.text) newCell.text = cell.text
					// Capture options
					if (cell.options && typeof cell.options === 'object') newCell.options = cell.options
				}

				// C: Set cell borders
				// A side nobody asked for is written as an explicit `{type:'none'}`. That is direct
				// formatting, so it beats any border a built-in table style would draw — which is
				// what keeps an unstyled table free of grid lines, PowerPoint's own no-style look
				// being a black hairline grid. Whatever the caller *did* author, on the cell or on
				// the table, is untouched.
				//
				// A table that names a `tableStyle` asked for that style's grid, so the default is
				// skipped there and the edges stay absent: direct formatting beats the style, and
				// filling in four no-fills erased the very thing the caller selected the style for
				// (#23). Author `border: { type: 'none' }` to erase a styled table's grid anyway.
				const authoredBorder = newCellOptions.border || opt.border
				newCellOptions.border =
					authoredBorder ||
					(opt.tableStyle ? undefined : [{ type: 'none' }, { type: 'none' }, { type: 'none' }, { type: 'none' }])
				let cellBorder = newCellOptions.border

				// CASE 1: border interface is: BorderOptions | [BorderOptions, BorderOptions, BorderOptions, BorderOptions]
				if (cellBorder && typeof cellBorder === 'object') {
					cellBorder = normalizeBorderTuple(cellBorder)
					newCellOptions.border = cellBorder
				}
				// Handle: [null, null, {type:'solid'}, null]. A null side is *omitted*, not erased:
				// it keeps inheriting from the built-in style/theme, exactly the distinction
				// `normalizeOuterBorder` documents for the perimeter ("a sparse side is NOT
				// `{ type: 'none' }`"). Converting a hole into an explicit `<a:noFill/>` overrides
				// that inheritance and there was no spelling left for "draw this edge, leave that one".
				// The whole tuple is absent when a styled table authored no border at all, which is
				// the same distinction one level up: nothing to complete, nothing to emit.
				const cellBorderTuple = newCellOptions.border as BorderTuple | undefined

				// set complete BorderOptions for the sides that were authored
				if (cellBorderTuple) {
					const arrSides = [0, 1, 2, 3] as const
					arrSides.forEach((idx) => {
						const side = cellBorderTuple[idx]
						if (!side) return
						// `withBorderDefaults` spreads first and overrides only the defaulted keys.
						// Rebuilding the side from a fixed key list dropped `cap` — public on
						// `BorderProps` and already read by `genTableCellBorderXml`, so every table
						// border emitted cap="flat" whatever the caller asked for (and `dashType`
						// would go the same way). The spread also gives each side its own object,
						// which the single-BorderProps form (one object shared across all four)
						// relies on.
						cellBorderTuple[idx] = withBorderDefaults(side)
					})
					newCellOptions.border = cellBorderTuple
				}

				// LAST:
				newRow.push(newCell)
			})
		} else {
			// The same condition STEP 1 rejects for row 0, reaching us on a later row. It used to log
			// and push an empty row, so a deck built fine and quietly lost a row of content.
			throw new InvalidOptionError(
				'table/rows-not-nested',
				`addTable: 'rows' should be an array of cells! Row ${idx} is ${JSON.stringify(row)}`
			)
		}

		arrRows.push(newRow)
	})
	return arrRows
}

/**
 * Register the slide media relationship behind every image fill this table uses, so the
 * cell emitter can resolve `<a:blipFill r:embed="rIdN">` instead of warning and dropping
 * the fill. Tables were the last major object kind not wired into `registerImageFillMedia`
 * (shapes do it at `define/shape.ts`, text boxes at `define/text.ts`).
 *
 * Three fill sources are walked, which between them cover every way a picture reaches a table:
 * the resolved cells (per-cell `options.fill`, plus `headerRow` and `columns[i]`, both baked
 * onto cells by the sugar step above), the table-level `fill`, which is *not* baked —
 * `gen/slide/objects/table.ts` copies it onto each cell at emit time — and `tableFill`, which
 * lands on `a:tblPr` itself and so needs its own relationship for the same reason.
 *
 * Keyed on fill **object identity**, not on the image source. The sugar spreads its options
 * shallowly, so one `headerRow` fill object is shared by every cell it styles; registering
 * per cell would mint a redundant relationship each time and let the later call overwrite
 * the `_imgRid` the first one stashed. Two cells given separately-authored fills for the
 * same file still get one relationship each — matching how two shapes sharing an image
 * behave — while the bytes are deduped into a single media part downstream.
 *
 * Must run only after auto-paging has shredded the rows across slides, for the same reason
 * `createHyperlinkRels` does: `target` has to be the slide the cell actually lands on, or
 * every relationship piles onto the first slide.
 */
function registerTableImageFills(target: PresSlideInternal, rows: TableCell[][], opt: TableProps): void {
	const seen = new Set<ShapeFillProps>()
	const register = (fill: ShapeFillProps | undefined): void => {
		// `type:'image'` and a bare `image:{…}` are both accepted, mirroring the shape path.
		if (!fill || typeof fill !== 'object') return
		if (fill.type !== 'image' && !fill.image) return
		if (seen.has(fill)) return
		seen.add(fill)
		registerImageFillMedia(target, fill)
	}

	register(opt.fill)
	register(opt.tableFill)
	rows.forEach((row) => {
		row.forEach((cell) => {
			register(cell.options?.fill)
		})
	})
}

export function addTableDefinition(
	target: PresSlideInternal,
	tableRows: TableRow[],
	options: TableProps,
	slideLayout: SlideLayoutInternal | null,
	presLayout: PresLayout,
	addSlide: (options?: AddSlideProps) => PresSlideInternal,
	getSlide: (slideNumber: number) => PresSlideInternal | undefined
): PresSlideInternal[] {
	const slides: PresSlideInternal[] = [target] // Create array of Slides as more may be added by auto-paging
	// Take ownership of the options before touching them, the same way `addTextDefinition` does.
	// Everything below normalizes in place — `objectName`, `fontSize`, `margin`, `color`, the
	// `autoPage*` family, the resolved `w`/`colW` — and STEP 5 hands this object to every plain
	// string cell as that cell's options, so the cell emitters write onto it too. Without the copy
	// all of that lands on the CALLER's object, and a style literal reused across tables carries one
	// table's settings (and its `objectName`) into the next.
	//
	// Identity WITHIN one call is kept on purpose: the string cells in STEP 5 all share this one
	// object, which is what `gen/slide/objects/table.ts` reads back. Copying per cell would change
	// the emitted bytes.
	//
	// `border` is copied because the array normalization below writes `withBorderDefaults` results
	// back into its slots; `fill` and the cell-level objects stay shared by reference, since rel ids
	// are registered through them and read back at emit time.
	const opt: TableProps = options && typeof options === 'object' ? { ...options } : {}
	if (Array.isArray(opt.border)) opt.border = [...opt.border] as typeof opt.border
	opt.objectName = opt.objectName
		? encodeXmlAttrValue(validateObjectName(opt.objectName, 'table'))
		: `Table ${target._slideObjects.filter((obj) => obj._type === SlideObjectType.table).length}`

	// STEP 0: PLACEHOLDER — a table targeting a layout placeholder inherits that placeholder's
	// position/size for any of x/y/w/h the caller omits, mirroring the image and
	// text placeholder inheritance. Explicit values always win; this only fills the gaps so
	// the table fills the placeholder geometry rather than the default 1in/full-width fallback.
	if (opt.placeholder && slideLayout?._slideObjects) {
		const placeHold = slideLayout._slideObjects.find(
			(item) => item._type === SlideObjectType.placeholder && item.options?.placeholder === opt.placeholder
		)
		if (placeHold?.options) {
			if (opt.x === undefined) opt.x = placeHold.options.x
			if (opt.y === undefined) opt.y = placeHold.options.y
			if (opt.w === undefined) opt.w = placeHold.options.w
			if (opt.h === undefined) opt.h = placeHold.options.h
		}
	}

	// STEP 1: REALITY-CHECK
	{
		// A: check for empty
		if (tableRows === null || tableRows.length === 0 || !Array.isArray(tableRows)) {
			throw new InvalidOptionError(
				'table/rows-not-an-array',
				"addTable: Array expected! EX: 'slide.addTable( [rows], {options} );'"
			)
		}

		// B: check for non-well-formatted array (ex: rows=['a','b'] instead of [['a','b']])
		if (!tableRows[0] || !Array.isArray(tableRows[0])) {
			throw new InvalidOptionError(
				'table/rows-not-nested',
				"addTable: 'rows' should be an array of cells! EX: 'slide.addTable( [ ['A'], ['B'], {text:'C',options:{align:'center'}} ] );'"
			)
		}
	}

	// STEP 1.5: `headerRow` / `columns` inline sugar — bake blanket styling into cells as
	// direct per-cell formatting so it flows through the normal cell pipeline (incl. border
	// defaulting below). Precedence (highest wins), matching how PowerPoint resolves styling
	// (direct formatting overrides a style region): explicit per-cell `options` > `headerRow`
	// (row 0) > `columns[colIdx]` > `tableStyle`/defaults. The merge is property-level, so a
	// header cell keeps `headerRow` typography and takes its column's fill when they differ.
	// Setting `headerRow` implies `hasHeader` unless the caller set it explicitly. The caller's
	// `tableRows` array is not mutated — only affected rows (and their cells) are shallow-copied.
	const srcRows = applyTableHeaderColumnSugar(tableRows, opt)

	// STEP 2: Transform `tableRows` into well-formatted TableCell's
	// tableRows can be object or plain text array: `[{text:'cell 1'}, {text:'cell 2', options:{color:'ff0000'}}]` | `["cell 1", "cell 2"]`
	const arrRows = normalizeTableRows(srcRows, opt)

	// STEP 3: Set options
	// Keep x/y/w/h as raw user `Coord` (inches/percent/unit-string). They are resolved to EMU
	// exactly once at emission (`gen/slide/objects/table.ts`) and by the auto-pager (getSlidesForTableRows); no
	// pre-conversion here, so a value is never parsed twice. Default position is 0.5in.
	if (opt.x === undefined || opt.x === null) opt.x = 0.5
	if (opt.y === undefined || opt.y === null) opt.y = 0.5
	// NOTE: Dont set default `h` - leaving it null triggers auto-rowH in `makeXMLSlide()`
	opt.fontSize = opt.fontSize || DEF_FONT_SIZE
	opt.margin = opt.margin === 0 || opt.margin ? opt.margin : DEF_CELL_MARGIN_IN
	if (typeof opt.margin === 'number')
		opt.margin = [Number(opt.margin), Number(opt.margin), Number(opt.margin), Number(opt.margin)]
	// defensive fallback - if `opt.margin` is not a 4-element array of finite numbers, use defaults so non-numeric table-level margins don't leak NaN into <a:tcPr>
	if (
		!Array.isArray(opt.margin) ||
		opt.margin.length !== 4 ||
		opt.margin.some((v: unknown) => typeof v !== 'number' || !Number.isFinite(v))
	) {
		opt.margin = DEF_CELL_MARGIN_IN
	}
	// Black lands on the table and is inherited by every cell that has no colour of its own
	// (`gen/slide/objects/table.ts`), as direct formatting. One thing stands it down: a
	// hyperlink anywhere in the grid, because the default paints the whole run, so the words
	// *after* a link come out black instead of following the link colour. That run is then
	// emitted with no `<a:solidFill>` and PowerPoint resolves the text from the theme's `tx1`.
	if (!JSON.stringify({ arrRows }).includes('hyperlink')) {
		if (!opt.color) opt.color = DEF_FONT_COLOR // table option > inherit from Slide > default to black
	}
	if (typeof opt.border === 'string') {
		warn('table/invalid-border', "addTable `border` option must be an object. Ex: `{border: {type:'none'}}`")
		opt.border = undefined
	} else if (Array.isArray(opt.border)) {
		const border = opt.border
		;([0, 1, 2, 3] as const).forEach((idx) => {
			const side = border[idx]
			border[idx] = side ? withBorderDefaults(side) : { type: 'none' }
		})
	}
	if (typeof opt.outerBorder === 'string') {
		warn(
			'table/invalid-outer-border',
			"addTable `outerBorder` option must be an object. Ex: `{outerBorder: {type:'solid'}}`"
		)
		opt.outerBorder = undefined
	}
	// The perimeter is resolved here but applied at emit time: which cells sit on the table's
	// outer edge is only knowable once the merge grid exists, and the serializer is what builds
	// it — a colspan reaching the last column puts that column's rule on a *covered* cell.
	opt.outerBorder = normalizeOuterBorder(opt.outerBorder)

	opt.autoPage = typeof opt.autoPage === 'boolean' ? opt.autoPage : false
	opt.autoPagePlaceholder = typeof opt.autoPagePlaceholder === 'boolean' ? opt.autoPagePlaceholder : false
	opt.autoPageRepeatHeader = typeof opt.autoPageRepeatHeader === 'boolean' ? opt.autoPageRepeatHeader : false
	opt.autoPageHeaderRows =
		typeof opt.autoPageHeaderRows !== 'undefined' && !isNaN(Number(opt.autoPageHeaderRows))
			? Number(opt.autoPageHeaderRows)
			: 1
	opt.autoPageLineWeight =
		typeof opt.autoPageLineWeight !== 'undefined' && !isNaN(Number(opt.autoPageLineWeight))
			? Number(opt.autoPageLineWeight)
			: 0
	if (opt.autoPageLineWeight) {
		if (opt.autoPageLineWeight > 1) opt.autoPageLineWeight = 1
		else if (opt.autoPageLineWeight < -1) opt.autoPageLineWeight = -1
	}
	// autoPage ^^^

	// Set/Calc table width
	// Get slide margins - start with default values, then adjust if master or slide margins exist
	let arrTableMargin = DEF_SLIDE_MARGIN_IN
	// Master margins override the defaults, if present
	if (slideLayout && typeof slideLayout._margin !== 'undefined') {
		if (Array.isArray(slideLayout._margin)) arrTableMargin = slideLayout._margin
		else if (!isNaN(Number(slideLayout._margin))) {
			arrTableMargin = [
				Number(slideLayout._margin),
				Number(slideLayout._margin),
				Number(slideLayout._margin),
				Number(slideLayout._margin),
			]
		}
	}

	/**
	 * Calc table width depending upon what data we have - several scenarios exist (including bad data, eg: colW doesnt match col count)
	 * The API does not require a `w` value, but XML generation does, hence, code to calc a width below using colW value(s)
	 */
	if (opt.colW) {
		const firstRowColCnt = (arrRows[0] ?? []).reduce((totalLen, c) => {
			if (c?.options?.colspan && typeof c.options.colspan === 'number') {
				totalLen += c.options.colspan
			} else {
				totalLen += 1
			}
			return totalLen
		}, 0)

		if (typeof opt.colW === 'string' || typeof opt.colW === 'number') {
			// Ex: `colW = 3` or `colW = '3'`
			opt.w = Math.floor(Number(opt.colW) * firstRowColCnt)
			opt.colW = undefined // IMPORTANT: Unset `colW` so table is created using `opt.w`, which will evenly divide cols
		} else if (opt.colW && Array.isArray(opt.colW) && opt.colW.length === 1 && firstRowColCnt > 1) {
			// Ex: `colW=[3]` but with >1 cols (same as above, user is saying "use this width for all")
			opt.w = Math.floor(Number(opt.colW) * firstRowColCnt)
			opt.colW = undefined // IMPORTANT: Unset `colW` so table is created using `opt.w`, which will evenly divide cols
		} else if (opt.colW && Array.isArray(opt.colW) && opt.colW.length !== firstRowColCnt) {
			// Err: Mismatched colW and cols count
			warn(
				'table/col-width-count-mismatch',
				'addTable: mismatch: (colW.length != data.length) Therefore, defaulting to evenly distributed col widths.'
			)
			opt.colW = undefined
		}
	} else if (opt.w) {
		// Keep raw user `Coord` — resolved to EMU once at emission. (No pre-conversion.)
	} else {
		opt.w = Math.floor((presLayout._sizeW || presLayout.width) / EMU_PER_INCH - arrTableMargin[1] - arrTableMargin[3])
	}

	// Shrink-to-fit (`fitColumns: 'shrink'`): proportionally scale columns down so a
	// too-wide table fits between `x` and the right slide margin. Runs after the width-calc
	// above so it sees the resolved form — either a surviving per-column `colW` array, or a
	// single `w`. Rewriting the widths here (the table definition) means both the emitter and
	// the measured-fit pass inherit the fitted grid; shrink only, no minimum-width floor.
	if (opt.fitColumns === 'shrink') {
		const slideWin = (presLayout._sizeW || presLayout.width) / EMU_PER_INCH
		const xIn = getSmartParseNumber(opt.x, 'X', presLayout) / EMU_PER_INCH
		// `arrTableMargin` is TRBL, so the margin to the RIGHT of the table is index 1. This
		// read index 3 (the left margin) — invisible with the default symmetric [0.5 × 4], and
		// wrong by the difference for any master whose `_margin` is asymmetric, which is
		// exactly the layout a caller sets a left gutter on.
		const availWin = slideWin - xIn - arrTableMargin[1]
		if (availWin > 0) {
			if (Array.isArray(opt.colW)) {
				const sumIn = opt.colW.reduce((p, n) => p + (Number.isFinite(n) ? n : 0), 0)
				if (sumIn > availWin) {
					const factor = availWin / sumIn
					opt.colW = opt.colW.map((n) => (Number.isFinite(n) ? n * factor : n))
					opt.w = availWin
				}
			} else if (typeof opt.w === 'number' && opt.w > availWin) {
				opt.w = availWin
			}
		}
	}

	// STEP 5: Loop over cells: transform each to TableCell; check to see whether to unset `autoPage` while here
	arrRows.forEach((row) => {
		row.forEach((cell, idy) => {
			// A: Transform cell data if needed
			/* Table rows can be an object or plain text - transform into object when needed
				// EX:
				const arrTabRows1 = [
					[ { text:'A1\nA2', options:{rowspan:2, fill:'99FFCC'} } ]
					,[ 'B2', 'C2', 'D2', 'E2' ]
				]
			*/
			if (typeof cell === 'number' || typeof cell === 'string') {
				// Grab table formatting `opts` to use here so text style/format inherits as it should
				row[idy] = { _type: SlideObjectType.tablecell, text: String(row[idy]), options: opt }
			} else if (typeof cell === 'object') {
				const target = row[idy]
				if (!target) return
				// ARG0: `text` (numeric input was already coerced to a string in the first pass)
				if (typeof cell.text === 'undefined' || cell.text === null) target.text = ''

				// ARG1: `options`: ensure options exists
				target.options = cell.options || {}

				// Set type to tabelcell
				target._type = SlideObjectType.tablecell
			}

			// B: Check for fine-grained formatting, disable auto-page when found
			// Since genXmlTextBody already checks for text array ( text:[{},..{}] ) we're done!
			// Text in individual cells will be formatted as they are added by calls to genXmlTextBody within table builder
		})
	})

	// If autoPage = true, we need to return references to newly created slides if any
	const newAutoPagedSlides: PresSlideInternal[] = []

	// STEP 6: Auto-Paging: (via {options} and used internally)
	// (used internally by `tableToSlides()` to not engage recursion - we've already paged the table data, just add this one)
	if (opt && !opt.autoPage) {
		// Create hyperlink rels (IMPORTANT: Wait until table has been shredded across Slides or all rels will end-up on Slide 1!)
		createHyperlinkRels(target, arrRows)

		// Same timing rule as the hyperlink rels above: resolve cell image fills to media rels
		// on the slide this table actually landed on.
		registerTableImageFills(target, arrRows, opt)

		// Add slideObjects (NOTE: Use `extend` to avoid mutation)
		// `columns` (per-column TableCellProps[]) is consumed in STEP 1.5 and baked into cells;
		// drop it here so it is not carried into `ObjectOptions`, where `columns` is the unrelated
		// text-column *count* (`number`). Leaving it in would be both meaningless and a type clash.
		target._slideObjects.push({
			_type: SlideObjectType.table,
			arrTabRows: arrRows,
			options: { ...opt, columns: undefined },
		})
	} else {
		if (opt.autoPageRepeatHeader)
			opt._arrObjTabHeadRows = arrRows.filter((_row, idx) => idx < (opt.autoPageHeaderRows || 1))

		// snapshot populated placeholders on the source slide (e.g. a title added via
		// `addText(text, { placeholder })`) so they can be re-rendered on each overflow slide.
		// Overflow slides otherwise inherit only the layout's empty placeholders. Captured before
		// the loop so the table object added per-slide below is never included.
		const sourcePlaceholders =
			opt.autoPagePlaceholder && Array.isArray(target._slideObjects)
				? target._slideObjects.filter((obj) => obj._type !== SlideObjectType.table && obj.options?.placeholder)
				: []

		// Loop over rows and create 1-N tables as needed
		getSlidesForTableRows(arrRows, opt, presLayout, slideLayout).forEach((slide, idx) => {
			// A: Create new Slide when needed, otherwise, use existing (NOTE: More than 1 table can be on a Slide, so we will go up AND down the Slide chain)
			let newSlide = getSlide(target._slideNum + idx)
			if (!newSlide) {
				newSlide = addSlide({ masterTitle: slideLayout?._name || undefined })
				slides.push(newSlide)
			}

			// B: Reset opt.y to `option`/`margin` after first Slide
			// Keep raw inches — resolved to EMU once at emission. (No pre-conversion.)
			if (idx > 0) opt.y = opt.autoPageSlideStartY || arrTableMargin[0]

			// C: Add this table to new Slide
			{
				opt.autoPage = false

				// copy the source slide's populated placeholders onto each overflow slide
				// (idx 0 is the source slide itself and already has them).
				if (idx > 0 && sourcePlaceholders.length > 0) {
					sourcePlaceholders.forEach((ph) => newSlide._slideObjects.push(structuredClone(ph)))
				}

				// Create hyperlink rels (IMPORTANT: Wait until table has been shredded across Slides or all rels will end-up on Slide 1!)
				createHyperlinkRels(newSlide, slide.rows)

				// Add rows to new slide. When `rowH` is an array it is keyed by *original* row index,
				// which no longer matches the per-slide physical row order after pagination; use the
				// per-slide heights the auto-pager resolved so each row keeps its configured height
				// instead of inheriting whatever row lands at the same index.
				// `slide.rowH` may contain `undefined` holes (auto-height rows); the table serializer
				// treats a falsy per-row height as "auto", so the cast to number[] is safe.
				newSlide.addTable(slide.rows, {
					...opt,
					rowH: Array.isArray(opt.rowH) && slide.rowH ? (slide.rowH as number[]) : opt.rowH,
				})

				// Add reference to the new slide so it can be returned, but don't add the first one because the user already has a reference to that one.
				if (idx > 0) newAutoPagedSlides.push(newSlide)
			}
		})
	}
	return newAutoPagedSlides
}
