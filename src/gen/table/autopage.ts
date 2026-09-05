/**
 * ts-pptx: Table auto-paging core
 *
 * The DOM-independent heart of table generation: given rows that overflow a slide,
 * split them across as many slides as needed, measuring wrapped text to compute line
 * counts and row heights. `getSlidesForTableRows` is the in-memory core (fully
 * Node-testable); `parseTextToLines` is its private cell-wrapping helper.
 */

import { SlideObjectType } from '../../enums.js'
import { DEF_FONT_SIZE, DEF_SLIDE_MARGIN_IN } from '../../constants-internal.js'
import type { PresLayout, TableRowSlide, TableCellProps } from '../../types/index.js'
import type {
	SlideLayoutInternal,
	TableCellInternal,
	TableRowInternal,
	TableToSlidesPropsInternal,
} from '../../types/internal.js'
import {
	autoPageLineHeightEmu,
	getSmartParseNumber,
	inch2Emu,
	marginToEmu,
	pinnedRowHeightInches,
	resolveCellMarginsInches,
	resolveSlideMarginsInches,
	resolveTableColWidthsEmu,
	usableTableWidthEmu,
} from '../../units-internal.js'
import { warn } from '../../diagnostics.js'
import { type GridPlacement, tableColCount, walkTableGrid } from './grid.js'
import { resolveSpan, withCheckedSpans } from './spans.js'
import { EMU_PER_INCH, POINTS_PER_INCH } from '../../units.js'

type AutoPageCell = TableCellInternal & {
	_lineHeight: number
	_lines: TableCellInternal[][]
	options: TableCellProps
	text: TableCellInternal[]
}

/**
 * A cell for the pager's working grid, carrying the source cell's options only when it has some.
 *
 * Every one of these used to be an inline `{ …, options: cell.options }`, which put an `options`
 * key holding `undefined` on a cell whose source had none. Readers cannot see the difference —
 * they all go through `cell.options?.…` — but a `TableCellProps` bag *is* spread in
 * `gen/define/table.ts` (`{ ...colDef, ...headerRow, ...cell.options }`), where the two states
 * are not the same thing. So the model keeps one spelling of absent.
 * @param text - the cell's content, a string or the run list the pager is accumulating
 * @param options - the source cell's options, if it had any
 */
function workingCell(text: string | TableCellInternal[], options: TableCellProps | undefined): TableCellInternal {
	return options === undefined
		? { _type: SlideObjectType.tablecell, text }
		: { _type: SlideObjectType.tablecell, text, options }
}

// ===== Cell text wrapping =====

/**
 * Break cell text into lines based upon table column width (e.g.: Magic Happens Here(tm))
 *
 * `charWeight` arrives as an argument rather than being read off `cell.options` because the
 * caller used to *stamp* it there: the table's weight was written onto the caller's own cell
 * options bag, or `delete`d from it when the table stated none. That mutated caller input --
 * which the chart path has a dedicated guard against -- and it made the option's precedence
 * differ by call site: the row-height probe read the cell's own value while the main loop had
 * just overwritten it with the table's. Both callers now resolve it the same way, and the more
 * specific statement wins, as it does everywhere else in this codebase.
 * @param {TableCellInternal} cell - table cell
 * @param {number} colWidth - table column width (inches)
 * @param {number} charWeight - resolved `autoPageCharWeight`: the cell's, else the table's, else 0
 * @param {boolean} [verbose] - dump the four wrapping stages; carries `addTable({ verbose })`
 *   down, which is the only stage of the pager that flag did not reach
 * @return {TableRowInternal[]} - cell's text objects grouped into lines
 */
function parseTextToLines(
	cell: TableCellInternal,
	colWidth: number,
	charWeight: number,
	verbose?: boolean
): TableCellInternal[][] {
	// FYI: CPL = Width / (font-size / font-constant)
	// FYI: CHAR:2.3, colWidth:10, fontSize:12 => CPL=138, (actual chars per line in PPT)=145 [14.5 CPI]
	// FYI: CHAR:2.3, colWidth:7 , fontSize:12 => CPL= 97, (actual chars per line in PPT)=100 [14.3 CPI]
	// FYI: CHAR:2.3, colWidth:9 , fontSize:16 => CPL= 96, (actual chars per line in PPT)=84  [ 9.3 CPI]
	const FOCO = 2.3 + charWeight // Character Constant
	// `colWidth` is inches, so the column's width in points is `colWidth * 72`. This used to be
	// spelled `(colWidth / EMU_PER_POINT) * EMU_PER_INCH`, which is the same 72 with two EMU
	// constants that cancel -- except that the detour through EMU is not exact in binary floating
	// point, and `Math.floor` takes the hit. Over the 200,000 widths from 0.001in to 200in the two
	// forms disagree on 51, always by one point low: 6.625in is exactly 477 points, and the old
	// form floored 476.9999... to 476.
	const CPL =
		Math.floor(colWidth * POINTS_PER_INCH) / ((cell.options?.fontSize ? cell.options.fontSize : DEF_FONT_SIZE) / FOCO) // Chars-Per-Line

	// The number the four dumps below exist to explain: every wrap decision in step 4 is
	// `strCurrLine.length + word.length > CPL`, so a wrap that looks wrong is almost always
	// a CPL that is wrong, and CPL is not recoverable from the lines themselves.
	if (verbose)
		console.log(
			`[0/4] colWidth=${colWidth}in fontSize=${cell.options?.fontSize ?? DEF_FONT_SIZE} FOCO=${FOCO} CPL=${CPL}`
		)

	const parsedLines: TableCellInternal[][] = []
	let inputCells: TableCellInternal[] = []
	const inputLines1: TableCellInternal[][] = []
	const inputLines2: TableCellInternal[][] = []

	/**
	 * EX INPUTS: `cell.text`
	 * - string....: "Account Name Column"
	 * - object....: { text:"Account Name Column" }
	 * - object[]..: [{ text:"Account Name", options:{ bold:true } }, { text:" Column" }]
	 * - object[]..: [{ text:"Account Name", options:{ breakLine:true } }, { text:"Input" }]
	 */

	/**
	 * EX OUTPUTS:
	 * - string....: [{ text:"Account Name Column" }]
	 * - object....: [{ text:"Account Name Column" }]
	 * - object[]..: [{ text:"Account Name", options:{ breakLine:true } }, { text:"Input" }]
	 * - object[]..: [{ text:"Account Name", options:{ breakLine:true } }, { text:"Input" }]
	 */

	// STEP 1: Ensure inputCells is an array of TableCells
	if (cell.text && cell.text.toString().trim().length === 0) {
		// Allow a single space/whitespace as cell text (user-requested feature)
		inputCells.push({ _type: SlideObjectType.tablecell, text: ' ' })
	} else if (typeof cell.text === 'number' || typeof cell.text === 'string') {
		inputCells.push({ _type: SlideObjectType.tablecell, text: (cell.text || '').toString().trim() })
	} else if (Array.isArray(cell.text)) {
		inputCells = cell.text
	}
	if (verbose) {
		console.log('[1/4] inputCells')
		inputCells.forEach((cell, idx) => console.log(`[1/4] [${idx + 1}] cell: ${JSON.stringify(cell)}`))
	}

	// STEP 2: Group table cells into lines based on "\n" or `breakLine` prop
	/**
	 * - EX: `[{ text:"Input Output" }, { text:"Extra" }]`                       == 1 line
	 * - EX: `[{ text:"Input" }, { text:"Output", options:{ breakLine:true } }]` == 1 line
	 * - EX: `[{ text:"Input\nOutput" }]`                                        == 2 lines
	 * - EX: `[{ text:"Input", options:{ breakLine:true } }, { text:"Output" }]` == 2 lines
	 */
	let newLine: TableCellInternal[] = []
	inputCells.forEach((cell) => {
		if (typeof cell.text !== 'string') return

		if (cell.text.includes('\n')) {
			// Each non-last \n-split part ends with a forced paragraph break; the last
			// part accumulates in newLine so subsequent runs stay on the same paragraph.
			const parts = cell.text.split('\n')
			parts.forEach((part, partIdx) => {
				const isLastPart = partIdx === parts.length - 1
				if (isLastPart) {
					newLine.push(workingCell(part, cell.options))
				} else {
					newLine.push({
						_type: SlideObjectType.tablecell,
						text: part,
						options: { ...cell.options, breakLine: true },
					})
					inputLines1.push(newLine)
					newLine = []
				}
			})
		} else {
			newLine.push(workingCell(cell.text.trim(), cell.options))
		}

		if (cell.options?.breakLine) {
			if (verbose) console.log(`inputCells: new line > ${JSON.stringify(newLine)}`)
			inputLines1.push(newLine)
			newLine = []
		}
	})
	// Flush remaining buffer after all cells are processed
	if (newLine.length > 0) {
		inputLines1.push(newLine)
		newLine = []
	}
	if (verbose) {
		console.log(`[2/4] inputLines1 (${inputLines1.length})`)
		inputLines1.forEach((line, idx) => console.log(`[2/4] [${idx + 1}] line: ${JSON.stringify(line)}`))
	}

	// STEP 3: Tokenize every text object into words. All runs of one logical paragraph
	// are flattened into a single token list so step 4 tracks column position across
	// styled-run boundaries (fixes independent-reset bug for rich-text cells).
	inputLines1.forEach((line) => {
		const lineTokens: TableCellInternal[] = []
		line.forEach((cell) => {
			const cellTextStr = String(cell.text) // force convert to string (compiled JS is better with this than a cast)
			const lineWords = cellTextStr.split(' ')

			lineWords.forEach((word, idx) => {
				const cellProps = { ...cell.options }
				// IMPORTANT: Handle `breakLine` prop - we cannot apply to each word - only apply to very last word!
				if (cellProps?.breakLine) cellProps.breakLine = idx + 1 === lineWords.length
				lineTokens.push({
					_type: SlideObjectType.tablecell,
					text: word + (idx + 1 < lineWords.length ? ' ' : ''),
					options: cellProps,
				})
			})
		})
		inputLines2.push(lineTokens)
	})
	if (verbose) {
		console.log(`[3/4] inputLines2 (${inputLines2.length})`)
		inputLines2.forEach((line) => console.log(`[3/4] line: ${JSON.stringify(line)}`))
	}

	// STEP 4: Group cells/words into lines based upon space consumed by word letters
	inputLines2.forEach((line) => {
		let lineCells: TableCellInternal[] = []
		let strCurrLine = ''

		line.forEach((word) => {
			const wordText = String(word.text || '')
			// A: create new line when horizontal space is exhausted.
			// `lineCells.length > 0` is the same guard the flush below carries, and for the same
			// reason: on the first word `strCurrLine` is `''`, so a word wider than the column on
			// its own fires this branch with nothing buffered and used to push an EMPTY line. The
			// cell's text survived -- the pager concatenates lines back together -- but its height
			// did not, because `_lines.length` is what prices the row. Twelve single-word rows in a
			// 0.4in column paged onto two slides when the word was long and one when it was short.
			// The word still overflows its column: nothing here breaks inside a word.
			if (strCurrLine.length + wordText.length > CPL) {
				if (lineCells.length > 0) parsedLines.push(lineCells)
				lineCells = []
				strCurrLine = ''
			}

			// B: add current word to line cells
			lineCells.push(word)

			// C: add current word to `strCurrLine` which we use to keep track of line's char length
			strCurrLine += wordText
		})

		// Flush buffer: Only create a line when there's text to avoid empty row
		if (lineCells.length > 0) parsedLines.push(lineCells)
	})
	if (verbose) {
		console.log(`[4/4] parsedLines (${parsedLines.length})`)
		parsedLines.forEach((line, idx) => console.log(`[4/4] [Line ${idx + 1}]:\n${JSON.stringify(line)}`))
		console.log('...............................................\n\n')
	}

	// Done:
	return parsedLines
}

/**
 * The font size in points one auto-paged cell is measured at: the cell's own `fontSize`,
 * else the table's, else the library default. Pass `null` for the cell to ask the same
 * question of the table alone.
 */
function resolveCellFontSize(
	cellOpts: TableCellProps | undefined | null,
	tableOpts: TableToSlidesPropsInternal
): number {
	const size = cellOpts?.fontSize ?? tableOpts.fontSize
	return typeof size === 'number' ? size : DEF_FONT_SIZE
}

/**
 * The `autoPageCharWeight` a cell wraps at: its own, else the table's, else none.
 *
 * The sibling of {@link resolveCellFontSize}, and named for the same reason -- the two call
 * sites of {@link parseTextToLines} had resolved this one differently. The row-height probe
 * read the cell's own value; the main loop stamped the table's over it first, and `delete`d the
 * cell's when the table stated none, so a weight set on a cell alone was accepted and dropped.
 * The precedence here is the one every other paired option in this codebase uses.
 * @param cellOpts - the cell's own options, if any
 * @param tableOpts - the table's options
 */
function resolveCellCharWeight(
	cellOpts: TableCellProps | undefined | null,
	tableOpts: TableToSlidesPropsInternal
): number {
	const weight = cellOpts?.autoPageCharWeight ?? tableOpts.autoPageCharWeight
	return typeof weight === 'number' ? weight : 0
}

// ===== Auto-page engine =====
// The in-memory core (fully Node-testable). Internally signposted by the
// `// STEP 1..7` markers below: margins → column count → widths → the main
// row-iteration/overflow loop (STEP 6) → final-slide flush (STEP 7).

/**
 * The largest top and bottom cell margin in one row, in EMU.
 *
 * A cell's own `margin` wins over the table's when it states one, and a stated `0` counts as
 * stating one. Both sites that read this gated on truthiness and on `Array.isArray`, so a cell
 * asking for `margin: [0, …]` fell through to the table's margin instead of taking its own, and
 * a scalar `margin: 0.2` was not seen at all -- while the emitter broadcasts a scalar to four
 * sides. `resolveCellMarginsInches` is the resolver the emitter itself reads.
 * @param row - the row's cells
 * @param tableProps - the table's options, for the fallback margin
 * @returns the row's top and bottom margin allowance in EMU
 */
function rowMarginsEmu(
	row: TableRowInternal,
	tableProps: TableToSlidesPropsInternal
): { topEmu: number; btmEmu: number } {
	let topEmu = 0
	let btmEmu = 0
	row.forEach((cell) => {
		const stated = cell?.options?.margin ?? tableProps.margin
		if (stated === undefined || stated === null) return
		const [top, , btm] = resolveCellMarginsInches(stated)
		if (marginToEmu(top) > topEmu) topEmu = marginToEmu(top)
		if (marginToEmu(btm) > btmEmu) btmEmu = marginToEmu(btm)
	})
	return { topEmu, btmEmu }
}

/**
 * The vertical space one repeated header row occupies, in EMU: its own top/bottom cell margins
 * plus its tallest cell's wrapped line count times that row's line height.
 *
 * This is the same arithmetic the main loop applies to a body row, and the reason it has to be
 * spelled again is that the pager used to read `cell._lineHeight` off `_arrObjTabHeadRows` --
 * which holds the DEFINER's plain `TableCellInternal`s. `_lineHeight` is written only onto the pager's
 * own working cells, so it was always absent there and every repeated header row was priced at
 * zero: each continuation page took the header for free and then packed the same number of body
 * rows the first page fits, so the last row hung off the bottom of the slide. That is the same
 * failure the margin reset a few lines below documents having already been fixed once.
 *
 * @param row - one header row, as the definer built it
 * @param colWidthsIn - the resolved column grid, in inches
 * @param numCols - the grid's column count
 * @param tableProps - the table's options, for font size and the two weights
 * @returns the row's height in EMU
 */
function headerRowHeightEmu(
	row: TableRowInternal,
	colWidthsIn: number[],
	numCols: number,
	tableProps: TableToSlidesPropsInternal
): number {
	let maxLines = 0
	let maxLineHeightEmu = 0
	const { topEmu: marTopEmu, btmEmu: marBtmEmu } = rowMarginsEmu(row, tableProps)
	let colCursor = 0
	row.forEach((cell) => {
		const cellOpts = cell.options
		const cellColspan = resolveSpan(cellOpts?.colspan, 'colspan')
		const colStart = colCursor
		colCursor = Math.min(colCursor + cellColspan, numCols)
		const totalColW = colWidthsIn.slice(colStart, colStart + cellColspan).reduce((prev, curr) => prev + curr, 0)

		const lines = parseTextToLines(cell, totalColW, resolveCellCharWeight(cellOpts, tableProps), false).length
		if (lines > maxLines) maxLines = lines
		const lineHeightEmu = autoPageLineHeightEmu(
			resolveCellFontSize(cellOpts, tableProps),
			tableProps.autoPageLineWeight || 0
		)
		if (lineHeightEmu > maxLineHeightEmu) maxLineHeightEmu = lineHeightEmu
	})
	return marTopEmu + marBtmEmu + maxLines * maxLineHeightEmu
}

/**
 * Takes an array of table rows and breaks into an array of slides, which contain the calculated amount of table rows that fit on that slide
 * @param {TableCellInternal[][]} tableRows - table rows
 * @param {TableToSlidesPropsInternal} tableProps - table2slides properties
 * @param {PresLayout} presLayout - presentation layout
 * @param {SlideLayoutInternal} masterSlide - master slide
 * @return {TableRowSlide[]} array of table rows
 */
export function getSlidesForTableRows(
	rows: TableCellInternal[][] = [],
	tableProps: TableToSlidesPropsInternal = {},
	presLayout: PresLayout,
	masterSlide?: SlideLayoutInternal | null
): TableRowSlide[] {
	// `MAX_TABLE_SPAN` guards the two allocations a span decides, and its own module says it
	// covers both paths -- but `tableToSlides` reaches this function directly, without the
	// definer that applies the check, and its own `gridSpan` has no ceiling. So a
	// `colspan="4000000000"` in a source table reached the column walk and the per-column depth
	// array, which is the allocation that takes the host process down with no exception to catch.
	// Checking here rather than at the call site keeps the one entry point guarded whoever calls
	// it; a grid whose spans are all fine is passed through by identity.
	const tableRows = withCheckedSpans(rows)
	let arrInchMargins = DEF_SLIDE_MARGIN_IN
	let emuSlideTabW: number
	let colWidthsIn: number[] = []
	let emuSlideTabH = EMU_PER_INCH * 1
	let emuTabCurrH = 0
	let numCols = 0
	let warnedNoTabH = false
	const tableRowSlides: TableRowSlide[] = []
	const tablePropX = getSmartParseNumber(tableProps.x, 'X', presLayout)
	const tablePropY = getSmartParseNumber(tableProps.y, 'Y', presLayout)
	const tablePropW = getSmartParseNumber(tableProps.w, 'X', presLayout)
	const tablePropH = getSmartParseNumber(tableProps.h, 'Y', presLayout)
	let tableCalcW: number = tablePropW

	/**
	 * Where this page's table starts, in EMU from the top of the slide.
	 *
	 * There used to be three rules for two cases. The first page and "every page after it"
	 * were handled together, then a block gated on `tableRowSlides.length > 1` overrode the
	 * result -- so despite its own comment it began on the THIRD page, and its first arm
	 * recomputed exactly what had just been computed. Page 2 therefore got a different usable
	 * height from pages 3 and up: with `y` above the top margin and an explicit `h`, a
	 * 60-row table paged 9 / 7 / 9 / 9 / 9 / 9 / 8, which is the same "pages disagree about
	 * how many identical rows fit" symptom the margin-reset note in `calcSlideTabH` records.
	 *
	 * It is a function rather than two expressions because the too-small-height fallback in
	 * `calcSlideTabH` had spelled the continuation arm its own way and drifted the same way
	 * twice over: `autoPageSlideStartY || margin` discarded an explicit `0` that the main rule
	 * honoured, and the `Math.min` clause below was missing outright. Two readings of one rule
	 * is what the paragraph above is about; there is now one.
	 */
	function startYEmu(): number {
		if (tableRowSlides.length === 0) return tablePropY || inch2Emu(arrInchMargins[0])
		if (typeof tableProps.autoPageSlideStartY === 'number') return inch2Emu(tableProps.autoPageSlideStartY)
		// RULE: after the first page the table starts at the top margin rather than at its
		// own `y` -- unless `y` is ABOVE the margin, in which case paging must not push the
		// table down and lose the space. Whichever is higher on the slide wins.
		return inch2Emu(tablePropY ? Math.min(tablePropY / EMU_PER_INCH, arrInchMargins[0]) : arrInchMargins[0])
	}

	function calcSlideTabH(): void {
		const emuStartY = startYEmu()
		emuSlideTabH = (tablePropH || presLayout.height) - emuStartY - inch2Emu(arrInchMargins[2])
		// EXPLICIT-H FIX: an explicit `h` is the table's height (an extent), not a bottom
		// coordinate, so -- unlike `presLayout.height` -- the start-Y must not be subtracted from
		// it. Otherwise a table that begins mid-slide gets a tiny page while the pages that do
		// clamp to `h` render normally. Applied on every page: the floor used to reach the first
		// page and pages three and up, leaving page two as the one that could shrink below `h`.
		if (tablePropH && emuSlideTabH < tablePropH) emuSlideTabH = tablePropH

		// GUARD: a non-positive height, or one too small to fit even a single line of the base font,
		// means no row ever fits. That previously emitted degenerate empty overflow pages (rows:[]),
		// which made the recursive addTable throw "Array expected", or — with `h` smaller than one
		// line — placed one row per slide forever. Ignore the unusable height, fall back to the full
		// slide area between margins, and warn once rather than emit a broken table.
		const emuBaseLineH = autoPageLineHeightEmu(
			resolveCellFontSize(null, tableProps),
			tableProps.autoPageLineWeight || 0
		)
		if (emuSlideTabH < emuBaseLineH) {
			const fallbackH = presLayout.height - startYEmu() - inch2Emu(arrInchMargins[2])
			if (!warnedNoTabH) {
				warn(
					'table/autopage-height-too-small',
					'addTable/autoPage: the table height (`h`) leaves no room to paginate; ignoring it and using the slide height. Increase `h` or decrease `y`.'
				)
				warnedNoTabH = true
			}
			emuSlideTabH = fallbackH > 0 ? fallbackH : presLayout.height
		}
	}

	if (tableProps.verbose) {
		console.log('[[VERBOSE MODE]]')
		console.log('|-- TABLE PROPS --------------------------------------------------------|')
		console.log(`| presLayout.width ................................ = ${(presLayout.width / EMU_PER_INCH).toFixed(1)}`)
		console.log(
			`| presLayout.height ............................... = ${(presLayout.height / EMU_PER_INCH).toFixed(1)}`
		)
		console.log(
			`| tableProps.x .................................... = ${typeof tableProps.x === 'number' ? (tableProps.x / EMU_PER_INCH).toFixed(1) : tableProps.x}`
		)
		console.log(
			`| tableProps.y .................................... = ${typeof tableProps.y === 'number' ? (tableProps.y / EMU_PER_INCH).toFixed(1) : tableProps.y}`
		)
		console.log(
			`| tableProps.w .................................... = ${typeof tableProps.w === 'number' ? (tableProps.w / EMU_PER_INCH).toFixed(1) : tableProps.w}`
		)
		console.log(
			`| tableProps.h .................................... = ${typeof tableProps.h === 'number' ? (tableProps.h / EMU_PER_INCH).toFixed(1) : tableProps.h}`
		)
		console.log(
			`| tableProps.slideMargin .......................... = ${tableProps.slideMargin ? String(tableProps.slideMargin) : ''}`
		)
		console.log(`| tableProps.margin ............................... = ${String(tableProps.margin)}`)
		console.log(`| tableProps.colW ................................. = ${String(tableProps.colW)}`)
		console.log(`| tableProps.autoPageSlideStartY .................. = ${tableProps.autoPageSlideStartY}`)
		console.log(`| tableProps.autoPageCharWeight ................... = ${tableProps.autoPageCharWeight}`)
		console.log('|-- CALCULATIONS -------------------------------------------------------|')
		console.log(`| tablePropX ...................................... = ${tablePropX / EMU_PER_INCH}`)
		console.log(`| tablePropY ...................................... = ${tablePropY / EMU_PER_INCH}`)
		console.log(`| tablePropW ...................................... = ${tablePropW / EMU_PER_INCH}`)
		console.log(`| tablePropH ...................................... = ${tablePropH / EMU_PER_INCH}`)
		console.log(`| tableCalcW ...................................... = ${tableCalcW / EMU_PER_INCH}`)
	}

	// STEP 1: Calculate margins
	{
		// The caller's value goes through unaltered, including an absent one: `resolveSlideMarginsInches`
		// owns all three tiers, and defaulting to `DEF_SLIDE_MARGIN_IN[0]` here handed it a stated
		// caller margin that no caller had stated, which suppressed the master's own `margin`.
		arrInchMargins = resolveSlideMarginsInches(masterSlide?._margin, tableProps.slideMargin)

		if (tableProps.verbose)
			console.log(`| arrInchMargins .................................. = [${arrInchMargins.join(', ')}]`)
	}

	// STEP 2: Calculate number of columns
	{
		numCols = tableColCount(tableRows)
		if (tableProps.verbose) console.log(`| numCols ......................................... = ${numCols}`)
	}

	// Where every cell actually lands. The pager asks the grid two questions -- which columns a
	// cell spans, so its text is wrapped against the right width, and whether a rowspan opened
	// on an earlier row is still open here, so a page break does not split a merged group -- and
	// it used to answer both with its own occupancy array, once forwards through the row and
	// once again after it. `walkTableGrid` is the traversal the measured-fit pass and the
	// emitter's own grid build follow, so the widths this pager wraps against are the widths the
	// cells are finally given.
	const placements: GridPlacement[][] = tableRows.map(() => [])
	/**
	 * How many of each row's grid columns are held by a rowspan opened ABOVE it, accumulated as
	 * the placements arrive rather than searched for afterwards.
	 *
	 * The question used to be answered by scanning every placement in the table, once per row,
	 * which is quadratic in a table of any width. Two columns of short text and no spans at all:
	 * 1000 rows took 55ms, 2000 took 115ms, 4000 took 333ms and 8000 took 1036ms. The same four
	 * now take 24ms, 19ms, 40ms and 83ms -- so it was pure overhead, paid in full by the long
	 * report table that is the whole reason this code exists.
	 *
	 * Filling forward instead is bounded by the placement's own `rowSpan`, so the total work is
	 * the sum of all rowspans. That is safe to expand without a further check because
	 * `walkTableGrid` has already clamped each `rowSpan` to `rows.length - r`: a span cannot
	 * reach past the last row however large the caller wrote it, so the sum is bounded by the
	 * grid rather than by anything the caller states.
	 *
	 * The fill starts at `row + 1`. A `rowSpan` of 1 must contribute nothing -- starting at
	 * `row` would have every ordinary row counting itself as covered, and the
	 * `coveredFromAbove(iRow) === numCols` reader below would then fire on all of them.
	 */
	const coveredCols: number[] = Array.from({ length: tableRows.length }, () => 0)
	for (const placement of walkTableGrid(tableRows, numCols)) {
		placements[placement.row]?.push(placement)
		const last = Math.min(placement.row + placement.rowSpan, coveredCols.length)
		for (let r = placement.row + 1; r < last; r++) coveredCols[r] = (coveredCols[r] ?? 0) + placement.colSpan
	}
	/** How many of `iRow`'s grid columns are held by a rowspan opened above it. */
	const coveredFromAbove = (iRow: number): number => coveredCols[iRow] ?? 0
	/** Whether a rowspan opened above `iRow` is still covering it. */
	const spannedFromAbove = (iRow: number): boolean => coveredFromAbove(iRow) > 0

	// STEP 3: Calculate width using tableProps.colW if possible
	if (!tablePropW && tableProps.colW) {
		// The `0` seed is what `colW: []` needs: an empty array is truthy, so it reached an
		// unseeded `reduce` and threw a raw `TypeError` — a failure this library did not raise as
		// a `TsPptxError`. Seeded, it totals `0`, which falls through to `usableTableWidthEmu`
		// exactly where `colW: undefined` lands. That is the reading intended: an empty array is a
		// caller stating no columns, not a table of width zero.
		tableCalcW = Array.isArray(tableProps.colW)
			? tableProps.colW.reduce((p, n) => p + n, 0) * EMU_PER_INCH
			: tableProps.colW * numCols || 0
		if (tableProps.verbose)
			console.log(`| tableCalcW ...................................... = ${tableCalcW / EMU_PER_INCH}`)
	}

	// STEP 4: Calculate usable width now that total usable space is known (`emuSlideTabW`)
	{
		emuSlideTabW = tableCalcW || usableTableWidthEmu(presLayout, tablePropX, arrInchMargins)
		if (tableProps.verbose)
			console.log(`| emuSlideTabW .................................... = ${(emuSlideTabW / EMU_PER_INCH).toFixed(1)}`)
	}

	// STEP 5: Resolve the column grid. `resolveTableColWidthsEmu` is the same resolver the
	// table emitter and the measured-fit pass read, so the widths this pager wraps text
	// against are the widths the emitted `<a:gridCol>`s carry. It also normalizes the array
	// to `numCols`: a short `colW` used to leave the trailing columns measuring against a
	// zero width. Inches, because that is the unit each paged table's `colW` takes; the grid
	// rides out on every `TableRowSlide` rather than being written back onto the caller's bag.
	colWidthsIn = resolveTableColWidthsEmu(tableProps.colW, emuSlideTabW, numCols).map(
		(emu: number) => emu / EMU_PER_INCH
	)

	// Resolve a row's explicit height (inches) from the original `rowH` *array*, keyed by original
	// row index. A single-number `rowH` is left to propagate via table options (it applies uniformly,
	// so it needs no per-row remapping); only the array form is index-sensitive after pagination.
	// Whether a slot pins its row is `pinnedRowHeightInches`, the same rule the writer applies to
	// the array this builds — a `0` accepted here and rejected there is how the two readings drift.
	const resolveRowH = (origRowIdx: number): number | undefined =>
		Array.isArray(tableProps.rowH) ? (pinnedRowHeightInches(tableProps.rowH[origRowIdx]) ?? undefined) : undefined

	// What the repeated header rows cost a continuation page. Computed once: the rows and the
	// grid are the same on every page, so re-measuring per page would only be slower.
	const repeatHeaderRows =
		tableProps.autoPageRepeatHeader && tableProps._arrObjTabHeadRows ? tableProps._arrObjTabHeadRows : []
	const repeatHeaderHeightsEmu = repeatHeaderRows.map((row) =>
		headerRowHeightEmu(row, colWidthsIn, numCols, tableProps)
	)

	// STEP 6: **MAIN** Iterate over rows, add table content, create new slides as rows overflow
	let newTableRowSlide: TableRowSlide = {
		rows: [] as TableRowInternal[],
		rowH: [] as Array<number | undefined>,
		colW: colWidthsIn,
	}
	tableRows.forEach((row, iRow) => {
		// A: Row variables — detect active rowspan at the start of this row so we can
		// suppress page breaks that would split a rowspan group across slides.
		const hasActiveRowSpan = spannedFromAbove(iRow)
		const rowCellLines: AutoPageCell[] = []
		// B: Create new row in data model, calc `maxCellMar*`
		const { topEmu: maxCellMarTopEmu, btmEmu: maxCellMarBtmEmu } = rowMarginsEmu(row, tableProps)
		let currTableRow: TableRowInternal = []
		row.forEach((cell) => currTableRow.push(workingCell([], cell.options)))

		// C: Calc usable vertical space/table height. Set default value first, adjust below when necessary.
		calcSlideTabH()
		emuTabCurrH += maxCellMarTopEmu + maxCellMarBtmEmu // Start row height with margins
		if (tableProps.verbose && iRow === 0)
			console.log(
				`| SLIDE [${tableRowSlides.length}]: emuSlideTabH ...... = ${(emuSlideTabH / EMU_PER_INCH).toFixed(1)} `
			)

		// D: --==[[ BUILD DATA SET ]]==-- (iterate over cells: split text into lines[], set `lineHeight`)
		// Cells are keyed to grid columns, not to their position in the row: a colspan earlier in
		// the row shifts every later cell, and a rowspan opened above skips columns entirely.
		// Measuring a cell against `colW[iCell]` wrapped its text to another column's width.
		const rowPlacements = placements[iRow] ?? []
		row.forEach((cell, iCell) => {
			// A row longer than the grid runs off the end of it: those cells are placed nowhere
			// and measure against no column, which is what the cursor did by clamping.
			const placed = rowPlacements[iCell]
			const colStart = placed ? placed.col : numCols
			const cellColspan = placed ? placed.colSpan : resolveSpan(cell.options?.colspan, 'colspan')

			const newCellOptions = cell.options || {}
			const newCell: AutoPageCell = {
				_type: SlideObjectType.tablecell,
				_lines: [],
				_lineHeight: autoPageLineHeightEmu(
					resolveCellFontSize(cell.options, tableProps),
					tableProps.autoPageLineWeight || 0
				),
				text: [],
				options: newCellOptions,
			}

			// E-1: Exempt cells with `rowspan` from increasing lineHeight (or we could create a new slide when unecessary!)
			if (newCellOptions.rowspan) newCell._lineHeight = 0

			// E-2: **MAIN** Parse cell contents into lines based upon col width, font, etc.
			// A spanning cell is as wide as the columns it covers; the seed keeps a row longer
			// than the grid from throwing on an empty slice.
			const totalColW = colWidthsIn.slice(colStart, colStart + cellColspan).reduce((prev, curr) => prev + curr, 0)

			// E-3: Create lines based upon available column width
			newCell._lines = parseTextToLines(
				cell,
				totalColW,
				resolveCellCharWeight(cell.options, tableProps),
				tableProps.verbose
			)

			// E-4: Add cell to array
			rowCellLines.push(newCell)
		})

		/** E: --==[[ PAGE DATA SET ]]==--
		 * Add text one-line-a-time to this row's cells until: lines are exhausted OR table height limit is hit
		 *
		 * Design:
		 * - Building cells L-to-R/loop style wont work as one could be 100 lines and another 1 line
		 * - Therefore, build the whole row, one-line-at-a-time, across each table columns
		 * - Then, when the vertical size limit is hit is by any of the cells, make a new slide and continue adding any remaining lines
		 *
		 * Implementation:
		 * - `rowCellLines` is an array of cells, one for each column in the table, with each cell containing an array of lines
		 *
		 * Sample Data:
		 * - `rowCellLines` ..: [ TableCellInternal, TableCellInternal, TableCellInternal ]
		 * - `TableCellInternal` .....: { _type: 'tablecell', _lines: TableCellInternal[], _lineHeight: 10 }
		 * - `_lines` ........: [ {_type: 'tablecell', text: 'cell-1,line-1', options: {…}}, {_type: 'tablecell', text: 'cell-1,line-2', options: {…}} }
		 * - `_lines` is TableCellInternal[] (the 1-N words in the line)
		 * {
		 *    _lines: [{ text:'cell-1,line-1' }, { text:'cell-1,line-2' }],                                                     // TOTAL-CELL-HEIGHT = 2
		 *    _lines: [{ text:'cell-2,line-1' }, { text:'cell-2,line-2' }],                                                     // TOTAL-CELL-HEIGHT = 2
		 *    _lines: [{ text:'cell-3,line-1' }, { text:'cell-3,line-2' }, { text:'cell-3,line-3' }, { text:'cell-3,line-4' }], // TOTAL-CELL-HEIGHT = 4
		 * }
		 *
		 * Example: 2 rows, with the firstrow overflowing onto a new slide
		 * SLIDE 1:
		 *  |--------|--------|--------|--------|
		 *  | line-1 | line-1 | line-1 | line-1 |
		 *  |        |        | line-2 |        |
		 *  |        |        | line-3 |        |
		 *  |--------|--------|--------|--------|
		 *
		 * SLIDE 2:
		 *  |--------|--------|--------|--------|
		 *  |        |        | line-4 |        |
		 *  |--------|--------|--------|--------|
		 *  | line-1 | line-1 | line-1 | line-1 |
		 *  |--------|--------|--------|--------|
		 */
		if (tableProps.verbose) console.log(`\n| SLIDE [${tableRowSlides.length}]: ROW [${iRow}]: START...`)
		let currCellIdx = 0
		let emuLineMaxH = 0
		let isDone = false
		while (!isDone) {
			const srcCell = rowCellLines[currCellIdx]
			if (!srcCell) break
			let tgtCell = currTableRow[currCellIdx] // NOTE: may be redefined below (a new row may be created, thus changing this value)

			// 1: calc emuLineMaxH
			rowCellLines.forEach((cell) => {
				if (cell._lineHeight >= emuLineMaxH) emuLineMaxH = cell._lineHeight
			})

			// 2: create a new slide if there is insufficient room for the current row,
			// but never break inside a rowspan group — keep spanned rows together.
			if (emuTabCurrH + emuLineMaxH > emuSlideTabH && !hasActiveRowSpan) {
				if (tableProps.verbose) {
					console.log('\n|-----------------------------------------------------------------------|')
					// prettier-ignore
					console.log(`|-- NEW SLIDE CREATED (currTabH+currLineH > maxH) => ${(emuTabCurrH / EMU_PER_INCH).toFixed(2)} + ${(srcCell._lineHeight / EMU_PER_INCH).toFixed(2)} > ${emuSlideTabH / EMU_PER_INCH}`)
					console.log('|-----------------------------------------------------------------------|\n\n')
				}

				// A: add current row slide or it will be lost (only if it has rows and text)
				if (
					currTableRow.length > 0 &&
					currTableRow.map((cell) => (Array.isArray(cell.text) ? cell.text.length : 0)).reduce((p, n) => p + n) > 0
				) {
					newTableRowSlide.rows.push(currTableRow)
					newTableRowSlide.rowH?.push(resolveRowH(iRow))
				}

				// B: add current slide to Slides array (never push an empty page: a row that does not
				// fit yet has no content here, and an empty `rows` slide crashes the recursive addTable)
				if (newTableRowSlide.rows.length > 0) tableRowSlides.push(newTableRowSlide)

				// C: reset working/curr slide to hold rows as they're created
				const newRows: TableRowInternal[] = []
				newTableRowSlide = { rows: newRows, rowH: [] as Array<number | undefined>, colW: colWidthsIn }

				// D: reset working/curr row
				currTableRow = []
				row.forEach((cell) => currTableRow.push(workingCell([], cell.options)))

				// E: Calc usable vertical space/table height now as we may still be in the same row and code above ("C: Calc usable vertical space/table height.") calc may now be invalid
				calcSlideTabH()

				// F: reset current table height for this new Slide, then re-charge the row's cell
				// margins onto it — the row is starting over here, so it owes them again.
				//
				// ORDER IS LOAD-BEARING. These two statements used to run the other way round: the
				// margins were added and then wiped by the reset, so the first row of every
				// continuation slide was the only row in the table that paid no margin. That let a
				// continuation slide accept one row more than it had room for, and the extra row
				// hung off the bottom of the slide (upstream gitbrent/PptxGenJS#1200). The symptom
				// is easy to miss because the pager is *self*-consistent per slide: it only shows
				// up as the first slide and the continuation slides disagreeing about how many
				// identical rows fit the identical space — which is what
				// test/regression/table-autopage-continuation-budget.test.js pins.
				emuTabCurrH = maxCellMarTopEmu + maxCellMarBtmEmu
				if (tableProps.verbose)
					console.log(
						`| SLIDE [${tableRowSlides.length}]: emuSlideTabH ...... = ${(emuSlideTabH / EMU_PER_INCH).toFixed(1)} `
					)

				// G: handle repeat headers option /or/ Add new empty row to continue current lines into
				repeatHeaderRows.forEach((row, headIdx) => {
					newTableRowSlide.rows.push([...row])
					// Repeated header rows are the original leading rows, so carry their configured height.
					newTableRowSlide.rowH?.push(resolveRowH(headIdx))
					emuTabCurrH += repeatHeaderHeightsEmu[headIdx] ?? 0
				})

				// WIP: NEW: TEST THIS!!
				tgtCell = currTableRow[currCellIdx]
			}

			// 3: set array of words that comprise this line
			const currLine = srcCell._lines.shift()

			// 4: create new line by adding all words from curr line (or add empty if there are no words to avoid "needs repair" issue triggered when cells have null content)
			if (tgtCell && Array.isArray(tgtCell.text)) {
				if (currLine) tgtCell.text = tgtCell.text.concat(currLine)
				else if (tgtCell.text.length === 0)
					tgtCell.text = tgtCell.text.concat({ _type: SlideObjectType.tablecell, text: '' })
				// IMPORTANT: ^^^ add empty if there are no words to avoid "needs repair" issue triggered when cells have null content
			}

			// 5: increase table height by the curr line height (if we're on the last column)
			if (currCellIdx === rowCellLines.length - 1) emuTabCurrH += emuLineMaxH

			// 6: advance column/cell index (or circle back to first one to continue adding lines)
			currCellIdx = currCellIdx < rowCellLines.length - 1 ? currCellIdx + 1 : 0

			// 7: WIP: done?
			const brent = rowCellLines.map((cell) => cell._lines.length).reduce((prev, next) => prev + next)
			if (brent === 0) isDone = true
		}

		// F: Flush/capture row buffer before it resets at the top of this loop.
		//
		// A row that states no cells is kept only when every one of its grid columns is held by a
		// rowspan from above. That is a real row -- a source row states only the cells it *starts*,
		// so one entirely covered from above states none at all, which is what `<tr></tr>` between
		// two spanned rows means -- and the emitter fills it with `vMerge` continuations. Dropping
		// it moved every later row up one while the emitter went on synthesizing that grid row's
		// continuations, so they landed in the next row's `<a:tr>`: a 2-column table came out with
		// a 4-cell row.
		//
		// An empty row NOT covered that way is a different thing and is still dropped: nothing
		// would fill it, and a `<a:tr>` carrying fewer `<a:tc>` than the grid has columns is the
		// malformation PowerPoint offers to repair.
		if (currTableRow.length > 0 || (row.length === 0 && coveredFromAbove(iRow) === numCols)) {
			newTableRowSlide.rows.push(currTableRow)
			newTableRowSlide.rowH?.push(resolveRowH(iRow))
		}

		if (tableProps.verbose) {
			console.log(
				`- SLIDE [${tableRowSlides.length}]: ROW [${iRow}]: ...COMPLETE ...... emuTabCurrH = ${(emuTabCurrH / EMU_PER_INCH).toFixed(2)} ( emuSlideTabH = ${(
					emuSlideTabH / EMU_PER_INCH
				).toFixed(2)} )`
			)
		}
	})

	// STEP 7: Flush buffer / add final slide (skip an empty trailing buffer; always keep at least
	// one slide so a non-empty table is never reduced to zero pages)
	if (newTableRowSlide.rows.length > 0 || tableRowSlides.length === 0) tableRowSlides.push(newTableRowSlide)

	if (tableProps.verbose) {
		console.log('\n|================================================|')
		console.log(`| FINAL: tableRowSlides.length = ${tableRowSlides.length}`)
		tableRowSlides.forEach((slide) => console.log(slide))
		console.log('|================================================|\n\n')
	}

	// LAST:
	return tableRowSlides
}
