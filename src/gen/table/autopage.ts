/**
 * ts-pptx: Table auto-paging core
 *
 * The DOM-independent heart of table generation: given rows that overflow a slide,
 * split them across as many slides as needed, measuring wrapped text to compute line
 * counts and row heights. `getSlidesForTableRows` is the in-memory core (fully
 * Node-testable); `parseTextToLines` is its private cell-wrapping helper.
 */

import { SlideObjectType } from '../../enums.js'
import { DEF_FONT_SIZE, DEF_SLIDE_MARGIN_IN, LINEH_MODIFIER } from '../../constants-internal.js'
import type {
	PresLayout,
	TableCell,
	TableToSlidesProps,
	TableRow,
	TableRowSlide,
	TableCellProps,
} from '../../types/index.js'
import type { SlideLayoutInternal } from '../../types/internal.js'
import { getSmartParseNumber, inch2Emu, marginToEmu } from '../../units-internal.js'
import { warn } from '../../diagnostics.js'
import { EMU_PER_INCH, EMU_PER_POINT } from '../../units.js'

type AutoPageCell = TableCell & {
	_lineHeight: number
	_lines: TableCell[][]
	options: TableCellProps
	text: TableCell[]
}

// ===== Cell text wrapping =====

/**
 * Break cell text into lines based upon table column width (e.g.: Magic Happens Here(tm))
 * @param {TableCell} cell - table cell
 * @param {number} colWidth - table column width (inches)
 * @return {TableRow[]} - cell's text objects grouped into lines
 */
function parseTextToLines(cell: TableCell, colWidth: number, verbose?: boolean): TableCell[][] {
	// FYI: CPL = Width / (font-size / font-constant)
	// FYI: CHAR:2.3, colWidth:10, fontSize:12 => CPL=138, (actual chars per line in PPT)=145 [14.5 CPI]
	// FYI: CHAR:2.3, colWidth:7 , fontSize:12 => CPL= 97, (actual chars per line in PPT)=100 [14.3 CPI]
	// FYI: CHAR:2.3, colWidth:9 , fontSize:16 => CPL= 96, (actual chars per line in PPT)=84  [ 9.3 CPI]
	const FOCO = 2.3 + (cell.options?.autoPageCharWeight ? cell.options.autoPageCharWeight : 0) // Character Constant
	const CPL =
		Math.floor((colWidth / EMU_PER_POINT) * EMU_PER_INCH) /
		((cell.options?.fontSize ? cell.options.fontSize : DEF_FONT_SIZE) / FOCO) // Chars-Per-Line

	const parsedLines: TableCell[][] = []
	let inputCells: TableCell[] = []
	const inputLines1: TableCell[][] = []
	const inputLines2: TableCell[][] = []
	/*
		if (cell.options && cell.options.autoPageCharWeight) {
			let CHR1 = 2.3 + (cell.options && cell.options.autoPageCharWeight ? cell.options.autoPageCharWeight : 0) // Character Constant
			let CPL1 = ((colWidth / EMU_PER_POINT) * EMU_PER_INCH) / ((cell.options && cell.options.fontSize ? cell.options.fontSize : DEF_FONT_SIZE) / CHR1) // Chars-Per-Line
			console.log(`cell.options.autoPageCharWeight: '${cell.options.autoPageCharWeight}' => CPL: ${CPL1}`)
			let CHR2 = 2.3 + 0
			let CPL2 = ((colWidth / EMU_PER_POINT) * EMU_PER_INCH) / ((cell.options && cell.options.fontSize ? cell.options.fontSize : DEF_FONT_SIZE) / CHR2) // Chars-Per-Line
			console.log(`cell.options.autoPageCharWeight: '0' => CPL: ${CPL2}`)
		}
	*/

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
	let newLine: TableCell[] = []
	inputCells.forEach((cell) => {
		if (typeof cell.text !== 'string') return

		if (cell.text.includes('\n')) {
			// Each non-last \n-split part ends with a forced paragraph break; the last
			// part accumulates in newLine so subsequent runs stay on the same paragraph.
			const parts = cell.text.split('\n')
			parts.forEach((part, partIdx) => {
				const isLastPart = partIdx === parts.length - 1
				if (isLastPart) {
					newLine.push({ _type: SlideObjectType.tablecell, text: part, options: cell.options })
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
			newLine.push({ _type: SlideObjectType.tablecell, text: cell.text.trim(), options: cell.options })
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
		const lineTokens: TableCell[] = []
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
		let lineCells: TableCell[] = []
		let strCurrLine = ''

		line.forEach((word) => {
			const wordText = String(word.text || '')
			// A: create new line when horizontal space is exhausted
			if (strCurrLine.length + wordText.length > CPL) {
				parsedLines.push(lineCells)
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

// ===== Auto-page engine =====
// The in-memory core (fully Node-testable). Internally signposted by the
// `// STEP 1..7` markers below: margins → column count → widths → the main
// row-iteration/overflow loop (STEP 6) → final-slide flush (STEP 7).

/**
 * Takes an array of table rows and breaks into an array of slides, which contain the calculated amount of table rows that fit on that slide
 * @param {TableCell[][]} tableRows - table rows
 * @param {TableToSlidesProps} tableProps - table2slides properties
 * @param {PresLayout} presLayout - presentation layout
 * @param {SlideLayoutInternal} masterSlide - master slide
 * @return {TableRowSlide[]} array of table rows
 */
export function getSlidesForTableRows(
	tableRows: TableCell[][] = [],
	tableProps: TableToSlidesProps = {},
	presLayout: PresLayout,
	masterSlide?: SlideLayoutInternal | null
): TableRowSlide[] {
	let arrInchMargins = DEF_SLIDE_MARGIN_IN
	let emuSlideTabW: number
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

	function calcSlideTabH(): void {
		let emuStartY = 0
		if (tableRowSlides.length === 0) emuStartY = tablePropY || inch2Emu(arrInchMargins[0])
		if (tableRowSlides.length > 0) emuStartY = inch2Emu(tableProps.autoPageSlideStartY || arrInchMargins[0])
		emuSlideTabH = (tablePropH || presLayout.height) - emuStartY - inch2Emu(arrInchMargins[2])
		// EXPLICIT-H FIX: an explicit `h` is the table's height (an extent), not a
		// bottom coordinate, so — unlike `presLayout.height` — the first slide must NOT subtract the
		// start-Y from it. Otherwise a table that begins mid-slide gets a tiny first page (only a few
		// rows) while later pages, which already clamp to `h`, render normally. Mirror the
		// subsequent-slide rule below: never let an explicit `h` shrink the usable height below `h`.
		if (tableRowSlides.length === 0 && tablePropH && emuSlideTabH < tablePropH) emuSlideTabH = tablePropH
		if (tableRowSlides.length > 1) {
			// D: RULE: Use margins for starting point after the initial Slide, not `opt.y`
			if (typeof tableProps.autoPageSlideStartY === 'number') {
				emuSlideTabH = (tablePropH || presLayout.height) - inch2Emu(tableProps.autoPageSlideStartY + arrInchMargins[2])
			} else if (tablePropY) {
				emuSlideTabH =
					(tablePropH || presLayout.height) -
					inch2Emu(
						(tablePropY / EMU_PER_INCH < arrInchMargins[0] ? tablePropY / EMU_PER_INCH : arrInchMargins[0]) +
							arrInchMargins[2]
					)
				// Use whichever is greater: area between margins or the table H provided (dont shrink usable area - the whole point of over-riding Y on paging is to *increase* usable space)
				if (emuSlideTabH < tablePropH) emuSlideTabH = tablePropH
			}
		}

		// GUARD: a non-positive height, or one too small to fit even a single line of the base font,
		// means no row ever fits. That previously emitted degenerate empty overflow pages (rows:[]),
		// which made the recursive addTable throw "Array expected", or — with `h` smaller than one
		// line — placed one row per slide forever. Ignore the unusable height, fall back to the full
		// slide area between margins, and warn once rather than emit a broken table.
		const emuBaseLineH = inch2Emu(
			((typeof tableProps.fontSize === 'number' ? tableProps.fontSize : DEF_FONT_SIZE) *
				(LINEH_MODIFIER + (tableProps.autoPageLineWeight || 0))) /
				100
		)
		if (emuSlideTabH < emuBaseLineH) {
			const emuStartY =
				tableRowSlides.length === 0
					? tablePropY || inch2Emu(arrInchMargins[0])
					: inch2Emu(tableProps.autoPageSlideStartY || arrInchMargins[0])
			const fallbackH = presLayout.height - emuStartY - inch2Emu(arrInchMargins[2])
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
		// Important: Use default size as zero cell margin is causing our tables to be too large and touch bottom of slide!
		if (!tableProps.slideMargin && tableProps.slideMargin !== 0) tableProps.slideMargin = DEF_SLIDE_MARGIN_IN[0]

		if (masterSlide && typeof masterSlide._margin !== 'undefined') {
			if (Array.isArray(masterSlide._margin)) arrInchMargins = masterSlide._margin
			else if (!isNaN(Number(masterSlide._margin))) {
				arrInchMargins = [
					Number(masterSlide._margin),
					Number(masterSlide._margin),
					Number(masterSlide._margin),
					Number(masterSlide._margin),
				]
			}
		} else if (tableProps.slideMargin || tableProps.slideMargin === 0) {
			if (Array.isArray(tableProps.slideMargin)) arrInchMargins = tableProps.slideMargin
			else if (!isNaN(tableProps.slideMargin))
				arrInchMargins = [
					tableProps.slideMargin,
					tableProps.slideMargin,
					tableProps.slideMargin,
					tableProps.slideMargin,
				]
		}

		if (tableProps.verbose)
			console.log(`| arrInchMargins .................................. = [${arrInchMargins.join(', ')}]`)
	}

	// STEP 2: Calculate number of columns
	{
		// NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
		// ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
		const firstRow = tableRows[0] || []
		firstRow.forEach((cell) => {
			if (!cell) cell = { _type: SlideObjectType.tablecell }
			const cellOpts = cell.options || null
			numCols += Number(cellOpts?.colspan ? cellOpts.colspan : 1)
		})
		if (tableProps.verbose) console.log(`| numCols ......................................... = ${numCols}`)
	}

	// Track per-column remaining rowspan depths so we can suppress page breaks that would
	// fall inside a rowspan group. colSpanDepths[c] = how many more rows column c is still
	// occupied by a rowspan that started in a previous row.
	const colSpanDepths: number[] = new Array<number>(numCols).fill(0)

	// STEP 3: Calculate width using tableProps.colW if possible
	if (!tablePropW && tableProps.colW) {
		tableCalcW = Array.isArray(tableProps.colW)
			? tableProps.colW.reduce((p, n) => p + n) * EMU_PER_INCH
			: tableProps.colW * numCols || 0
		if (tableProps.verbose)
			console.log(`| tableCalcW ...................................... = ${tableCalcW / EMU_PER_INCH}`)
	}

	// STEP 4: Calculate usable width now that total usable space is known (`emuSlideTabW`)
	{
		emuSlideTabW =
			tableCalcW || inch2Emu((tablePropX ? tablePropX / EMU_PER_INCH : arrInchMargins[1]) + arrInchMargins[3])
		if (tableProps.verbose)
			console.log(`| emuSlideTabW .................................... = ${(emuSlideTabW / EMU_PER_INCH).toFixed(1)}`)
	}

	// STEP 5: Calculate column widths if not provided (emuSlideTabW will be used below to determine lines-per-col)
	if (!tableProps.colW || !Array.isArray(tableProps.colW)) {
		if (tableProps.colW && !isNaN(Number(tableProps.colW))) {
			const arrColW: number[] = []
			const colW = Number(tableProps.colW)
			const firstRow = tableRows[0] || []
			firstRow.forEach(() => arrColW.push(colW))
			tableProps.colW = []
			arrColW.forEach((val) => {
				if (Array.isArray(tableProps.colW)) tableProps.colW.push(val)
			})
		} else {
			// No column widths provided? Then distribute cols.
			tableProps.colW = []
			for (let iCol = 0; iCol < numCols; iCol++) {
				tableProps.colW.push(emuSlideTabW / EMU_PER_INCH / numCols)
			}
		}
	}

	// Resolve a row's explicit height (inches) from the original `rowH` *array*, keyed by original
	// row index. A single-number `rowH` is left to propagate via table options (it applies uniformly,
	// so it needs no per-row remapping); only the array form is index-sensitive after pagination.
	const resolveRowH = (origRowIdx: number): number | undefined =>
		Array.isArray(tableProps.rowH) && typeof tableProps.rowH[origRowIdx] === 'number'
			? tableProps.rowH[origRowIdx]
			: undefined

	// STEP 6: **MAIN** Iterate over rows, add table content, create new slides as rows overflow
	let newTableRowSlide: TableRowSlide = { rows: [] as TableRow[], rowH: [] as Array<number | undefined> }
	tableRows.forEach((row, iRow) => {
		// A: Row variables — detect active rowspan at the start of this row so we can
		// suppress page breaks that would split a rowspan group across slides.
		const hasActiveRowSpan = colSpanDepths.some((d) => d > 0)
		const rowCellLines: AutoPageCell[] = []
		let maxCellMarTopEmu = 0
		let maxCellMarBtmEmu = 0

		// B: Create new row in data model, calc `maxCellMar*`
		let currTableRow: TableRow = []
		row.forEach((cell) => {
			currTableRow.push({
				_type: SlideObjectType.tablecell,
				text: [],
				options: cell.options,
			})

			// Cell margins are inches (see `marginToEmu`); prefer the cell's own top/bottom margin, else the table's.
			const cellMargin = Array.isArray(cell.options?.margin) ? cell.options.margin : undefined
			const tableMargin = Array.isArray(tableProps.margin) ? tableProps.margin : null
			if (cellMargin?.[0] && marginToEmu(cellMargin[0]) > maxCellMarTopEmu)
				maxCellMarTopEmu = marginToEmu(cellMargin[0])
			else if (tableMargin?.[0] && marginToEmu(tableMargin[0]) > maxCellMarTopEmu)
				maxCellMarTopEmu = marginToEmu(tableMargin[0])
			if (cellMargin?.[2] && marginToEmu(cellMargin[2]) > maxCellMarBtmEmu)
				maxCellMarBtmEmu = marginToEmu(cellMargin[2])
			else if (tableMargin?.[2] && marginToEmu(tableMargin[2]) > maxCellMarBtmEmu)
				maxCellMarBtmEmu = marginToEmu(tableMargin[2])
		})

		// C: Calc usable vertical space/table height. Set default value first, adjust below when necessary.
		calcSlideTabH()
		emuTabCurrH += maxCellMarTopEmu + maxCellMarBtmEmu // Start row height with margins
		if (tableProps.verbose && iRow === 0)
			console.log(
				`| SLIDE [${tableRowSlides.length}]: emuSlideTabH ...... = ${(emuSlideTabH / EMU_PER_INCH).toFixed(1)} `
			)

		// D: --==[[ BUILD DATA SET ]]==-- (iterate over cells: split text into lines[], set `lineHeight`)
		row.forEach((cell, iCell) => {
			const newCellOptions = cell.options || {}
			const newCell: AutoPageCell = {
				_type: SlideObjectType.tablecell,
				_lines: [],
				_lineHeight: inch2Emu(
					((cell.options?.fontSize
						? cell.options.fontSize
						: tableProps.fontSize
							? tableProps.fontSize
							: DEF_FONT_SIZE) *
						(LINEH_MODIFIER + (tableProps.autoPageLineWeight ? tableProps.autoPageLineWeight : 0))) /
						100
				),
				text: [],
				options: newCellOptions,
			}

			// E-1: Exempt cells with `rowspan` from increasing lineHeight (or we could create a new slide when unecessary!)
			if (newCellOptions.rowspan) newCell._lineHeight = 0

			// E-2: The parseTextToLines method uses `autoPageCharWeight`, so inherit from table options
			newCellOptions.autoPageCharWeight = tableProps.autoPageCharWeight || undefined

			// E-3: **MAIN** Parse cell contents into lines based upon col width, font, etc
			const tableColW = Array.isArray(tableProps.colW) ? tableProps.colW : []
			let totalColW = tableColW[iCell] ?? 0
			const cellColspan = cell.options?.colspan
			if (cellColspan) {
				totalColW = tableColW
					.filter((_cell, idx) => idx >= iCell && idx < idx + cellColspan)
					.reduce((prev, curr) => prev + curr)
			}

			// E-4: Create lines based upon available column width
			newCell._lines = parseTextToLines(cell, totalColW, false)

			// E-5: Add cell to array
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
		 * - `rowCellLines` ..: [ TableCell, TableCell, TableCell ]
		 * - `TableCell` .....: { _type: 'tablecell', _lines: TableCell[], _lineHeight: 10 }
		 * - `_lines` ........: [ {_type: 'tablecell', text: 'cell-1,line-1', options: {…}}, {_type: 'tablecell', text: 'cell-1,line-2', options: {…}} }
		 * - `_lines` is TableCell[] (the 1-N words in the line)
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
				const newRows: TableRow[] = []
				newTableRowSlide = { rows: newRows, rowH: [] as Array<number | undefined> }

				// D: reset working/curr row
				currTableRow = []
				row.forEach((cell) => currTableRow.push({ _type: SlideObjectType.tablecell, text: [], options: cell.options }))

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
				if (tableProps.autoPageRepeatHeader && tableProps._arrObjTabHeadRows) {
					tableProps._arrObjTabHeadRows.forEach((row, headIdx) => {
						const newHeadRow: TableRow = []
						let maxLineHeight = 0
						row.forEach((cell) => {
							newHeadRow.push(cell)
							if ((cell._lineHeight || 0) > maxLineHeight) maxLineHeight = cell._lineHeight || 0
						})
						newTableRowSlide.rows.push(newHeadRow)
						// Repeated header rows are the original leading rows, so carry their configured height.
						newTableRowSlide.rowH?.push(resolveRowH(headIdx))
						// NOTE: possible imprecision — this accumulates line height only; cell top/bottom
						// margins are not added, so autoPage row-height estimates can run slightly short.
						emuTabCurrH += maxLineHeight
					})
				}

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

		// F: Flush/capture row buffer before it resets at the top of this loop
		if (currTableRow.length > 0) {
			newTableRowSlide.rows.push(currTableRow)
			newTableRowSlide.rowH?.push(resolveRowH(iRow))
		}

		// G: Update colSpanDepths for the next row's hasActiveRowSpan check.
		// Snapshot occupied columns *before* adding new spans from this row so that
		// cells in this row are placed correctly even when the row itself starts spans.
		const occupiedBefore = [...colSpanDepths]
		let colCursor = 0
		row.forEach((cell) => {
			while (colCursor < numCols && (occupiedBefore[colCursor] ?? 0) > 0) colCursor++
			const cellColspan = cell.options?.colspan ?? 1
			const cellRowspan = cell.options?.rowspan ?? 1
			if (cellRowspan > 1) {
				for (let c = 0; c < cellColspan && colCursor + c < numCols; c++) {
					colSpanDepths[colCursor + c] = cellRowspan
				}
			}
			colCursor += cellColspan
		})
		// Consume one row from every active span (including ones just opened above).
		for (let c = 0; c < numCols; c++) {
			const depth = colSpanDepths[c] ?? 0
			if (depth > 0) colSpanDepths[c] = depth - 1
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
