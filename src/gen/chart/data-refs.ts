/**
 * ts-pptx: Chart Series-Data ↔ Worksheet-Cell Mapping
 *
 * The tiny pure helpers that both chart package parts share: the empty-array-safe
 * accessors over the normalized (internal) chart-series arrays, and the builders that
 * turn a (column, row) into an absolute reference into the chart's embedded `Sheet1`.
 * The embedded-workbook builder ({@link ../chart/embed-xlsx}) writes the cells; the
 * chart-XML builder ({@link ../chart/chart-xml}) points `<c:f>` formulas back at them —
 * so this mapping lives in one place to keep column letters and `$`-anchoring from
 * drifting between the two sides.
 */

import { LETTERS } from '../../constants-internal.js'
import type { OptsChartDataInternal } from '../../types/internal.js'

// ===== Series-data accessors =====
// The normalized (internal) chart-series arrays are populated at addChart time but stay
// optional on OptsChartDataInternal; read them through these accessors with an empty-array
// fallback so the OOXML/worksheet assembly never dereferences `undefined`. They also
// tolerate an absent series (`data[0]` on an empty set) by returning an empty array.
export const dataLabels = (d: OptsChartDataInternal | undefined): string[][] => d?.labels ?? []
export const dataValues = (d: OptsChartDataInternal | undefined): number[] => d?.values ?? []
export const dataSizes = (d: OptsChartDataInternal | undefined): number[] => d?.sizes ?? []
// The first label group of a series (`labels[0]`), empty when the series or group is absent.
export const firstLabelGroup = (d: OptsChartDataInternal | undefined): string[] => dataLabels(d)[0] ?? []

// ===== Worksheet-cell references =====

/**
 * Calc and return excel column name for a given column length
 * @param colIndex column index
 * @return column name
 * @example 1 returns 'A'
 * @example 27 returns 'AA'
 */
export function getExcelColName(colIndex: number): string {
	let colStr: string
	const colIdx = colIndex - 1 // Subtract 1 so `LETTERS[columnIndex]` returns "A" etc

	if (colIdx <= 25) {
		// A-Z
		colStr = LETTERS[colIdx] ?? ''
	} else {
		// AA-ZZ (ZZ = index 702)
		colStr = `${LETTERS[Math.floor(colIdx / LETTERS.length - 1)]}${LETTERS[colIdx % LETTERS.length]}`
	}

	return colStr
}

/**
 * Build an absolute single-cell reference into the embedded workbook's `Sheet1`,
 * as emitted in a chart series `<c:f>` formula: `Sheet1!$C$2`. Centralizes the
 * `$`-anchoring so the column letter and `$` placement can't drift between sites.
 * @param colIndex 1-based column index (1 => 'A')
 * @param row 1-based row number
 */
export function sheetCellRef(colIndex: number, row: number): string {
	return `Sheet1!$${getExcelColName(colIndex)}$${row}`
}

/**
 * Build an absolute range reference into the embedded workbook's `Sheet1`, as
 * emitted in a chart series `<c:f>` formula: `Sheet1!$C$2:$C$6`. A same-column
 * range (the common case) passes the same index for `colFrom`/`colTo`; deriving
 * the column letter once per endpoint removes the duplicated `getExcelColName`
 * call that was a real off-by-one surface.
 * @param colFrom 1-based start column index
 * @param rowFrom 1-based start row
 * @param colTo 1-based end column index
 * @param rowTo 1-based end row
 */
export function sheetRangeRef(colFrom: number, rowFrom: number, colTo: number, rowTo: number): string {
	return `Sheet1!$${getExcelColName(colFrom)}$${rowFrom}:$${getExcelColName(colTo)}$${rowTo}`
}

/**
 * The shape of the embedded worksheet every series in one chart is laid out against.
 *
 * Both numbers are facts about the FIRST series, because that is how the workbook is written:
 * `data[0]`'s label groups take the leading columns and its categories decide the row count,
 * then every series gets one column. Neither is a fact about the series being emitted.
 */
export interface SheetLayout {
	/** How many leading columns the label groups occupy. */
	labelCols: number
	/** How many data rows the sheet has, header row excluded. */
	rowCount: number
}

/**
 * The worksheet layout for one chart's series set.
 *
 * The chart XML derived both numbers per series instead, from that series' own labels, so the
 * two sides of the mapping this module exists to keep in step disagreed the moment a caller
 * labelled only the first series -- which is the shape the plot builders' own worked example
 * shows. From series 1 on that produced `Sheet1!$A$2:$$1`, a `<c:val>` range pointing backwards
 * at series A's column, and a `<c:tx>` naming series A's header. The workbook was right
 * throughout; only the references were wrong.
 *
 * The row-count fallback is the workbook's: a category-less layout (a histogram feeds raw
 * observations with no labels) sizes the sheet from the longest value series instead.
 * @param data - the chart's normalized series, in data order
 */
export function sheetLayout(data: readonly OptsChartDataInternal[]): SheetLayout {
	return {
		labelCols: dataLabels(data[0]).length,
		rowCount: firstLabelGroup(data[0]).length || Math.max(0, ...data.map((series) => dataValues(series).length)),
	}
}

/**
 * The 1-based worksheet column a series' values occupy: the label groups come first, then one
 * column per series in data order.
 *
 * This is the one number the chart XML and the embedded workbook have to agree on, and six
 * emitters wrote the expression out. It belongs beside the reference builders for the same
 * reason they do.
 * @param d - the normalized series
 * @param sheet - the chart's worksheet layout, from {@link sheetLayout}
 */
export const seriesColumn = (d: OptsChartDataInternal, sheet: SheetLayout): number => d._dataIndex + sheet.labelCols + 1

/**
 * The category-name range in column A, header row excluded: `Sheet1!$A$2:$A$<count + 1>`.
 *
 * Six plot families built this string by hand around the very module whose header says the
 * mapping lives in one place so `$`-anchoring cannot drift between the two sides.
 * @param count - how many categories the series carries
 */
export const categoryRange = (count: number): string => sheetRangeRef(1, 2, 1, count + 1)
