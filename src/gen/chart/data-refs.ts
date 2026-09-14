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
import { InvalidOptionError } from '../../errors.js'
import type { OptsChartDataInternal, SlideRelChart } from '../../types/internal.js'
import { isBubbleChart, isScatterChart } from './chart-kind.js'

/**
 * The text a series' header cell holds, and so the text its `<c:tx>` caches: the series' name, or
 * the empty string for a series with none.
 *
 * The workbook and the chart part used to spell this apart. The sheet wrote an unnamed series as
 * `' '` where every cache held `''`, rewrote `X-Axis` to `X-Values` inside any name holding it, named
 * an unnamed bubble series `Y-Axis1`, and headed a bubble's X column `X-Axis` whatever it was called.
 * A bubble's size column has no name of its own and no cache reads its header, so it keeps
 * `Size<n>`.
 * @param series - the series
 * @param role - `values` for the column a series' values take, `sizes` for a bubble's size column
 */
export function seriesHeader(series: OptsChartDataInternal | undefined, role: 'values' | 'sizes' = 'values'): string {
	return role === 'sizes' ? `Size${series?._dataIndex ?? 0}` : (series?.name ?? '')
}

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

/**
 * The series a category sheet takes its label columns and row count from: the first one that
 * carries labels, else the first.
 *
 * On a plain category chart that is `data[0]`. On a combo it need not be: a scatter subchart's
 * first series is its X row, which carries no labels, so a combo that opens with a scatter gave the
 * sheet no label column while a bar or line after it still wrote `<c:cat>` against column A, which
 * then held the scatter's X values.
 * @param data - the chart's series, across every subchart
 */
export const labelSeries = (data: readonly OptsChartDataInternal[]): OptsChartDataInternal | undefined =>
	data.find((series) => dataLabels(series).length > 0) ?? data[0]

// ===== Worksheet-cell references =====

/** The last column a worksheet has, `XFD`. */
const MAX_WORKSHEET_COLUMN = 16384

/**
 * The worksheet column name for a 1-based column index: `A` to `Z`, `AA` to `ZZ`, `AAA` to `XFD`.
 *
 * Column names are bijective base 26. There is no zero digit, so each step takes one off before
 * dividing. This used to handle two letters at most, and column 703 came out as `undefinedA`.
 * @param colIndex column index, 1-based
 * @return column name
 * @example 1 returns 'A'
 * @example 27 returns 'AA'
 * @example 703 returns 'AAA'
 */
export function getExcelColName(colIndex: number): string {
	if (colIndex > MAX_WORKSHEET_COLUMN)
		throw new InvalidOptionError(
			'chart/too-many-columns',
			`A chart's worksheet would need column ${colIndex}, past XFD (column ${MAX_WORKSHEET_COLUMN}), the last column a worksheet has. Split the data across more than one chart.`
		)
	let name = ''
	for (let n = colIndex; n > 0; n = Math.floor((n - 1) / LETTERS.length)) {
		name = `${LETTERS[(n - 1) % LETTERS.length] ?? ''}${name}`
	}
	return name
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
 * Where everything in one chart's embedded worksheet falls: the workbook writer lays the sheet
 * out from it and every formula in the chart part points back into it.
 *
 * Every number is a fact about the whole chart, not about the series being emitted. The chart XML
 * once derived the label columns and row count per series, from that series' own labels, and the
 * two sides disagreed the moment a caller labelled only the first series.
 */
export interface WorksheetLayout {
	/**
	 * `category` for a sheet with a block of label columns whose header cells point at a blank string,
	 * whether or not any series carries labels; `xy` for a scatter or bubble sheet, which has neither.
	 * The workbook writer reads this rather than asking the chart type again.
	 */
	kind: 'category' | 'xy'
	/** Every column after the label block, in column order. */
	columns: readonly WorksheetColumn[]
	/** How many leading columns the label groups occupy. */
	labelCols: number
	/** How many data rows the sheet has, header row excluded. */
	rowCount: number
	/** How many columns the sheet has. */
	colCount: number
	/**
	 * The 1-based column a series' values occupy, and whose header row holds its name.
	 * @param dataIndex - the series' position across the whole chart, its `_dataIndex`
	 */
	valueColumn(dataIndex: number): number
}

/** One column of a chart's worksheet after its label block. */
export interface WorksheetColumn {
	/** The 1-based column index. */
	col: number
	/** The header cell's text. */
	header: string
	/**
	 * The value in one data row.
	 * @param row - the 0-based data row, header excluded
	 * @returns the value, or `null`/`undefined` for a gap
	 */
	value(row: number): number | null | undefined
}

/** A column holding one series' values under its name. */
const valuesColumn = (series: OptsChartDataInternal, col: number): WorksheetColumn => ({
	col,
	header: seriesHeader(series),
	value: (row) => dataValues(series)[row],
})

/**
 * The 1-based column a bubble series' sizes occupy on a bubble chart's sheet: the one after its
 * values.
 *
 * Only a standalone bubble chart has size columns. A combo cannot hold a bubble subchart
 * (`addChartDefinition` refuses one), so no other layout is ever asked.
 * @param dataIndex - the series' position in the chart, its `_dataIndex`; 1 or more
 */
export const bubbleSizeColumn = (dataIndex: number): number => dataIndex * 2 + 1

/**
 * The worksheet layout for one chart, decided by its kind.
 *
 * - A category chart, every combo and every chartEx layout: the label groups of
 *   {@link labelSeries} take the leading columns, outermost first, and its categories decide the
 *   row count. Then one column
 *   per series in data order. A category-less layout (a histogram feeds raw observations with no
 *   labels) sizes the sheet from the longest value series instead.
 * - Scatter: no label columns. The X row is column A and each Y series takes the next column,
 *   one row per X value.
 * - Bubble: no label columns. The X row is column A and each later series takes two columns, its
 *   values and then its sizes ({@link bubbleSizeColumn}).
 *
 * A combo is laid out as a category chart whatever its subcharts are, because one workbook is
 * written from the combined series. A scatter subchart's X row is therefore a column like any
 * other series.
 * @param rel - the chart, for its series and its normalized type
 */
export function worksheetLayout(rel: Pick<SlideRelChart, 'data' | 'opts'>): WorksheetLayout {
	const data = rel.data
	const type = rel.opts._type
	if (isBubbleChart(type)) {
		const valueColumn = (dataIndex: number): number => (dataIndex === 0 ? 1 : dataIndex * 2)
		return {
			kind: 'xy',
			columns: data.flatMap((series, idx): WorksheetColumn[] =>
				idx === 0
					? [valuesColumn(series, 1)]
					: [
							valuesColumn(series, valueColumn(series._dataIndex)),
							{
								col: bubbleSizeColumn(series._dataIndex),
								header: seriesHeader(series, 'sizes'),
								value: (row) => dataSizes(series)[row],
							},
						]
			),
			labelCols: 0,
			rowCount: dataValues(data[0]).length,
			// 1 for the X values, then 2 for every Y series.
			colCount: (data.length - 1) * 2 + 1,
			valueColumn,
		}
	}
	if (isScatterChart(type)) {
		return {
			kind: 'xy',
			columns: data.map((series) => valuesColumn(series, series._dataIndex + 1)),
			labelCols: 0,
			rowCount: dataValues(data[0]).length,
			colCount: data.length,
			valueColumn: (dataIndex) => dataIndex + 1,
		}
	}
	const labelled = labelSeries(data)
	const labelCols = dataLabels(labelled).length
	return {
		kind: 'category',
		columns: data.map((series) => valuesColumn(series, series._dataIndex + labelCols + 1)),
		labelCols,
		rowCount: firstLabelGroup(labelled).length || Math.max(0, ...data.map((series) => dataValues(series).length)),
		colCount: data.length + labelCols,
		valueColumn: (dataIndex) => dataIndex + labelCols + 1,
	}
}

/**
 * The category-name range in the leaf label column, header row excluded:
 * `Sheet1!$B$2:$B$<count + 1>` on a sheet with two label levels.
 *
 * Six plot families built this string by hand around the very module whose header says the mapping
 * lives in one place. It then hard-coded column A, which holds the outermost level: a leaf cache over
 * nested labels was referenced against the outer labels' cells. The workbook writes the leaf level in
 * the last label column, which for single-level labels is column A.
 * @param sheet - the chart's worksheet layout
 * @param count - how many categories the series carries
 */
export const categoryRange = (sheet: WorksheetLayout, count: number): string =>
	sheetRangeRef(sheet.labelCols, 2, sheet.labelCols, count + 1)
