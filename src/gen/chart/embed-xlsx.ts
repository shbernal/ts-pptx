/**
 * ts-pptx: Chart Embedded-Workbook Generation
 *
 * Builds the embedded `.xlsx` workbook that backs a chart's cached data — the data
 * source PowerPoint opens when a user edits the chart. `createExcelWorksheet` writes
 * the workbook plus the chart part + its `.rels` into the presentation package;
 * `buildEmbeddedWorksheet` / `buildChartRelsXml` are also reused by the read-side
 * injection path (`TsPptx.extractSlides`). Everything here is a pure string/bytes
 * builder — no I/O beyond the passed-in ZipWriter, no mutation of the presentation model.
 *
 * The chart's `chart.xml` DrawingML lives in `./chart-xml.ts`; the series↔worksheet-cell
 * mapping the two sides share lives in `./data-refs.ts`.
 */

import { XML_DECL } from '../../constants-internal.js'
import type { SlideRelChart, OptsChartDataInternal } from '../../types/internal.js'
import { ZipWriter } from '../../zip.js'
import { el, raw, voidEl } from '../oxml/el.js'
import {
	CHART_COLOR_STYLE_REL,
	CHART_STYLE_REL,
	CORE_PROPS_CONTENT_TYPE,
	CORE_PROPS_REL,
	EXTENDED_PROPS_CONTENT_TYPE,
	EXTENDED_PROPS_REL,
	OFFICE_DOCUMENT_REL,
	OFFICE_REL,
	PACKAGE_REL,
	RELATIONSHIPS_CONTENT_TYPE,
	THEME_CONTENT_TYPE,
	THEME_REL,
	OD_CONTENT,
} from '../../ooxml/rel-types.js'
import { OOXML_NS } from '../../ooxml/namespaces.js'
import { relationshipEl, relationshipsEl } from '../opc/rels.js'
import { CORE_PROPS_NS, coreTimestamp } from '../opc/core.js'
import { dataLabels, getExcelColName, labelSeries, type WorksheetLayout, worksheetLayout } from './data-refs.js'
import { makeXmlCharts } from './chart-xml.js'
import { makeXmlChartEx } from './chartex-xml.js'
import { makeChartExColorsXml, makeChartExStyleXml } from './chartex-style.js'
import { FMT_SCHEME_XML } from '../oxml/fmt-scheme.js'

/** MS chart-extension relationship types (chartEx style + color-style sidecar parts). */

/** The SpreadsheetML namespace every part of the embedded workbook is written in. */
const SML_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
/** The prefix the embedded workbook's SpreadsheetML part types share. */
const SML_CT = OD_CONTENT + 'spreadsheetml.'

/** One `<si>` shared string carrying literal text. */
const sharedString = (text: string): string => el('si', null, raw(el('t', null, text)))

/**
 * The embedded workbook's style sheet, captured verbatim from an Excel-authored chart workbook.
 * @raw-xml-asset
 */
const XLSX_STYLES_XML =
	'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="0" formatCode="General"/></numFmts><fonts count="4"><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="10"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="18"/><color indexed="8"/>' +
	'<name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><dxfs count="0"/><tableStyles count="0"/><colors><indexedColors><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ffff0000"/><rgbColor rgb="ff00ff00"/><rgbColor rgb="ff0000ff"/>' +
	'<rgbColor rgb="ffffff00"/><rgbColor rgb="ffff00ff"/><rgbColor rgb="ff00ffff"/><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ff878787"/><rgbColor rgb="fff9f9f9"/></indexedColors></colors></styleSheet>\n'

/**
 * The Office theme the embedded workbook ships, captured verbatim. PowerPoint reads it when a user
 * opens the chart's data, so it has to be exactly the bytes Office writes.
 *
 * Its `<a:fmtScheme>` comes from {@link FMT_SCHEME_XML} rather than sitting inline: `theme1.xml`
 * carried a byte-identical 2661-character copy of it, and one asset transcribed twice is the
 * arrangement where a correction lands in only one of them. The two literals around the splice
 * are unchanged character for character, which the byte-identity gate is what proves.
 * @raw-xml-asset
 */
const XLSX_THEME_XML =
	'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme>' +
	FMT_SCHEME_XML +
	'</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>'

/**
 * Build the chart's embedded Excel workbook as a standalone OPC package and
 * return its bytes — the data source PowerPoint opens when a user edits the
 * chart's data. Pure (no zip side effects), so both the package write path
 * ({@link createExcelWorksheet}) and the read-side injection path
 * (`TsPptx.extractSlides`) can reuse it.
 * @param {SlideRelChart} chartObject - chart object
 * @return {Uint8Array} the embedded `.xlsx` package bytes
 */
// ===== Embedded worksheet =====

export function buildEmbeddedWorksheet(chartObject: SlideRelChart): Uint8Array {
	const data = chartObject.data

	{
		// The embedded workbook is its own OPC package: build it in a nested ZipWriter,
		// then embed its bytes (no folder scaffolding; fflate emits no directory entries).
		const zipExcel = new ZipWriter()
		const layout = worksheetLayout(chartObject)

		// B: Add core contents
		{
			const override = (partName: string, contentType: string): string =>
				voidEl('Override', { PartName: partName, ContentType: contentType }, { openPrefix: '  ' })
			zipExcel.add(
				'[Content_Types].xml',
				XML_DECL +
					el('Types', { xmlns: OOXML_NS.ct }, [
						raw(
							voidEl('Default', { Extension: 'rels', ContentType: RELATIONSHIPS_CONTENT_TYPE }, { openPrefix: '  ' })
						),
						raw(voidEl('Default', { Extension: 'xml', ContentType: 'application/xml' }, { openPrefix: '  ' })),
						raw(override('/xl/workbook.xml', SML_CT + 'sheet.main+xml')),
						raw(override('/xl/worksheets/sheet1.xml', SML_CT + 'worksheet+xml')),
						raw(override('/xl/theme/theme1.xml', THEME_CONTENT_TYPE)),
						raw(override('/xl/styles.xml', SML_CT + 'styles+xml')),
						raw(override('/xl/sharedStrings.xml', SML_CT + 'sharedStrings+xml')),
						raw(override('/xl/tables/table1.xml', SML_CT + 'table+xml')),
						raw(override('/docProps/core.xml', CORE_PROPS_CONTENT_TYPE)),
						raw(override('/docProps/app.xml', EXTENDED_PROPS_CONTENT_TYPE)),
					]) +
					'\n'
			)
			zipExcel.add(
				'_rels/.rels',
				XML_DECL +
					relationshipsEl([
						relationshipEl('rId1', CORE_PROPS_REL, 'docProps/core.xml'),
						relationshipEl('rId2', EXTENDED_PROPS_REL, 'docProps/app.xml'),
						relationshipEl('rId3', OFFICE_DOCUMENT_REL, 'xl/workbook.xml'),
					]) +
					'\n'
			)
			const headingPairs = el('vt:vector', { size: 2, baseType: 'variant' }, [
				raw(el('vt:variant', null, raw(el('vt:lpstr', null, 'Worksheets')))),
				raw(el('vt:variant', null, raw(el('vt:i4', null, 1)))),
			])
			zipExcel.add(
				'docProps/app.xml',
				XML_DECL +
					el(
						'Properties',
						{
							xmlns: OOXML_NS.ep,
							'xmlns:vt': OOXML_NS.vt,
						},
						[
							raw(el('Application', null, 'Microsoft Macintosh Excel')),
							raw(el('DocSecurity', null, 0)),
							raw(el('ScaleCrop', null, 'false')),
							raw(el('HeadingPairs', null, raw(headingPairs))),
							raw(
								el(
									'TitlesOfParts',
									null,
									raw(el('vt:vector', { size: 1, baseType: 'lpstr' }, raw(el('vt:lpstr', null, 'Sheet1'))))
								)
							),
							raw(el('Company', null)),
							raw(el('LinksUpToDate', null, 'false')),
							raw(el('SharedDoc', null, 'false')),
							raw(el('HyperlinksChanged', null, 'false')),
							raw(el('AppVersion', null, '16.0300')),
						]
					) +
					'\n'
			)
			// One reading of the clock for both stamps: two calls make `created` and `modified`
			// disagree whenever the build crosses a millisecond, which no reader notices and
			// every byte-diff does.
			const now = coreTimestamp()
			zipExcel.add(
				'docProps/core.xml',
				XML_DECL +
					el('cp:coreProperties', CORE_PROPS_NS, [
						raw(el('dc:creator', null, 'TsPptx')),
						raw(el('cp:lastModifiedBy', null, 'TsPptx')),
						raw(el('dcterms:created', { 'xsi:type': 'dcterms:W3CDTF' }, now)),
						raw(el('dcterms:modified', { 'xsi:type': 'dcterms:W3CDTF' }, now)),
					])
			)
			zipExcel.add(
				'xl/_rels/workbook.xml.rels',
				XML_DECL +
					// Ids are deliberately out of order (3/2/1/4) — that is how this part has
					// always been emitted, and rel order is byte-significant.
					relationshipsEl([
						relationshipEl('rId3', OFFICE_REL + 'styles', 'styles.xml'),
						relationshipEl('rId2', THEME_REL, 'theme/theme1.xml'),
						relationshipEl('rId1', OFFICE_REL + 'worksheet', 'worksheets/sheet1.xml'),
						relationshipEl('rId4', OFFICE_REL + 'sharedStrings', 'sharedStrings.xml'),
					])
			)
			zipExcel.add('xl/styles.xml', XML_DECL + XLSX_STYLES_XML)
			zipExcel.add('xl/theme/theme1.xml', XML_DECL + XLSX_THEME_XML)
			zipExcel.add(
				'xl/workbook.xml',
				XML_DECL +
					el(
						'workbook',
						{
							xmlns: SML_NS,
							'xmlns:r': OOXML_NS.r,
							'xmlns:mc': OOXML_NS.mc,
							'mc:Ignorable': 'x15',
							'xmlns:x15': 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main',
						},
						[
							raw(voidEl('fileVersion', { appName: 'xl', lastEdited: 7, lowestEdited: 6, rupBuild: 10507 })),
							raw(voidEl('workbookPr')),
							raw(
								el(
									'bookViews',
									null,
									raw(
										voidEl('workbookView', {
											xWindow: 0,
											yWindow: 500,
											windowWidth: 20960,
											windowHeight: 15960,
										})
									)
								)
							),
							raw(el('sheets', null, raw(voidEl('sheet', { name: 'Sheet1', sheetId: 1, 'r:id': 'rId1' })))),
							raw(voidEl('calcPr', { calcId: 0, concurrentCalc: 0 })),
						]
					) +
					'\n'
			)
			zipExcel.add(
				'xl/worksheets/_rels/sheet1.xml.rels',
				XML_DECL + relationshipsEl([relationshipEl('rId1', OFFICE_REL + 'table', '../tables/table1.xml')]) + '\n'
			)
		}

		zipExcel.add('xl/sharedStrings.xml', buildXlsxSharedStrings(data, layout))
		zipExcel.add('xl/tables/table1.xml', buildXlsxTable(layout))
		zipExcel.add('xl/worksheets/sheet1.xml', buildXlsxSheet(data, layout))

		// Done — return the embedded workbook bytes for the caller to place.
		return zipExcel.toBytes()
	}
}

/** One category label the sheet writes: its label column (1-based), its row (0-based) and its text. */
interface LabelCell {
	col: number
	row: number
	label: string
}

/**
 * The label cells of a category chart's sheet, outermost level first and down each column: the
 * order their shared strings are written in.
 *
 * A blank label has no cell and no string. The shared-string index of the cell at position `k` is
 * therefore `data.length + 1 + k`, after the blank entry and the series names, whether or not any
 * label before it was blank. The one-level arm used to index every row as if none were, so from
 * the first blank label on each cell read the next label's string or ran past the end.
 * @param data - the chart's series; the labels are {@link labelSeries}'
 */
function categoryLabelCells(data: readonly OptsChartDataInternal[]): LabelCell[] {
	// labels[0] is the leaf level and the last is the outermost, which takes column A.
	return dataLabels(labelSeries(data))
		.slice()
		.reverse()
		.flatMap((group, level) =>
			group.flatMap((label, row) => (label && label !== '' ? [{ col: level + 1, row, label }] : []))
		)
}

/**
 * Build the embedded workbook's `xl/sharedStrings.xml`: the column headers, and on a category sheet a
 * leading blank and the category labels.
 *
 * The headers are the layout's, which the chart part caches as well ({@link seriesHeader}).
 */
function buildXlsxSharedStrings(data: OptsChartDataInternal[], layout: WorksheetLayout): string {
	const headers = layout.columns.map((column) => sharedString(column.header)).join('')
	let count = layout.columns.length
	let uniqueCount = count
	let blank = ''
	let labels = ''
	if (layout.kind === 'category') {
		// A blank, the headers, then every non-blank label. `count` is how many cells refer to a
		// string: each label column's header points at the blank, then one per header and per label.
		const cells = categoryLabelCells(data)
		count = layout.labelCols + layout.columns.length + cells.length
		uniqueCount = 1 + layout.columns.length + cells.length
		labels = cells.map((cell) => sharedString(cell.label)).join('')
		// The leading entry is the blank the header row's label columns point at. Its two spellings
		// are not interchangeable: `<t/>` is the empty string, `<t xml:space="preserve"></t>` is a
		// preserved one, and next to character data that difference is content rather than layout.
		blank =
			layout.labelCols > 1
				? el('si', null, raw(voidEl('t')))
				: el('si', null, raw(el('t', { 'xml:space': 'preserve' })))
	}
	return XML_DECL + el('sst', { xmlns: SML_NS, count, uniqueCount }, [raw(blank), raw(headers), raw(labels)]) + '\n'
}

/**
 * Build the embedded workbook's `xl/tables/table1.xml` (the data table over the sheet range).
 *
 * Its range comes from the same layout as the sheet's `<dimension>`. The two were derived
 * separately once, and the bubble range used the column count as its row count, so one workbook
 * said `ref="A1:C3"` in this part and `<dimension ref="A1:C5"/>` in the other.
 *
 * Each column after the label block is named by its header. A scatter's used to be `X-Values0` and
 * `Y-Value 1`, and a bubble's X column `X-Values`, names no header cell held.
 */
function buildXlsxTable(layout: WorksheetLayout): string {
	const ref = `A1:${getExcelColName(layout.colCount)}${layout.rowCount + 1}`
	// The leading columns are the label groups; the named columns follow them.
	const columns =
		Array.from({ length: layout.labelCols }, (_unused, idx) =>
			voidEl('tableColumn', { id: idx + 1, name: `Column${idx + 1}` })
		).join('') + layout.columns.map((column) => voidEl('tableColumn', { id: column.col, name: column.header })).join('')
	return (
		XML_DECL +
		el('table', { xmlns: SML_NS, id: 1, name: 'Table1', displayName: 'Table1', ref, totalsRowShown: 0 }, [
			raw(el('tableColumns', { count: layout.colCount }, raw(columns))),
			raw(
				voidEl('tableStyleInfo', {
					showFirstColumn: 0,
					showLastColumn: 0,
					showRowStripes: 1,
					showColumnStripes: 0,
				})
			),
		])
	)
}

/**
 * Build the embedded workbook's `xl/worksheets/sheet1.xml` (header row + per-series data rows).
 */
function buildXlsxSheet(data: OptsChartDataInternal[], layout: WorksheetLayout): string {
	const { labelCols, colCount, rowCount } = layout
	/**
	 * One cell. `t="s"` marks a shared-string index; without it the value is a number.
	 *
	 * A gap is written as a present-but-empty `<v></v>`, which is how a missing value has always
	 * reached this sheet (`?? ''` at the call sites) and which Excel reads back as an empty cell.
	 * A **non-finite** number is the same thing: `<v>Infinity</v>` (or `NaN`, or `INF`) is not a
	 * number Excel will parse, and it refuses the whole workbook with 0x3EC rather than skipping
	 * the cell — a failure PowerPoint hides, because it does not parse the embedding on open. It
	 * surfaces on "Edit Data", by which point the chart's own cache is already clean: `numCachePt`
	 * (`./chart-parts.ts`) and the chartEx numeric dimension (`./chartex-data.ts`) both drop a
	 * non-finite point, with a `chart/non-finite-value` warning, before it reaches the chart part.
	 * Every numeric cell written here is mirrored by one of those caches, so the drop is silent on
	 * this side rather than warned about twice for the same value.
	 */
	const cell = (col: number, row: number, value: string | number, shared = false): string => {
		const cellValue = typeof value === 'number' && !Number.isFinite(value) ? '' : value
		return el('c', { r: `${getExcelColName(col)}${row}`, t: shared ? 's' : undefined }, raw(el('v', null, cellValue)))
	}
	/** One row, spanning the sheet's full column count. */
	const sheetRow = (row: number, span: number, cells: string): string =>
		el('row', { r: row, spans: `1:${span}` }, raw(cells))

	/* EX: a category sheet, labels in A and one column per series:
				-|---A---|--B--|--C--|--D--|
				1|       | Red | Amb | Grn |
				2|Jan-17 |   11|   22|   33|
				3|Feb-17 |   55|   43|   70|
				-|-------|-----|-----|-----|
			A scatter sheet has no label block: X in A and one column per Y series. A bubble sheet
			follows each Y column with that series' sizes. */
	let rows = ''

	// Header row. On a category sheet the label columns point at the blank shared string, which the
	// column headers follow; an xy sheet has neither, so its first header is string 0.
	const headerOffset = layout.kind === 'category' ? 1 : 0
	let header = ''
	for (let col = 1; col <= labelCols; col++) header += cell(col, 1, 0, true)
	layout.columns.forEach((column, k) => (header += cell(column.col, 1, headerOffset + k, true)))
	rows += sheetRow(1, colCount, header)

	// Each non-blank label's cell, by row, in column order. Its shared string follows the blank entry
	// and the headers, in the order the strings part writes them.
	const labelsByRow = new Map<number, string>()
	if (layout.kind === 'category') {
		categoryLabelCells(data).forEach(({ col, row }, k) =>
			labelsByRow.set(row, (labelsByRow.get(row) ?? '') + cell(col, row + 2, layout.columns.length + 1 + k, true))
		)
	}

	// One row per category, or per X value. A category-less chartEx layout (a histogram feeds
	// PowerPoint raw observations with no labels) has no label groups, so the row count falls back to
	// the longest value series and the leading label columns are simply skipped: values land in A.
	for (let idx = 0; idx < rowCount; idx++) {
		let cells = labelsByRow.get(idx) ?? ''
		for (const column of layout.columns) cells += cell(column.col, idx + 2, column.value(idx) ?? '')
		rows += sheetRow(idx + 2, colCount, cells)
	}

	const sheetView = el(
		'sheetViews',
		null,
		raw(
			el(
				'sheetView',
				{ tabSelected: 1, workbookViewId: 0 },
				raw(voidEl('selection', { activeCell: 'B1', sqref: 'B1' }))
			)
		)
	)
	return (
		XML_DECL +
		el(
			'worksheet',
			{
				xmlns: SML_NS,
				'xmlns:r': OOXML_NS.r,
				'xmlns:mc': OOXML_NS.mc,
				'mc:Ignorable': 'x14ac',
				'xmlns:x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
			},
			[
				raw(voidEl('dimension', { ref: `A1:${getExcelColName(colCount)}${rowCount + 1}` })),
				raw(sheetView),
				raw(voidEl('sheetFormatPr', { baseColWidth: 10, defaultRowHeight: 16 })),
				raw(el('sheetData', null, raw(rows))),
				raw(voidEl('pageMargins', { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 })),
				// NOTE: Intentionally no `<tableParts>`. A tablePart only works for scatter charts; every
				// other chart type reports a "cannot find linked file" error. The chart data can be
				// edited / range-selected without it, so it is deliberately never emitted.
			]
		) +
		'\n'
	)
}

/**
 * Build the standalone `.rels` for a chart part: a single `rId1` relationship to
 * the chart's embedded workbook (`Target`). Shared by the package write path and
 * the read-side injection path, which pass different (relative) embedding targets.
 * @param {string} embeddingTarget - the workbook target, relative to the chart part
 * @return {string} the chart part's `.rels` XML
 */
function buildChartRelsXml(embeddingTarget: string): string {
	// `voidEl` escapes the Target. The one in-tree caller passes an internally built
	// `../embeddings/Microsoft_Excel_WorksheetN.xlsx`, so that is a no-op on bytes;
	// it matters only for the read-side injection path, which supplies its own target.
	return XML_DECL + relationshipsEl([relationshipEl('rId1', PACKAGE_REL, embeddingTarget)])
}

/**
 * Build the `.rels` for a chartEx chart part: the embedded workbook (rId1, via `<cx:externalData>`)
 * plus the mandatory color-style (rId2) and chart-style (rId3) sidecar parts. PowerPoint treats a
 * chartEx part without the style/color rels as corrupt, so these are not optional.
 * @param {string} embeddingTarget - workbook target, relative to the chart part
 * @param {string} colorsTarget - `colors{N}.xml`, relative to the chart part
 * @param {string} styleTarget - `style{N}.xml`, relative to the chart part
 */
function buildChartExRelsXml(embeddingTarget: string, colorsTarget: string, styleTarget: string): string {
	return (
		XML_DECL +
		relationshipsEl([
			relationshipEl('rId1', PACKAGE_REL, embeddingTarget),
			relationshipEl('rId2', CHART_COLOR_STYLE_REL, colorsTarget),
			relationshipEl('rId3', CHART_STYLE_REL, styleTarget),
		])
	)
}

/**
 * Create the chart's embedded Excel worksheet and add the chart + workbook parts
 * to `zip` (package write path). The read-side injection path builds the same
 * parts itself from {@link buildEmbeddedWorksheet}, {@link buildChartRelsXml}, and
 * {@link makeXmlCharts}.
 * @param {SlideRelChart} chartObject - chart object
 * @param {ZipWriter} zip - zip writer the resulting XLSX (and chart parts) are added to
 * @return {Promise} promise of generating the XLSX file
 */
export async function createExcelWorksheet(chartObject: SlideRelChart, zip: ZipWriter): Promise<string> {
	// 1: Embed the workbook. The xlsx is itself a zip, so STORE it — re-DEFLATING
	//    already-compressed bytes wastes CPU.
	zip.add(`ppt/embeddings/Microsoft_Excel_Worksheet${chartObject.globalId}.xlsx`, buildEmbeddedWorksheet(chartObject), {
		store: true,
	})

	// 2: Create the chart part, its rels, and (for chartEx) the required style/color sidecar parts.
	const embeddingTarget = `../embeddings/Microsoft_Excel_Worksheet${chartObject.globalId}.xlsx`
	if (chartObject.isChartEx) {
		// chartEx charts REQUIRE a chart-style + color-style part or PowerPoint reports the deck as
		// corrupt (schema-valid but unopenable) — see gen/chart/chartex-style.ts.
		const colorsName = `colors${chartObject.globalId}.xml`
		const styleName = `style${chartObject.globalId}.xml`
		zip.add(`ppt/charts/${colorsName}`, makeChartExColorsXml())
		zip.add(`ppt/charts/${styleName}`, makeChartExStyleXml())
		zip.add(
			'ppt/charts/_rels/' + chartObject.fileName + '.rels',
			buildChartExRelsXml(embeddingTarget, colorsName, styleName)
		)
		zip.add(`ppt/charts/${chartObject.fileName}`, makeXmlChartEx(chartObject))
	} else {
		zip.add('ppt/charts/_rels/' + chartObject.fileName + '.rels', buildChartRelsXml(embeddingTarget))
		zip.add(`ppt/charts/${chartObject.fileName}`, makeXmlCharts(chartObject))
	}

	return ''
}
