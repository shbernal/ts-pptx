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
	CORE_PROPS_REL,
	EXTENDED_PROPS_REL,
	OFFICE_DOCUMENT_REL,
	OFFICE_REL,
	PACKAGE_REL,
	RELATIONSHIPS_CONTENT_TYPE,
	THEME_REL,
} from '../../ooxml/rel-types.js'
import { OOXML_NS } from '../../ooxml/namespaces.js'
import { relationshipEl, relationshipsEl } from '../opc/rels.js'
import { CORE_PROPS_NS, coreTimestamp } from '../opc/core.js'
import { dataLabels, dataValues, dataSizes, firstLabelGroup, getExcelColName } from './data-refs.js'
import { makeXmlCharts } from './chart-xml.js'
import { makeXmlChartEx } from './chartex-xml.js'
import { makeChartExColorsXml, makeChartExStyleXml } from './chartex-style.js'
import { isBubbleChart, isScatterChart } from './chart-kind.js'
import { FMT_SCHEME_XML } from '../oxml/fmt-scheme.js'

/** MS chart-extension relationship types (chartEx style + color-style sidecar parts). */

/** The SpreadsheetML namespace every part of the embedded workbook is written in. */
const SML_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
/** OPC content types, and the prefix the four SpreadsheetML part types share. */
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
const SML_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.'

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
		const intBubbleCols = (data.length - 1) * 2 + 1 // 1 for "X-Values", then 2 for every Y-Axis
		const IS_MULTI_CAT_AXES = (data[0]?.labels?.length ?? 0) > 1

		// B: Add core contents
		{
			const override = (partName: string, contentType: string): string =>
				voidEl('Override', { PartName: partName, ContentType: contentType }, { openPrefix: '  ' })
			zipExcel.add(
				'[Content_Types].xml',
				XML_DECL +
					el('Types', { xmlns: CT_NS }, [
						raw(
							voidEl('Default', { Extension: 'rels', ContentType: RELATIONSHIPS_CONTENT_TYPE }, { openPrefix: '  ' })
						),
						raw(voidEl('Default', { Extension: 'xml', ContentType: 'application/xml' }, { openPrefix: '  ' })),
						raw(override('/xl/workbook.xml', SML_CT + 'sheet.main+xml')),
						raw(override('/xl/worksheets/sheet1.xml', SML_CT + 'worksheet+xml')),
						raw(override('/xl/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml')),
						raw(override('/xl/styles.xml', SML_CT + 'styles+xml')),
						raw(override('/xl/sharedStrings.xml', SML_CT + 'sharedStrings+xml')),
						raw(override('/xl/tables/table1.xml', SML_CT + 'table+xml')),
						raw(override('/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml')),
						raw(override('/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml')),
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
							xmlns: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
							'xmlns:vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
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

		zipExcel.add('xl/sharedStrings.xml', buildXlsxSharedStrings(chartObject, data, intBubbleCols, IS_MULTI_CAT_AXES))
		zipExcel.add('xl/tables/table1.xml', buildXlsxTable(chartObject, data, intBubbleCols))
		zipExcel.add('xl/worksheets/sheet1.xml', buildXlsxSheet(chartObject, data, intBubbleCols, IS_MULTI_CAT_AXES))

		// Done — return the embedded workbook bytes for the caller to place.
		return zipExcel.toBytes()
	}
}

/**
 * Build the embedded workbook's `xl/sharedStrings.xml` (series names + category labels).
 */
function buildXlsxSharedStrings(
	chartObject: SlideRelChart,
	data: OptsChartDataInternal[],
	intBubbleCols: number,
	IS_MULTI_CAT_AXES: boolean
): string {
	const isBubble = isBubbleChart(chartObject.opts._type)
	const isScatter = isScatterChart(chartObject.opts._type)
	let count: number
	let uniqueCount: number
	// The leading entry is the blank the header row's label columns point at. Its two spellings are
	// not interchangeable: `<t/>` is the empty string, `<t xml:space="preserve"></t>` is a preserved
	// one, and next to character data that difference is content rather than layout.
	let blank = ''
	if (isBubble) {
		count = uniqueCount = intBubbleCols
	} else if (isScatter) {
		count = uniqueCount = data.length
	} else if (IS_MULTI_CAT_AXES) {
		let totCount = data.length + 1 // +1 for the blank entry at index 0
		dataLabels(data[0]).forEach((arrLabel) => (totCount += arrLabel.filter((label) => label && label !== '').length))
		count = uniqueCount = totCount
		blank = el('si', null, raw(voidEl('t')))
	} else {
		// series names + all labels of one series + number of label groups (data.labels.length) of one
		// series (i.e. how many times the blank string is used)
		count = data.length + dataLabels(data[0]).length * firstLabelGroup(data[0]).length + dataLabels(data[0]).length
		// series names + labels of one series + blank string (same for all label groups)
		uniqueCount = data.length + dataLabels(data[0]).length * firstLabelGroup(data[0]).length + 1
		blank = el('si', null, raw(el('t', { 'xml:space': 'preserve' })))
	}

	// Series names. A bubble series contributes both a name and a size column.
	const names = isBubble
		? data
				.map((objData, idx) =>
					idx === 0 ? sharedString('X-Axis') : sharedString(objData.name || `Y-Axis${idx}`) + sharedString(`Size${idx}`)
				)
				.join('')
		: data.map((objData) => sharedString((objData.name || ' ').replace('X-Axis', 'X-Values'))).join('')

	// Category labels, outermost group first — the order the sheet's label columns index into.
	// Blank entries are skipped: they share the single blank string at index 0.
	const labels =
		isBubble || isScatter
			? ''
			: dataLabels(data[0])
					.slice()
					.reverse()
					.map((labelsGroup) =>
						labelsGroup
							.filter((label) => label && label !== '')
							.map((label) => sharedString(label))
							.join('')
					)
					.join('')

	return XML_DECL + el('sst', { xmlns: SML_NS, count, uniqueCount }, [raw(blank), raw(names), raw(labels)]) + '\n'
}

/**
 * The embedded worksheet's extent: how many columns and how many rows (header included) the
 * sheet builder actually writes for this chart kind.
 *
 * Three sites derived these two numbers three ways, and one of them was wrong on its face:
 * `table1.xml`'s bubble range used the *column* count as its row count, so the gate deck's own
 * bubble chart said `ref="A1:C3"` in one part and `<dimension ref="A1:C5"/>` in the other. The
 * `<dimension>` in turn always used the first series' value count, which is not how the
 * category branch decides its rows.
 *
 * @param chartObject - the chart rel, for its normalized type
 * @param data - the chart's series
 * @param intBubbleCols - the bubble column count (`1 + 2 per Y series`), computed once by the caller
 */
function sheetExtent(
	chartObject: SlideRelChart,
	data: OptsChartDataInternal[],
	intBubbleCols: number
): { colCount: number; rowCount: number } {
	const first = data[0]
	const valueRows = dataValues(first).length + 1
	if (isBubbleChart(chartObject.opts._type)) return { colCount: intBubbleCols, rowCount: valueRows }
	if (isScatterChart(chartObject.opts._type)) return { colCount: data.length, rowCount: valueRows }
	// A category-less chartEx layout (a histogram feeds PowerPoint raw observations with no
	// labels) has no label groups, so the row count falls back to the longest value series --
	// the same rule the category branch of `buildXlsxSheet` applies when it writes those rows.
	const categoryRows = firstLabelGroup(first).length || Math.max(0, ...data.map((series) => dataValues(series).length))
	return { colCount: data.length + dataLabels(first).length, rowCount: categoryRows + 1 }
}

/**
 * Build the embedded workbook's `xl/tables/table1.xml` (the data table over the sheet range).
 */
function buildXlsxTable(chartObject: SlideRelChart, data: OptsChartDataInternal[], intBubbleCols: number): string {
	const labelCols = dataLabels(data[0]).length
	const { colCount, rowCount } = sheetExtent(chartObject, data, intBubbleCols)
	const ref = `A1:${getExcelColName(colCount)}${rowCount}`
	let columns: string
	if (isBubbleChart(chartObject.opts._type)) {
		let idxColLtr = 1
		columns = data
			.map((obj, idx) => {
				if (idx === 0) return voidEl('tableColumn', { id: idx + 1, name: 'X-Values' })
				// `?? ''`, like the category branch below: `name` is required on a `tableColumn`,
				// so an absent one omitted the attribute rather than writing an empty string.
				const nameCol = voidEl('tableColumn', { id: idx + idxColLtr, name: obj.name ?? '' })
				idxColLtr++
				return nameCol + voidEl('tableColumn', { id: idx + idxColLtr, name: `Size${idx}` })
			})
			.join('')
	} else if (isScatterChart(chartObject.opts._type)) {
		columns = data
			.map((_obj, idx) => voidEl('tableColumn', { id: idx + 1, name: `${idx === 0 ? 'X-Values' : 'Y-Value '}${idx}` }))
			.join('')
	} else {
		// The leading columns are the label groups; the series follow them.
		columns =
			dataLabels(data[0])
				.map((_labelsGroup, idx) => voidEl('tableColumn', { id: idx + 1, name: `Column${idx + 1}` }))
				.join('') +
			data.map((obj, idx) => voidEl('tableColumn', { id: idx + labelCols + 1, name: obj.name ?? '' })).join('')
	}
	return (
		XML_DECL +
		el('table', { xmlns: SML_NS, id: 1, name: 'Table1', displayName: 'Table1', ref, totalsRowShown: 0 }, [
			raw(el('tableColumns', { count: colCount }, raw(columns))),
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
function buildXlsxSheet(
	chartObject: SlideRelChart,
	data: OptsChartDataInternal[],
	intBubbleCols: number,
	IS_MULTI_CAT_AXES: boolean
): string {
	const isBubble = isBubbleChart(chartObject.opts._type)
	const isScatter = isScatterChart(chartObject.opts._type)
	const labelCols = dataLabels(data[0]).length
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

	const { colCount, rowCount } = sheetExtent(chartObject, data, intBubbleCols)
	let rows = ''

	if (isBubble) {
		/* EX: INPUT: `data`
				[
					{ name:'X-Axis'  , values:[10,11,12,13,14,15,16,17,18,19,20] },
					{ name:'Y-Axis 1', values:[ 1, 6, 7, 8, 9], sizes:[ 4, 5, 6, 7, 8] },
					{ name:'Y-Axis 2', values:[33,32,42,53,63], sizes:[11,12,13,14,15] }
				];
				*/
		/* EX: OUTPUT: bubbleChart Worksheet:
					-|----A-----|------B-----|------C-----|------D-----|------E-----|
					1| X-Values | Y-Values 1 | Y-Sizes 1  | Y-Values 2 | Y-Sizes 2  |
					2|    11    |     22     |      4     |     33     |      8     |
					-|----------|------------|------------|------------|------------|
				*/
		// Header row. Every column is a shared-string index; column A is the 'X-Axis' name.
		let header = ''
		for (let idx = 0; idx < intBubbleCols; idx++) header += cell(idx + 1, 1, idx, true)
		rows += sheetRow(1, intBubbleCols, header)

		// One row per X value; each series contributes a value column and a size column.
		dataValues(data[0]).forEach((val, idx) => {
			let cells = cell(1, idx + 2, val)
			let idxColLtr = 2
			for (let idy = 1; idy < data.length; idy++) {
				cells += cell(idxColLtr++, idx + 2, dataValues(data[idy])[idx] ?? '')
				cells += cell(idxColLtr++, idx + 2, dataSizes(data[idy])[idx] ?? '')
			}
			rows += sheetRow(idx + 2, intBubbleCols, cells)
		})
	} else if (isScatter) {
		/* EX: INPUT: `data`
					[
						{ name:'X-AxisA', values:[ 1, 2, 3, 4, 5] },
						{ name:'Y-AxisB', values:[ 2,22,42,52,62] },
						{ name:'Y-AxisC', values:[ 3,33,43,53,63] }
					];
				*/
		/* EX: OUTPUT: sheet1.xml:
					-|----A----|----B----|----C----|
					1| X-AxisA | Y-AxisB | Y-AxisC |
					2|    1    |    2    |    3    |
					-|---------|---------|---------|
				*/
		let header = ''
		for (let idx = 0; idx < data.length; idx++) header += cell(idx + 1, 1, idx, true)
		rows += sheetRow(1, data.length, header)

		dataValues(data[0]).forEach((val, idx) => {
			// The leading column is the X value; the rest are one Y series each.
			let cells = cell(1, idx + 2, val)
			for (let idy = 1; idy < data.length; idy++) {
				const yValue = dataValues(data[idy])[idx]
				cells += cell(idy + 1, idx + 2, yValue || yValue === 0 ? yValue : '')
			}
			rows += sheetRow(idx + 2, data.length, cells)
		})
	} else if (!IS_MULTI_CAT_AXES) {
		/* EX: INPUT: `data`
					[
						{ name:'Red', labels:['Jan..May-17'], values:[11,13,14,15,16] },
						{ name:'Amb', labels:['Jan..May-17'], values:[22, 6, 7, 8, 9] },
						{ name:'Grn', labels:['Jan..May-17'], values:[33,32,42,53,63] }
					];
				*/
		/* EX: OUTPUT: lineChart Worksheet:
					-|---A---|--B--|--C--|--D--|
					1|       | Red | Amb | Grn |
					2|Jan-17 |   11|   22|   33|
					3|Feb-17 |   55|   43|   70|
					4|Mar-17 |   56|  143|   99|
					5|Apr-17 |   65|    3|  120|
					6|May-17 |   75|   93|  170|
					-|-------|-----|-----|-----|
				*/
		// Header row: the label columns point at the blank shared string, then the series names.
		let header = ''
		for (let idx = 0; idx < labelCols; idx++) header += cell(idx + 1, 1, 0, true)
		for (let idx = 0; idx < data.length; idx++) header += cell(idx + 1 + labelCols, 1, idx + 1, true)
		rows += sheetRow(1, data.length + labelCols, header)

		// Normally one row per category; a category-less chartEx layout (a histogram feeds PowerPoint
		// raw observations with no labels) has no label groups, so the row count falls back to the
		// longest value series and the leading label columns are simply skipped — values land in A.
		const rowCount = firstLabelGroup(data[0]).length || Math.max(0, ...data.map((series) => dataValues(series).length))
		for (let idx = 0; idx < rowCount; idx++) {
			let cells = ''
			for (let idx2 = labelCols - 1; idx2 >= 0; idx2--)
				cells += cell(labelCols - idx2, idx + 2, data.length + idx + 1, true)
			for (let idy = 0; idy < data.length; idy++) {
				cells += cell(labelCols + idy + 1, idx + 2, dataValues(data[idy])[idx] ?? '')
			}
			rows += sheetRow(idx + 2, data.length + labelCols, cells)
		}
	} else {
		const TOT_SER = data.length
		const TOT_CAT = firstLabelGroup(data[0]).length
		// labels[0] is the leaf (inner) level; labels[TOT_LVL-1] is the outermost.
		// Reversed so that the outermost group occupies column A and the leaf occupies column TOT_LVL.
		const revLabelGroups = dataLabels(data[0]).slice().reverse()

		// Pre-build a map from (revLevelIdx, rowIdx) -> shared-string index.
		// SST layout: 0=blank, 1..TOT_SER=series names, then non-empty labels per
		// reversed level in appearance order.
		const ssLabelMap = new Map<string, number>()
		let ssIdx = TOT_SER + 1
		revLabelGroups.forEach((labelsGroup, revLevelIdx) => {
			labelsGroup.forEach((label, rowIdx) => {
				if (label && label !== '') ssLabelMap.set(`${revLevelIdx}:${rowIdx}`, ssIdx++)
			})
		})

		// Header row: label columns blank (index 0), series name columns use indices 1..TOT_SER
		let header = ''
		for (let col = 1; col <= labelCols; col++) header += cell(col, 1, 0, true)
		for (let ser = 0; ser < TOT_SER; ser++) header += cell(labelCols + ser + 1, 1, ser + 1, true)
		rows += sheetRow(1, TOT_SER + labelCols, header)

		// One data row per leaf category
		for (let idx = 0; idx < TOT_CAT; idx++) {
			// Label columns: column idy+1 holds revLabelGroups[idy]; emit only non-empty cells
			let cells = revLabelGroups
				.map((labelsGroup, idy) => {
					const colLabel = labelsGroup[idx]
					return colLabel && colLabel !== '' ? cell(idy + 1, idx + 2, ssLabelMap.get(`${idy}:${idx}`) ?? '', true) : ''
				})
				.join('')
			for (let idy = 0; idy < TOT_SER; idy++) {
				cells += cell(labelCols + idy + 1, idx + 2, dataValues(data[idy])[idx] ?? '')
			}
			rows += sheetRow(idx + 2, TOT_SER + labelCols, cells)
		}
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
				raw(voidEl('dimension', { ref: `A1:${getExcelColName(colCount)}${rowCount}` })),
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
