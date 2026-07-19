/**
 * PptxGenJS: Chart Embedded-Workbook Generation
 *
 * Builds the embedded `.xlsx` workbook that backs a chart's cached data — the data
 * source PowerPoint opens when a user edits the chart. `createExcelWorksheet` writes
 * the workbook plus the chart part + its `.rels` into the presentation package;
 * `buildEmbeddedWorksheet` / `buildChartRelsXml` are also reused by the read-side
 * injection path (`PptxGenJS.extractSlides`). Everything here is a pure string/bytes
 * builder — no I/O beyond the passed-in ZipWriter, no mutation of the presentation model.
 *
 * The chart's `chart.xml` DrawingML lives in `./chart-xml.ts`; the series↔worksheet-cell
 * mapping the two sides share lives in `./data-refs.ts`.
 */

import { ChartType, XML_DECL } from '../../core-enums.js'
import type { SlideRelChart, OptsChartDataInternal } from '../../core-interfaces.js'
import { encodeXmlEntities } from '../../gen-utils.js'
import { ZipWriter } from '../../zip.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { dataLabels, dataValues, dataSizes, firstLabelGroup, getExcelColName } from './data-refs.js'
import { makeXmlCharts } from './chart-xml.js'

const SCHEMA_BASE = 'http://schemas.openxmlformats.org/'
const PACKAGE_REL_NS = SCHEMA_BASE + 'package/2006/relationships'
const OFFICE_REL = SCHEMA_BASE + 'officeDocument/2006/relationships/'

function relationship(id: string, type: string, target: string): string {
	return voidEl('Relationship', { Id: id, Type: type, Target: target })
}

/** `<Relationships>` wrapper shared by the workbook's `.rels` parts (all flat, no indent). */
function relationships(rels: string[]): string {
	return el('Relationships', { xmlns: PACKAGE_REL_NS }, rels.map(raw))
}

/**
 * Build the chart's embedded Excel workbook as a standalone OPC package and
 * return its bytes — the data source PowerPoint opens when a user edits the
 * chart's data. Pure (no zip side effects), so both the package write path
 * ({@link createExcelWorksheet}) and the read-side injection path
 * (`PptxGenJS.extractSlides`) can reuse it.
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
			zipExcel.add(
				'[Content_Types].xml',
				XML_DECL +
					'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
					'  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
					'  <Default Extension="xml" ContentType="application/xml"/>' +
					'  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
					'  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
					'  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
					'  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
					'  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
					'  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' +
					'  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
					'  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
					'</Types>\n'
			)
			zipExcel.add(
				'_rels/.rels',
				XML_DECL +
					relationships([
						relationship('rId1', PACKAGE_REL_NS + '/metadata/core-properties', 'docProps/core.xml'),
						relationship('rId2', OFFICE_REL + 'extended-properties', 'docProps/app.xml'),
						relationship('rId3', OFFICE_REL + 'officeDocument', 'xl/workbook.xml'),
					]) +
					'\n'
			)
			zipExcel.add(
				'docProps/app.xml',
				XML_DECL +
					'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
					'<Application>Microsoft Macintosh Excel</Application>' +
					'<DocSecurity>0</DocSecurity>' +
					'<ScaleCrop>false</ScaleCrop>' +
					'<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>' +
					'<TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts>' +
					'<Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion>' +
					'</Properties>\n'
			)
			zipExcel.add(
				'docProps/core.xml',
				XML_DECL +
					'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
					'<dc:creator>PptxGenJS</dc:creator>' +
					'<cp:lastModifiedBy>PptxGenJS</cp:lastModifiedBy>' +
					'<dcterms:created xsi:type="dcterms:W3CDTF">' +
					new Date().toISOString() +
					'</dcterms:created>' +
					'<dcterms:modified xsi:type="dcterms:W3CDTF">' +
					new Date().toISOString() +
					'</dcterms:modified>' +
					'</cp:coreProperties>'
			)
			zipExcel.add(
				'xl/_rels/workbook.xml.rels',
				XML_DECL +
					// Ids are deliberately out of order (3/2/1/4) — that is how this part has
					// always been emitted, and rel order is byte-significant.
					relationships([
						relationship('rId3', OFFICE_REL + 'styles', 'styles.xml'),
						relationship('rId2', OFFICE_REL + 'theme', 'theme/theme1.xml'),
						relationship('rId1', OFFICE_REL + 'worksheet', 'worksheets/sheet1.xml'),
						relationship('rId4', OFFICE_REL + 'sharedStrings', 'sharedStrings.xml'),
					])
			)
			zipExcel.add(
				'xl/styles.xml',
				XML_DECL +
					'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="0" formatCode="General"/></numFmts><fonts count="4"><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="10"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="18"/><color indexed="8"/>' +
					'<name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><dxfs count="0"/><tableStyles count="0"/><colors><indexedColors><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ffff0000"/><rgbColor rgb="ff00ff00"/><rgbColor rgb="ff0000ff"/>' +
					'<rgbColor rgb="ffffff00"/><rgbColor rgb="ffff00ff"/><rgbColor rgb="ff00ffff"/><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ff878787"/><rgbColor rgb="fff9f9f9"/></indexedColors></colors></styleSheet>\n'
			)
			zipExcel.add(
				'xl/theme/theme1.xml',
				XML_DECL +
					'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>'
			)
			zipExcel.add(
				'xl/workbook.xml',
				XML_DECL +
					'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">' +
					'<fileVersion appName="xl" lastEdited="7" lowestEdited="6" rupBuild="10507"/>' +
					'<workbookPr/>' +
					'<bookViews><workbookView xWindow="0" yWindow="500" windowWidth="20960" windowHeight="15960"/></bookViews>' +
					'<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
					'<calcPr calcId="0" concurrentCalc="0"/>' +
					'</workbook>\n'
			)
			zipExcel.add(
				'xl/worksheets/_rels/sheet1.xml.rels',
				XML_DECL + relationships([relationship('rId1', OFFICE_REL + 'table', '../tables/table1.xml')]) + '\n'
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
	let strSharedStrings = XML_DECL
	if (chartObject.opts._type === ChartType.bubble || chartObject.opts._type === ChartType.bubble3d) {
		strSharedStrings += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${intBubbleCols}" uniqueCount="${intBubbleCols}">`
	} else if (chartObject.opts._type === ChartType.scatter) {
		strSharedStrings += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${data.length}" uniqueCount="${data.length}">`
	} else if (IS_MULTI_CAT_AXES) {
		let totCount = data.length + 1 // +1 for the blank entry at index 0
		dataLabels(data[0]).forEach((arrLabel) => (totCount += arrLabel.filter((label) => label && label !== '').length))
		strSharedStrings += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${totCount}" uniqueCount="${totCount}">`
		strSharedStrings += '<si><t/></si>'
	} else {
		// series names + all labels of one series + number of label groups (data.labels.length) of one series (i.e. how many times the blank string is used)
		const totCount =
			data.length + dataLabels(data[0]).length * firstLabelGroup(data[0]).length + dataLabels(data[0]).length
		// series names + labels of one series + blank string (same for all label groups)
		const unqCount = data.length + dataLabels(data[0]).length * firstLabelGroup(data[0]).length + 1
		// start `sst`
		strSharedStrings += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${totCount}" uniqueCount="${unqCount}">`
		// B: Add 'blank' for A1, B1, ..., of every label group inside data[n].labels
		strSharedStrings += '<si><t xml:space="preserve"></t></si>'
	}

	// C: Add `name`/Series
	if (chartObject.opts._type === ChartType.bubble || chartObject.opts._type === ChartType.bubble3d) {
		data.forEach((objData, idx) => {
			if (idx === 0) strSharedStrings += '<si><t>X-Axis</t></si>'
			else {
				strSharedStrings += `<si><t>${encodeXmlEntities(objData.name || `Y-Axis${idx}`)}</t></si>`
				strSharedStrings += `<si><t>${encodeXmlEntities(`Size${idx}`)}</t></si>`
			}
		})
	} else {
		data.forEach((objData) => {
			strSharedStrings += `<si><t>${encodeXmlEntities((objData.name || ' ').replace('X-Axis', 'X-Values'))}</t></si>`
		})
	}

	// D: Add `labels`/Categories
	if (
		chartObject.opts._type !== ChartType.bubble &&
		chartObject.opts._type !== ChartType.bubble3d &&
		chartObject.opts._type !== ChartType.scatter
	) {
		// Use forEach backwards & check for '' to support multi-cat axes
		dataLabels(data[0])
			.slice()
			.reverse()
			.forEach((labelsGroup) => {
				labelsGroup
					.filter((label) => label && label !== '')
					.forEach((label) => {
						strSharedStrings += `<si><t>${encodeXmlEntities(label)}</t></si>`
					})
			})
	}

	// DONE:
	strSharedStrings += '</sst>\n'
	return strSharedStrings
}

/**
 * Build the embedded workbook's `xl/tables/table1.xml` (the data table over the sheet range).
 */
function buildXlsxTable(chartObject: SlideRelChart, data: OptsChartDataInternal[], intBubbleCols: number): string {
	let strTableXml = XML_DECL
	if (chartObject.opts._type === ChartType.bubble || chartObject.opts._type === ChartType.bubble3d) {
		strTableXml += `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:${getExcelColName(intBubbleCols)}${intBubbleCols}" totalsRowShown="0">`
		strTableXml += `<tableColumns count="${intBubbleCols}">`
		let idxColLtr = 1
		data.forEach((obj, idx) => {
			if (idx === 0) {
				strTableXml += `<tableColumn id="${idx + 1}" name="X-Values"/>`
			} else {
				strTableXml += `<tableColumn id="${idx + idxColLtr}" name="${obj.name}"/>`
				idxColLtr++
				strTableXml += `<tableColumn id="${idx + idxColLtr}" name="Size${idx}"/>`
			}
		})
	} else if (chartObject.opts._type === ChartType.scatter) {
		strTableXml += `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:${getExcelColName(data.length)}${dataValues(data[0]).length + 1}" totalsRowShown="0">`
		strTableXml += `<tableColumns count="${data.length}">`
		data.forEach((_obj, idx) => {
			strTableXml += `<tableColumn id="${idx + 1}" name="${idx === 0 ? 'X-Values' : 'Y-Value '}${idx}"/>`
		})
	} else {
		strTableXml += `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:${getExcelColName(data.length + dataLabels(data[0]).length)}${firstLabelGroup(data[0]).length + 1}" totalsRowShown="0">`
		strTableXml += `<tableColumns count="${data.length + dataLabels(data[0]).length}">`
		dataLabels(data[0]).forEach((_labelsGroup, idx) => {
			strTableXml += `<tableColumn id="${idx + 1}" name="Column${idx + 1}"/>`
		})
		data.forEach((obj, idx) => {
			strTableXml += `<tableColumn id="${idx + dataLabels(data[0]).length + 1}" name="${encodeXmlEntities(obj.name ?? '')}"/>`
		})
	}
	strTableXml += '</tableColumns>'
	strTableXml += '<tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>'
	strTableXml += '</table>'
	return strTableXml
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
	let strSheetXml = XML_DECL
	strSheetXml +=
		'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'

	if (chartObject.opts._type === ChartType.bubble || chartObject.opts._type === ChartType.bubble3d) {
		strSheetXml += `<dimension ref="A1:${getExcelColName(intBubbleCols)}${dataValues(data[0]).length + 1}"/>`
	} else if (chartObject.opts._type === ChartType.scatter) {
		strSheetXml += `<dimension ref="A1:${getExcelColName(data.length)}${dataValues(data[0]).length + 1}"/>`
	} else {
		strSheetXml += `<dimension ref="A1:${getExcelColName(data.length + dataLabels(data[0]).length)}${dataValues(data[0]).length + 1}"/>`
	}

	strSheetXml +=
		'<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B1" sqref="B1"/></sheetView></sheetViews>'
	strSheetXml += '<sheetFormatPr baseColWidth="10" defaultRowHeight="16"/>'
	if (chartObject.opts._type === ChartType.bubble || chartObject.opts._type === ChartType.bubble3d) {
		// UNUSED: strSheetXml += `<cols><col min="1" max="${data.length}" width="11" customWidth="1" /></cols>`

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
		strSheetXml += '<sheetData>'

		// A: Create header row first (NOTE: Start at index=1 as headers cols start with 'B')
		strSheetXml += `<row r="1" spans="1:${intBubbleCols}">`
		strSheetXml += '<c r="A1" t="s"><v>0</v></c>'
		for (let idx = 1; idx < intBubbleCols; idx++) {
			strSheetXml += `<c r="${getExcelColName(idx + 1)}1" t="s"><v>${idx}</v></c>` // NOTE: add `t="s"` for label cols!
		}
		strSheetXml += '</row>'

		// B: Add row for each X-Axis value (Y-Axis* value is optional)
		dataValues(data[0]).forEach((val, idx) => {
			// Leading col is reserved for the 'X-Axis' value, so hard-code it, then loop over col values
			strSheetXml += `<row r="${idx + 2}" spans="1:${intBubbleCols}">`
			strSheetXml += `<c r="A${idx + 2}"><v>${val}</v></c>`
			// Add Y-Axis 1->N (idy=0 = Xaxis)
			let idxColLtr = 2
			for (let idy = 1; idy < data.length; idy++) {
				// y-value
				strSheetXml += `<c r="${getExcelColName(idxColLtr)}${idx + 2}"><v>${dataValues(data[idy])[idx] ?? ''}</v></c>`
				idxColLtr++
				// y-size
				strSheetXml += `<c r="${getExcelColName(idxColLtr)}${idx + 2}"><v>${dataSizes(data[idy])[idx] ?? ''}</v></c>`
				idxColLtr++
			}
			strSheetXml += '</row>'
		})
	} else if (chartObject.opts._type === ChartType.scatter) {
		/* UNUSED:
					strSheetXml += '<cols>'
					strSheetXml += '<col min="1" max="' + data.length + '" width="11" customWidth="1" />'
					//data.forEach((obj,idx)=>{ strSheetXml += '<col min="'+(idx+1)+'" max="'+(idx+1)+'" width="11" customWidth="1" />' });
					strSheetXml += '</cols>'
				*/
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
		strSheetXml += '<sheetData>'

		// A: Create header row first (every `name` row provided)
		strSheetXml += `<row r="1" spans="1:${data.length}">`
		for (let idx = 0; idx < data.length; idx++) {
			strSheetXml += `<c r="${getExcelColName(idx + 1)}1" t="s"><v>${idx}</v></c>` // NOTE: add `t="s"` for label cols!
		}
		strSheetXml += '</row>'

		// B: Add row for each X-Axis value (Y-Axis* value is optional)
		dataValues(data[0]).forEach((val, idx) => {
			// Leading col is reserved for the 'X-Axis' value, so hard-code it, then loop over col values
			strSheetXml += `<row r="${idx + 2}" spans="1:${data.length}">`
			strSheetXml += `<c r="A${idx + 2}"><v>${val}</v></c>`
			// Add Y-Axis 1->N
			for (let idy = 1; idy < data.length; idy++) {
				strSheetXml += `<c r="${getExcelColName(idy + 1)}${idx + 2}"><v>${
					dataValues(data[idy])[idx] || dataValues(data[idy])[idx] === 0 ? dataValues(data[idy])[idx] : ''
				}</v></c>`
			}
			strSheetXml += '</row>'
		})
	} else {
		strSheetXml += '<sheetData>'

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

		if (!IS_MULTI_CAT_AXES) {
			// A: Create header row first
			strSheetXml += `<row r="1" spans="1:${data.length + dataLabels(data[0]).length}">`
			dataLabels(data[0]).forEach((_labelsGroup, idx) => {
				strSheetXml += `<c r="${getExcelColName(idx + 1)}1" t="s"><v>0</v></c>`
			})
			for (let idx = 0; idx < data.length; idx++) {
				strSheetXml += `<c r="${getExcelColName(idx + 1 + dataLabels(data[0]).length)}1" t="s"><v>${idx + 1}</v></c>` // NOTE: use `t="s"` for label cols!
			}
			strSheetXml += '</row>'

			// B: Add data row(s) for each category
			firstLabelGroup(data[0]).forEach((_cat, idx) => {
				strSheetXml += `<row r="${idx + 2}" spans="1:${data.length + dataLabels(data[0]).length}">`
				// Leading cols are reserved for the label groups
				for (let idx2 = dataLabels(data[0]).length - 1; idx2 >= 0; idx2--) {
					strSheetXml += `<c r="${getExcelColName(dataLabels(data[0]).length - idx2)}${idx + 2}" t="s">`
					strSheetXml += `<v>${data.length + idx + 1}</v>`
					strSheetXml += '</c>'
				}
				for (let idy = 0; idy < data.length; idy++) {
					strSheetXml += `<c r="${getExcelColName(dataLabels(data[0]).length + idy + 1)}${idx + 2}"><v>${dataValues(data[idy])[idx] ?? ''}</v></c>`
				}
				strSheetXml += '</row>'
			})
		} else {
			const TOT_SER = data.length
			const TOT_CAT = firstLabelGroup(data[0]).length
			const TOT_LVL = dataLabels(data[0]).length
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
			strSheetXml += `<row r="1" spans="1:${TOT_SER + TOT_LVL}">`
			for (let col = 1; col <= TOT_LVL; col++) {
				strSheetXml += `<c r="${getExcelColName(col)}1" t="s"><v>0</v></c>`
			}
			for (let ser = 0; ser < TOT_SER; ser++) {
				strSheetXml += `<c r="${getExcelColName(TOT_LVL + ser + 1)}1" t="s"><v>${ser + 1}</v></c>`
			}
			strSheetXml += '</row>'

			// One data row per leaf category
			for (let idx = 0; idx < TOT_CAT; idx++) {
				strSheetXml += `<row r="${idx + 2}" spans="1:${TOT_SER + TOT_LVL}">`
				// Label columns: column idy+1 holds revLabelGroups[idy]; emit only non-empty cells
				revLabelGroups.forEach((labelsGroup, idy) => {
					const colLabel = labelsGroup[idx]
					if (colLabel && colLabel !== '') {
						strSheetXml += `<c r="${getExcelColName(idy + 1)}${idx + 2}" t="s"><v>${ssLabelMap.get(`${idy}:${idx}`)}</v></c>`
					}
				})
				// Data columns
				for (let idy = 0; idy < TOT_SER; idy++) {
					strSheetXml += `<c r="${getExcelColName(TOT_LVL + idy + 1)}${idx + 2}"><v>${dataValues(data[idy])[idx] ?? ''}</v></c>`
				}
				strSheetXml += '</row>'
			}
		}
	}
	strSheetXml += '</sheetData>'

	strSheetXml += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
	// Link the `table1.xml` file to define an actual Table in Excel
	// NOTE: Intentionally no `<tableParts>` here. A tablePart only works for scatter charts;
	// every other chart type reports a "cannot find linked file" error. The chart data can be
	// edited / range-selected without it, so it is deliberately never emitted.
	strSheetXml += '</worksheet>\n'
	return strSheetXml
}

/**
 * Build the standalone `.rels` for a chart part: a single `rId1` relationship to
 * the chart's embedded workbook (`Target`). Shared by the package write path and
 * the read-side injection path, which pass different (relative) embedding targets.
 * @param {string} embeddingTarget - the workbook target, relative to the chart part
 * @return {string} the chart part's `.rels` XML
 */
export function buildChartRelsXml(embeddingTarget: string): string {
	// `voidEl` escapes the Target. The one in-tree caller passes an internally built
	// `../embeddings/Microsoft_Excel_WorksheetN.xlsx`, so that is a no-op on bytes;
	// it matters only for the read-side injection path, which supplies its own target.
	return XML_DECL + relationships([relationship('rId1', OFFICE_REL + 'package', embeddingTarget)])
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

	// 2: Create the chart.xml and rel files
	zip.add(
		'ppt/charts/_rels/' + chartObject.fileName + '.rels',
		buildChartRelsXml(`../embeddings/Microsoft_Excel_Worksheet${chartObject.globalId}.xlsx`)
	)
	zip.add(`ppt/charts/${chartObject.fileName}`, makeXmlCharts(chartObject))

	return ''
}
