import JSZip from 'jszip'
import { build, defineRegressionSuite, assert, assertEqual } from '../helpers.js'

const LABELS = [
	['Gear', 'Berg', 'Motr', 'Swch', 'Plug', 'Cord', 'Pump', 'Leak', 'Seal'], // leaf (inner) — labels[0]
	['Mech', '', '', 'Elec', '', '', 'Hydr', '', ''], // outer group — labels[1]
]

const DATA = [
	{ name: 'West', labels: LABELS, values: [11, 8, 3, 0, 11, 3, 0, 0, 0] },
	{ name: 'Ctrl', labels: LABELS, values: [0, 11, 6, 19, 12, 5, 0, 0, 0] },
	{ name: 'East', labels: LABELS, values: [0, 3, 2, 0, 0, 0, 4, 3, 1] },
]

// Extract the embedded workbook XLSX from inside the PPTX and parse both XML files.
async function getWorkbookXml(buf) {
	const pptxZip = await JSZip.loadAsync(buf)
	const xlsxEntry = pptxZip.file('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')
	assert(xlsxEntry, 'embedded xlsx not found in pptx')
	const xlsxBuf = await xlsxEntry.async('arraybuffer')
	const xlsxZip = await JSZip.loadAsync(xlsxBuf)

	const sharedStringsXml = await xlsxZip.file('xl/sharedStrings.xml').async('string')
	const sheetXml = await xlsxZip.file('xl/worksheets/sheet1.xml').async('string')
	return { sharedStringsXml, sheetXml }
}

// Parse all <si> entries from sharedStrings.xml into a string array.
function parseSharedStrings(xml) {
	const entries = []
	for (const match of xml.matchAll(/<si><t[^>]*>([^<]*)<\/t><\/si>|<si><t\/><\/si>/g)) {
		entries.push(match[1] ?? '')
	}
	return entries
}

// Return the v-element value of a cell by address (e.g. "A2").
function cellValue(sheetXml, addr) {
	const re = new RegExp(`<c r="${addr}"[^>]*>(?:<v>([^<]*)<\\/v>)?`, 'i')
	const m = sheetXml.match(re)
	assert(m, `cell ${addr} not found in sheet XML`)
	return m[1] ?? null
}

// Return the t-attribute of a cell (e.g. "s" for shared-string).
function cellType(sheetXml, addr) {
	const re = new RegExp(`<c r="${addr}"([^>]*)>`)
	const m = sheetXml.match(re)
	assert(m, `cell ${addr} not found in sheet XML`)
	const t = m[1].match(/t="([^"]*)"/)
	return t ? t[1] : null
}

defineRegressionSuite('Multi-level category chart embedded workbook', 'upstream-pr-1330', [
	{
		name: 'shared-string indices for outer and leaf label cells are correct',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(p.charts.BAR, DATA, { x: 1, y: 1, w: 6, h: 4 })
			})
			const { sharedStringsXml, sheetXml } = await getWorkbookXml(buf)
			const ss = parseSharedStrings(sharedStringsXml)

			// SST layout: 0=blank, 1=West, 2=Ctrl, 3=East, 4=Mech, 5=Elec, 6=Hydr, 7=Gear, 8=Berg, ...
			assertEqual(ss[0], '', 'SST[0] should be blank')
			assertEqual(ss[1], 'West', 'SST[1] should be West')
			assertEqual(ss[2], 'Ctrl', 'SST[2] should be Ctrl')
			assertEqual(ss[3], 'East', 'SST[3] should be East')
			// reversed labels: outer group first (revLabelGroups[0] = labels[1])
			assertEqual(ss[4], 'Mech', 'SST[4] should be Mech (first outer group)')
			assertEqual(ss[5], 'Elec', 'SST[5] should be Elec (second outer group)')
			assertEqual(ss[6], 'Hydr', 'SST[6] should be Hydr (third outer group)')
			// then leaf labels (revLabelGroups[1] = labels[0])
			assertEqual(ss[7], 'Gear', 'SST[7] should be Gear (first leaf)')
			assertEqual(ss[8], 'Berg', 'SST[8] should be Berg (second leaf)')
			assertEqual(ss[15], 'Seal', 'SST[15] should be Seal (last leaf)')

			// Header row: label cols blank, series name cols use 1..3
			assertEqual(cellType(sheetXml, 'A1'), 's', 'A1 should be a string cell')
			assertEqual(cellValue(sheetXml, 'A1'), '0', 'A1 should reference blank (SST index 0)')
			assertEqual(cellValue(sheetXml, 'B1'), '0', 'B1 should reference blank (SST index 0)')
			assertEqual(cellValue(sheetXml, 'C1'), '1', 'C1 should reference West (SST index 1)')
			assertEqual(cellValue(sheetXml, 'D1'), '2', 'D1 should reference Ctrl (SST index 2)')
			assertEqual(cellValue(sheetXml, 'E1'), '3', 'E1 should reference East (SST index 3)')

			// Row 2: Mech (outer) in col A, Gear (leaf) in col B, values in C/D/E
			assertEqual(cellType(sheetXml, 'A2'), 's', 'A2 should be a string cell')
			assertEqual(cellValue(sheetXml, 'A2'), '4', 'A2 should reference Mech (SST index 4)')
			assertEqual(cellValue(sheetXml, 'B2'), '7', 'B2 should reference Gear (SST index 7)')
			assertEqual(cellValue(sheetXml, 'C2'), '11', 'C2 West value for Gear should be 11')
			assertEqual(cellValue(sheetXml, 'D2'), '0', 'D2 Ctrl value for Gear should be 0')
			assertEqual(cellValue(sheetXml, 'E2'), '0', 'E2 East value for Gear should be 0')

			// Row 3: outer blank (no A3 cell), Berg in B3, values in C3/D3/E3
			assert(!sheetXml.match(/<c r="A3"/), 'A3 should be absent (outer label is blank for Berg row)')
			assertEqual(cellValue(sheetXml, 'B3'), '8', 'B3 should reference Berg (SST index 8)')
			assertEqual(cellValue(sheetXml, 'C3'), '8', 'C3 West value for Berg should be 8')
			assertEqual(cellValue(sheetXml, 'D3'), '11', 'D3 Ctrl value for Berg should be 11')
			assertEqual(cellValue(sheetXml, 'E3'), '3', 'E3 East value for Berg should be 3')

			// Row 5 (idx=3): Elec (outer) in A5, Swch (leaf) in B5
			assertEqual(cellValue(sheetXml, 'A5'), '5', 'A5 should reference Elec (SST index 5)')
			assertEqual(cellValue(sheetXml, 'B5'), '10', 'B5 should reference Swch (SST index 10)')
		},
	},
])
