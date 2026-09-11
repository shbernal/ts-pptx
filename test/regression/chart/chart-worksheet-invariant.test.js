import { DOMParser } from '@xmldom/xmldom'
import JSZip from 'jszip'
import { ChartType } from '../../../dist/node.js'
import { defineRegressionSuite, build, listEntries, assert } from '../../helpers.js'

// Every formula a chart part carries names cells in its own embedded workbook, and the cache
// beside it is a copy of what those cells hold. The workbook writer and the chart emitters work
// the layout out on their own, so this checks the two against each other rather than either one
// against a hard-coded reference: resolve each `<c:f>` and `<cx:f>` through `sheet1.xml` and
// `sharedStrings.xml`, and compare what comes back with the cache point by point.
//
// A chart whose formulas are wrong still paints, because PowerPoint draws from the cache. The
// damage shows on "Edit Data", which opens the workbook the formulas describe.

const parser = new DOMParser()

/** The children of `node` with the qualified name `name`. */
const kids = (node, name) => Array.from(node.childNodes).filter((child) => child.nodeName === name)
const kid = (node, name) => kids(node, name)[0]

/** 1-based column index for a column name. */
function columnIndex(name) {
	let n = 0
	for (const ch of name) n = n * 26 + (ch.charCodeAt(0) - 64)
	return n
}

/** `Sheet1!$A$2:$C$4` or `Sheet1!$B$1` as column and row bounds. */
function parseRef(f) {
	const m = /^Sheet1!\$([A-Z]+)\$(\d+)(?::\$([A-Z]+)\$(\d+))?$/.exec(f)
	assert(m, `unparseable reference ${JSON.stringify(f)}`)
	const c1 = columnIndex(m[1])
	const r1 = Number(m[2])
	return { c1, r1, c2: m[3] ? columnIndex(m[3]) : c1, r2: m[4] ? Number(m[4]) : r1 }
}

/** The embedded workbook as a cell lookup: `(col, row) => string`, blank for an absent cell. */
async function readWorkbook(zip) {
	const name = listEntries(zip).find((entry) => /^ppt\/embeddings\/.*\.xlsx$/.test(entry))
	assert(name, 'expected an embedded workbook')
	const xlsx = await JSZip.loadAsync(await zip.file(name).async('arraybuffer'))
	const sst = parser.parseFromString(await xlsx.file('xl/sharedStrings.xml').async('string'), 'text/xml')
	const strings = Array.from(sst.getElementsByTagName('si')).map((si) => si.textContent ?? '')
	const sheet = parser.parseFromString(await xlsx.file('xl/worksheets/sheet1.xml').async('string'), 'text/xml')
	const cells = new Map()
	for (const c of Array.from(sheet.getElementsByTagName('c'))) {
		const v = c.getElementsByTagName('v')[0]?.textContent ?? ''
		if (c.getAttribute('t') !== 's') {
			cells.set(c.getAttribute('r'), v)
			continue
		}
		const idx = Number(v)
		cells.set(c.getAttribute('r'), idx < strings.length ? strings[idx] : `<shared string ${idx} of ${strings.length}>`)
	}
	return (col, row) => {
		let name = ''
		for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name
		return cells.get(`${name}${row}`) ?? ''
	}
}

/** Two values agree when they are the same number, or the same text. Blank and absent are one thing. */
function same(cell, cached) {
	if (cell === cached) return true
	if (cell === '' || cached === '') return false
	return Number.isFinite(Number(cell)) && Number(cell) === Number(cached)
}

/** The cells a one-dimensional range spans, in order. */
function rangeCells(ref) {
	const cells = []
	for (let r = ref.r1; r <= ref.r2; r++) for (let c = ref.c1; c <= ref.c2; c++) cells.push([c, r])
	return cells
}

/** One level's points as a sparse array of `idx -> value`. */
function points(lvl, ptName, valueOf) {
	const out = []
	for (const pt of kids(lvl, ptName)) out[Number(pt.getAttribute('idx'))] = valueOf(pt)
	return out
}

/**
 * Compare one level of cached points with a run of cells. `ptCount` is the length of the range the
 * formula spans, so it has to equal the number of cells. Points at or past it are left to the
 * point-count checks, which are a separate question from where the range points.
 */
function compareLevel(problems, where, f, cells, ptCount, pts, read) {
	if (ptCount !== cells.length) problems.push(`${where}: ${f} spans ${cells.length} cells, the cache says ${ptCount}`)
	cells.forEach(([c, r], idx) => {
		const cell = read(c, r)
		const cached = pts[idx] ?? ''
		if (!same(cell, cached))
			problems.push(`${where}: ${f} cell ${idx} reads ${JSON.stringify(cell)}, cached ${JSON.stringify(cached)}`)
	})
}

/** A classic `<c:strRef>` / `<c:numRef>`. */
function checkRef(problems, node, read) {
	const f = kid(node, 'c:f').textContent
	const ref = parseRef(f)
	const cache = kid(node, 'c:strCache') ?? kid(node, 'c:numCache')
	const ptCount = Number(kid(cache, 'c:ptCount')?.getAttribute('val'))
	if (ref.c1 !== ref.c2 && ref.r1 !== ref.r2) problems.push(`${node.nodeName}: ${f} is two-dimensional`)
	const pts = points(cache, 'c:pt', (pt) => kid(pt, 'c:v')?.textContent ?? '')
	compareLevel(problems, node.nodeName, f, rangeCells(ref), ptCount, pts, read)
}

/**
 * Levels run leaf first, while the workbook puts the outermost level in the range's first column,
 * so level `i` is column `c2 - i`.
 */
function checkLevels(problems, where, f, levels, ptCountOf, ptName, valueOf, read) {
	const ref = parseRef(f)
	const columns = ref.c2 - ref.c1 + 1
	if (columns !== levels.length) problems.push(`${where}: ${f} spans ${columns} columns for ${levels.length} levels`)
	levels.forEach((lvl, i) => {
		const col = ref.c2 - i
		const cells = []
		for (let r = ref.r1; r <= ref.r2; r++) cells.push([col, r])
		compareLevel(problems, `${where} level ${i}`, f, cells, ptCountOf(lvl), points(lvl, ptName, valueOf), read)
	})
}

/** Every formula in a classic chart part, against the workbook. */
function checkClassic(xml, read) {
	const problems = []
	const doc = parser.parseFromString(xml, 'text/xml')
	for (const tag of ['c:strRef', 'c:numRef']) {
		for (const node of Array.from(doc.getElementsByTagName(tag))) checkRef(problems, node, read)
	}
	for (const node of Array.from(doc.getElementsByTagName('c:multiLvlStrRef'))) {
		const cache = kid(node, 'c:multiLvlStrCache')
		const ptCount = Number(kid(cache, 'c:ptCount').getAttribute('val'))
		checkLevels(
			problems,
			'c:multiLvlStrRef',
			kid(node, 'c:f').textContent,
			kids(cache, 'c:lvl'),
			() => ptCount,
			'c:pt',
			(pt) => kid(pt, 'c:v')?.textContent ?? '',
			read
		)
	}
	// A series index is how PowerPoint tells the series of one chart apart, across every plot group.
	for (const name of ['c:idx', 'c:order']) {
		const seen = Array.from(doc.getElementsByTagName('c:ser')).map((ser) => kid(ser, name).getAttribute('val'))
		if (new Set(seen).size !== seen.length) problems.push(`${name} repeats across series: ${seen.join(', ')}`)
	}
	return problems
}

/** Every formula in a chartEx part, against the workbook. */
function checkChartEx(xml, read) {
	const problems = []
	const doc = parser.parseFromString(xml, 'text/xml')
	for (const tag of ['cx:strDim', 'cx:numDim']) {
		for (const node of Array.from(doc.getElementsByTagName(tag))) {
			checkLevels(
				problems,
				tag,
				kid(node, 'cx:f').textContent,
				kids(node, 'cx:lvl'),
				(lvl) => Number(lvl.getAttribute('ptCount')),
				'cx:pt',
				(pt) => pt.textContent ?? '',
				read
			)
		}
	}
	for (const node of Array.from(doc.getElementsByTagName('cx:txData'))) {
		const f = kid(node, 'cx:f').textContent
		const ref = parseRef(f)
		compareLevel(problems, 'cx:txData', f, rangeCells(ref), 1, [kid(node, 'cx:v').textContent], read)
	}
	return problems
}

/** Build one chart and return every disagreement between its formulas and its workbook. */
async function problemsFor(addChart) {
	const { zip } = await build((p) => addChart(p.addSlide()))
	const read = await readWorkbook(zip)
	const classic = listEntries(zip).find((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name))
	if (classic) return checkClassic(await zip.file(classic).async('string'), read)
	const chartEx = listEntries(zip).find((name) => /^ppt\/charts\/chartEx\d+\.xml$/.test(name))
	assert(chartEx, 'expected a chart part')
	return checkChartEx(await zip.file(chartEx).async('string'), read)
}

const FRAME = { x: 1, y: 1, w: 6, h: 4 }
const LABELS = ['Jan', 'Feb', 'Mar']

/** `[name, addChart]`, every shape the matrix covers. */
const MATRIX = [
	[
		'bar',
		(s) =>
			s.addChart(
				[
					{ name: 'A', labels: LABELS, values: [1, 2, 3] },
					{ name: 'B', labels: LABELS, values: [4, 5, 6] },
				],
				{ type: ChartType.bar, ...FRAME }
			),
	],
	[
		'line labelled on the first series only',
		(s) =>
			s.addChart(
				[
					{ name: 'A', labels: LABELS, values: [1, 2, 3] },
					{ name: 'B', values: [4, 5, 6] },
				],
				{ type: ChartType.line, ...FRAME }
			),
	],
	[
		'bar with multi-level labels',
		(s) =>
			s.addChart(
				[
					{ name: 'A', labels: [LABELS, ['Q1', '', '']], values: [1, 2, 3] },
					{ name: 'B', labels: [LABELS, ['Q1', '', '']], values: [4, 5, 6] },
				],
				{ type: ChartType.bar, ...FRAME }
			),
	],
	['pie', (s) => s.addChart([{ name: 'P', labels: LABELS, values: [1, 2, 3] }], { type: ChartType.pie, ...FRAME })],
	[
		'doughnut',
		(s) => s.addChart([{ name: 'P', labels: LABELS, values: [1, 2, 3] }], { type: ChartType.doughnut, ...FRAME }),
	],
	['unlabelled pie', (s) => s.addChart([{ name: 'P', values: [1, 2, 3] }], { type: ChartType.pie, ...FRAME })],
	[
		'pie with two label levels',
		(s) =>
			s.addChart([{ name: 'P', labels: [LABELS, ['Q1', '', '']], values: [1, 2, 3] }], {
				type: ChartType.pie,
				...FRAME,
			}),
	],
	[
		'scatter',
		(s) =>
			s.addChart(
				[
					{ name: 'X', values: [10, 20, 30] },
					{ name: 'Y1', values: [5, 6, 7] },
					{ name: 'Y2', values: [8, 9, 10] },
				],
				{ type: ChartType.scatter, ...FRAME }
			),
	],
	[
		'bubble',
		(s) =>
			s.addChart(
				[
					{ name: 'X', values: [10, 20, 30] },
					{ name: 'Y1', values: [5, 6, 7], sizes: [1, 2, 3] },
					{ name: 'Y2', values: [8, 9, 10], sizes: [4, 5, 6] },
				],
				{ type: ChartType.bubble, ...FRAME }
			),
	],
	[
		'bar and line combo',
		(s) =>
			s.addChart(
				[
					{ type: ChartType.bar, data: [{ name: 'A', labels: LABELS, values: [1, 2, 3] }], options: {} },
					{
						type: ChartType.line,
						data: [{ name: 'B', labels: LABELS, values: [4, 5, 6] }],
						options: { secondaryValAxis: true, secondaryCatAxis: true },
					},
				],
				FRAME
			),
	],
	[
		'bar and scatter combo',
		(s) =>
			s.addChart(
				[
					{ type: ChartType.bar, data: [{ name: 'A', labels: LABELS, values: [1, 2, 3] }], options: {} },
					{
						type: ChartType.scatter,
						data: [
							{ name: 'X', values: [10, 20, 30] },
							{ name: 'Y', values: [5, 6, 7] },
						],
						options: { secondaryValAxis: true, secondaryCatAxis: true },
					},
				],
				FRAME
			),
	],
	[
		'bar with a blank label',
		(s) =>
			s.addChart(
				[
					{ name: 'A', labels: ['A', '', 'C'], values: [1, 2, 3] },
					{ name: 'B', labels: ['A', '', 'C'], values: [4, 5, 6] },
				],
				{ type: ChartType.bar, ...FRAME }
			),
	],
	[
		'stock',
		(s) =>
			s.addChart(
				[
					{ name: 'High', labels: LABELS, values: [10, 12, 11] },
					{ name: 'Low', labels: LABELS, values: [7, 8, 6] },
					{ name: 'Close', labels: LABELS, values: [9, 11, 7] },
				],
				{ type: ChartType.stock, ...FRAME }
			),
	],
	[
		'surface',
		(s) =>
			s.addChart(
				[
					{ name: 'A', labels: LABELS, values: [1, 2, 3] },
					{ name: 'B', labels: LABELS, values: [4, 5, 6] },
				],
				{ type: ChartType.surface, ...FRAME }
			),
	],
	[
		'waterfall',
		(s) => s.addChart([{ name: 'W', labels: LABELS, values: [1, -2, 3] }], { type: ChartType.waterfall, ...FRAME }),
	],
	[
		'waterfall with more values than labels',
		(s) =>
			s.addChart([{ name: 'W', labels: ['a', 'b'], values: [1, 2, 3, 4] }], { type: ChartType.waterfall, ...FRAME }),
	],
	[
		'treemap',
		(s) =>
			s.addChart(
				[
					{
						name: 'T',
						labels: [
							['leaf1', 'leaf2', 'leaf3'],
							['branch1', 'branch1', 'branch2'],
						],
						values: [1, 2, 3],
					},
				],
				{ type: ChartType.treemap, ...FRAME }
			),
	],
	['histogram', (s) => s.addChart([{ name: 'H', values: [1, 2, 2, 3, 5] }], { type: ChartType.histogram, ...FRAME })],
]

/** The shapes that fail today, each named for what is wrong with it. */
const BROKEN = new Set([
	'unlabelled pie',
	'pie with two label levels',
	'bar and scatter combo',
	'bar with a blank label',
	'waterfall with more values than labels',
])

defineRegressionSuite('Chart formulas resolve to their cache through the embedded workbook', [
	...MATRIX.map(([name, addChart]) => ({
		name,
		fails: BROKEN.has(name),
		fn: async () => {
			const problems = await problemsFor(addChart)
			assert(problems.length === 0, `${name}:\n  ${problems.join('\n  ')}`)
		},
	})),
	{
		// The check has to be able to fail: a resolver that read every cell as blank, or a comparison
		// that skipped every point, would pass the matrix above. Point one series at its neighbour's
		// column and it has to say so.
		name: 'a formula moved to the wrong column is reported',
		fn: async () => {
			const { zip } = await build((p) =>
				p.addSlide().addChart(
					[
						{ name: 'A', labels: LABELS, values: [1, 2, 3] },
						{ name: 'B', labels: LABELS, values: [4, 5, 6] },
					],
					{ type: ChartType.bar, ...FRAME }
				)
			)
			const read = await readWorkbook(zip)
			const xml = await zip.file('ppt/charts/chart1.xml').async('string')
			assert(checkClassic(xml, read).length === 0, 'the unperturbed chart is clean')
			const moved = xml.replace('<c:f>Sheet1!$C$2:$C$4</c:f>', '<c:f>Sheet1!$B$2:$B$4</c:f>')
			assert(moved !== xml, 'the perturbation applies')
			const problems = checkClassic(moved, read)
			assert(problems.length === 3, `three misread points; got ${JSON.stringify(problems)}`)
		},
	},
])
