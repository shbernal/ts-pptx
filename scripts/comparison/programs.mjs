/**
 * The bundle corpus: whole decks a consumer might actually write, measured as bundles.
 *
 * `scripts/comparison/probes.mjs` asks what each library *can emit*, one construct at a
 * time. This asks a different question with the same discipline: what a consumer's build
 * weighs once they use the library for something. A hello world answers that for the
 * smallest program anyone writes and for no other, and on a tree-shaken bundle the smallest
 * program is exactly where two libraries look most alike -- most of what either one costs
 * has been shaken out. So the corpus climbs: text, tables, charts, then a deck that uses
 * everything both libraries share.
 *
 * Three rules, all of them the probe corpus's rules for the same reasons:
 *
 *   - **Both arms build the same deck.** A program only one library can build measures two
 *     different pieces of work and calls the difference a size. Every program here is
 *     restricted to the constructs the shared baseline shows both libraries emitting.
 *   - **Each arm is written in its own library's idiom**, not transcribed from the other.
 *     Where the two come out character-identical anyway the page says so and prints one
 *     block.
 *   - **The programs are executed before they are bundled** (`scripts/comparison/hygiene.mjs`).
 *     esbuild will happily bundle a call that does not exist, and a misspelled method is a
 *     smaller bundle rather than an error: the tree-shaker simply keeps less. Running each
 *     program first is what stops a typo from being published as a saving.
 *
 * The result of `write` is passed to `console.log` on purpose. An export whose value is
 * discarded is dead code a minifier can prove away, and the row would become a measurement
 * of how well each library annotates side effects rather than of what a deck costs.
 */
import { corpusData } from './corpus-data.mjs'
import { functionBody, literal, renderSource } from './source.mjs'

/** A 1x1 transparent PNG, small enough that the image path is measured and not the payload. */
const PNG_1PX_B64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
/** ts-pptx takes a whole `data:` URL. */
const PNG_1PX_URL = 'data:image/png;base64,' + PNG_1PX_B64
/** pptxgenjs takes the MIME and payload without the `data:` scheme, as its own docs show. */
const PNG_1PX_BARE = 'image/png;base64,' + PNG_1PX_B64

/** Table content, cell objects rather than bare strings, as a real deck writes it. */
const REGION_ROWS = [
	[{ text: 'Region' }, { text: 'Revenue' }, { text: 'Growth' }],
	[{ text: 'North America' }, { text: '24.9' }, { text: '14.2%' }],
	[{ text: 'EMEA' }, { text: '12.6' }, { text: '16.8%' }],
	[{ text: 'APAC' }, { text: '6.8' }, { text: '9.4%' }],
]

/** Chart data in the shape both libraries take unchanged, one series and three. */
const BAR_DATA = [{ name: 'Revenue', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [12, 19, 7, 24] }]
const LINE_DATA = [
	{ name: 'Net retention', labels: ['Jul', 'Aug', 'Sep'], values: [108, 111, 114] },
	{ name: 'Gross retention', labels: ['Jul', 'Aug', 'Sep'], values: [94, 94, 96] },
]
const PIE_DATA = [{ name: 'Revenue mix', labels: ['Platform', 'Services', 'Licensing'], values: [25.8, 13.4, 7.2] }]

/**
 * Every value a program names, for the preamble {@link programSource} prints above it.
 *
 * `PNG_1PX_B64` is absent for the reason it is absent from the probe corpus: no program
 * names it, and the two names below are how the two libraries' image arms differ.
 *
 * Through {@link corpusData} for the reason given there, which bites harder here than it
 * does on the probes: these programs are run before they are rendered *and* before they are
 * bundled, so without it a library's edits would reach the page and the compiled entry
 * alike. {@link resetProgramData} is called between arms.
 */
const { constants: PROGRAM_CONSTANTS, reset: resetProgramData } = corpusData({
	PNG_1PX_URL,
	PNG_1PX_BARE,
	REGION_ROWS,
	BAR_DATA,
	LINE_DATA,
	PIE_DATA,
})

export { resetProgramData }

/**
 * The frame every program is compiled inside, per library.
 *
 * Stated once here and stated once on the page, rather than repeated at the top of five
 * programs where a reader would have to check five times that it did not change.
 * @type {Record<string, {import: string, construct: string}>}
 */
export const FRAME = {
	'ts-pptx': { import: "import TsPptx from 'pptx-ts'", construct: 'const pres = new TsPptx()' },
	pptxgenjs: { import: "import PptxGenJS from 'pptxgenjs'", construct: 'const pres = new PptxGenJS()' },
}

/** The last line of every program, closing the frame. */
export const EXPORT_LINE = "console.log(await pres.write({ outputType: 'arraybuffer' }))"

/**
 * One consumer program, built by both libraries and bundled twice.
 * @typedef {object} Program
 * @property {string} id - kebab-case, stable; the bundle table's row key
 * @property {string} label - how the row reads on the page
 * @property {string} what - one line saying what the deck contains
 * @property {Record<string, (pres: any) => unknown>} build - per subject; every program has both arms
 */

/** @type {Program[]} */
export const PROGRAMS = [
	{
		id: 'hello-world',
		label: 'Hello world',
		what: 'One slide with one text box.',
		build: {
			'ts-pptx': (pres) => {
				pres.addSlide().addText('hello', { x: 1, y: 1, w: 4, h: 1 })
			},
			pptxgenjs: (pres) => {
				pres.addSlide().addText('hello', { x: 1, y: 1, w: 4, h: 1 })
			},
		},
	},
	{
		id: 'text-deck',
		label: 'Text deck',
		what: 'A defined master, two sections, formatted and bulleted text, a hyperlink, a slide background and speaker notes.',
		build: {
			'ts-pptx': (pres) => {
				pres.defineSlideMaster({
					title: 'NARRATIVE',
					objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.5, w: 8.8, h: 1 } } }],
				})
				pres.addSection({ title: 'Findings' })
				pres.addSection({ title: 'Next steps' })

				const findings = pres.addSlide({ masterTitle: 'NARRATIVE', sectionTitle: 'Findings' })
				findings.background = { color: 'F2F2F2' }
				findings.addText('What we found', { placeholder: 'title' })
				findings.addText(
					[
						{ text: 'Revenue is up', options: { bullet: true, bold: true } },
						{ text: 'Retention held', options: { bullet: true } },
						{ text: 'Payback lengthened', options: { bullet: true, color: 'C00000' } },
					],
					{ x: 0.6, y: 1.8, w: 8.8, h: 3, fontSize: 18 }
				)
				findings.addNotes('Twenty minutes. Hold questions until the last slide.')

				const next = pres.addSlide({ masterTitle: 'NARRATIVE', sectionTitle: 'Next steps' })
				next.addText('Where to read the rest', { placeholder: 'title' })
				next.addText('The full report', {
					x: 0.6,
					y: 1.8,
					w: 8.8,
					h: 0.6,
					fontSize: 16,
					hyperlink: { url: 'https://example.com/report' },
				})
				next.addNotes('The link is the internal report, not the public summary.')
			},
			pptxgenjs: (pres) => {
				pres.defineSlideMaster({
					title: 'NARRATIVE',
					objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.5, w: 8.8, h: 1 } } }],
				})
				pres.addSection({ title: 'Findings' })
				pres.addSection({ title: 'Next steps' })

				const findings = pres.addSlide({ masterName: 'NARRATIVE', sectionTitle: 'Findings' })
				findings.background = { color: 'F2F2F2' }
				findings.addText('What we found', { placeholder: 'title' })
				findings.addText(
					[
						{ text: 'Revenue is up', options: { bullet: true, bold: true } },
						{ text: 'Retention held', options: { bullet: true } },
						{ text: 'Payback lengthened', options: { bullet: true, color: 'C00000' } },
					],
					{ x: 0.6, y: 1.8, w: 8.8, h: 3, fontSize: 18 }
				)
				findings.addNotes('Twenty minutes. Hold questions until the last slide.')

				const next = pres.addSlide({ masterName: 'NARRATIVE', sectionTitle: 'Next steps' })
				next.addText('Where to read the rest', { placeholder: 'title' })
				next.addText('The full report', {
					x: 0.6,
					y: 1.8,
					w: 8.8,
					h: 0.6,
					fontSize: 16,
					hyperlink: { url: 'https://example.com/report' },
				})
				next.addNotes('The link is the internal report, not the public summary.')
			},
		},
	},
	{
		id: 'table-deck',
		label: 'Table deck',
		what: 'A titled slide and a bordered table with a header row, fixed column widths and per-cell options.',
		build: {
			// Border thickness is the one call in this program the two libraries spell differently:
			// ts-pptx renamed `pt` to `width`, and warns rather than silently ignoring the old key.
			'ts-pptx': (pres) => {
				const slide = pres.addSlide()
				slide.addText('Revenue by region', { x: 0.6, y: 0.5, w: 8.8, h: 0.8, fontSize: 24, bold: true })
				slide.addTable(REGION_ROWS, {
					x: 0.6,
					y: 1.5,
					w: 8.8,
					colW: [4, 2.4, 2.4],
					rowH: 0.4,
					border: { type: 'solid', width: 1, color: 'D9D9D9' },
					fill: { color: 'FFFFFF' },
					fontSize: 12,
					valign: 'middle',
				})
			},
			pptxgenjs: (pres) => {
				const slide = pres.addSlide()
				slide.addText('Revenue by region', { x: 0.6, y: 0.5, w: 8.8, h: 0.8, fontSize: 24, bold: true })
				slide.addTable(REGION_ROWS, {
					x: 0.6,
					y: 1.5,
					w: 8.8,
					colW: [4, 2.4, 2.4],
					rowH: 0.4,
					border: { type: 'solid', pt: 1, color: 'D9D9D9' },
					fill: { color: 'FFFFFF' },
					fontSize: 12,
					valign: 'middle',
				})
			},
		},
	},
	{
		id: 'chart-deck',
		label: 'Chart deck',
		what: 'Three charts on three slides: a column chart with value labels, a two-series line chart, and a pie chart with percentages.',
		build: {
			// The signatures diverged at the detach: ts-pptx puts `type` in the options object,
			// upstream takes it as the first argument. Same three charts either way.
			'ts-pptx': (pres) => {
				pres.addSlide().addChart(BAR_DATA, {
					type: 'bar',
					barDir: 'col',
					x: 0.6,
					y: 0.6,
					w: 8.8,
					h: 5,
					showValue: true,
					showLegend: true,
					legendPos: 'b',
				})
				pres.addSlide().addChart(LINE_DATA, {
					type: 'line',
					x: 0.6,
					y: 0.6,
					w: 8.8,
					h: 5,
					lineSmooth: true,
					showLegend: true,
					legendPos: 'b',
				})
				pres.addSlide().addChart(PIE_DATA, {
					type: 'pie',
					x: 0.6,
					y: 0.6,
					w: 8.8,
					h: 5,
					showPercent: true,
					dataLabelColor: 'FFFFFF',
				})
			},
			pptxgenjs: (pres) => {
				pres.addSlide().addChart('bar', BAR_DATA, {
					barDir: 'col',
					x: 0.6,
					y: 0.6,
					w: 8.8,
					h: 5,
					showValue: true,
					showLegend: true,
					legendPos: 'b',
				})
				pres.addSlide().addChart('line', LINE_DATA, {
					x: 0.6,
					y: 0.6,
					w: 8.8,
					h: 5,
					lineSmooth: true,
					showLegend: true,
					legendPos: 'b',
				})
				pres.addSlide().addChart('pie', PIE_DATA, {
					x: 0.6,
					y: 0.6,
					w: 8.8,
					h: 5,
					showPercent: true,
					dataLabelColor: 'FFFFFF',
				})
			},
		},
	},
	{
		id: 'full-deck',
		label: 'Full deck',
		what: 'Every construct the shared baseline covers, in one deck: master, sections, background, text, hyperlink, notes, a preset shape, an image, a table and a chart.',
		build: {
			'ts-pptx': (pres) => {
				pres.defineSlideMaster({
					title: 'REVIEW',
					objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.5, w: 8.8, h: 1 } } }],
				})
				pres.addSection({ title: 'Quarter' })

				const cover = pres.addSlide({ masterTitle: 'REVIEW', sectionTitle: 'Quarter' })
				cover.background = { color: 'F2F2F2' }
				cover.addText('Q3 review', { placeholder: 'title' })
				cover.addShape('roundRect', { x: 0.6, y: 1.8, w: 3, h: 0.6, fill: { color: '4472C4' } })
				cover.addText('Read the report', {
					x: 0.6,
					y: 2.6,
					w: 4,
					h: 0.5,
					hyperlink: { url: 'https://example.com/report' },
				})
				cover.addImage({ data: PNG_1PX_URL, x: 6.4, y: 1.8, w: 2, h: 2 })
				cover.addNotes('Open on the number, not the agenda.')

				const numbers = pres.addSlide({ masterTitle: 'REVIEW', sectionTitle: 'Quarter' })
				numbers.addText('Where it came from', { placeholder: 'title' })
				numbers.addTable(REGION_ROWS, { x: 0.6, y: 1.6, w: 8.8, colW: [4, 2.4, 2.4], fontSize: 12 })

				pres.addSlide({ masterTitle: 'REVIEW', sectionTitle: 'Quarter' }).addChart(BAR_DATA, {
					type: 'bar',
					barDir: 'col',
					x: 0.6,
					y: 1.6,
					w: 8.8,
					h: 4.6,
					showValue: true,
				})
			},
			pptxgenjs: (pres) => {
				pres.defineSlideMaster({
					title: 'REVIEW',
					objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.5, w: 8.8, h: 1 } } }],
				})
				pres.addSection({ title: 'Quarter' })

				const cover = pres.addSlide({ masterName: 'REVIEW', sectionTitle: 'Quarter' })
				cover.background = { color: 'F2F2F2' }
				cover.addText('Q3 review', { placeholder: 'title' })
				cover.addShape('roundRect', { x: 0.6, y: 1.8, w: 3, h: 0.6, fill: { color: '4472C4' } })
				cover.addText('Read the report', {
					x: 0.6,
					y: 2.6,
					w: 4,
					h: 0.5,
					hyperlink: { url: 'https://example.com/report' },
				})
				cover.addImage({ data: PNG_1PX_BARE, x: 6.4, y: 1.8, w: 2, h: 2 })
				cover.addNotes('Open on the number, not the agenda.')

				const numbers = pres.addSlide({ masterName: 'REVIEW', sectionTitle: 'Quarter' })
				numbers.addText('Where it came from', { placeholder: 'title' })
				numbers.addTable(REGION_ROWS, { x: 0.6, y: 1.6, w: 8.8, colW: [4, 2.4, 2.4], fontSize: 12 })

				pres.addSlide({ masterName: 'REVIEW', sectionTitle: 'Quarter' }).addChart('bar', BAR_DATA, {
					barDir: 'col',
					x: 0.6,
					y: 1.6,
					w: 8.8,
					h: 4.6,
					showValue: true,
				})
			},
		},
	},
]

/**
 * One subject's arm, or a throw naming the program.
 *
 * Every program has both arms by construction, and the throw is what keeps that true: a
 * program added with one arm would otherwise be measured on one side and print a bundle
 * table row comparing a number with nothing.
 * @param {Program} program
 * @param {string} subject
 * @returns {(pres: any) => unknown}
 */
export function programArm(program, subject) {
	const build = program.build[subject]
	if (!build)
		throw new Error(
			'the bundle program "' + program.id + '" has no ' + subject + ' arm; every program must build with both libraries'
		)
	return build
}

/**
 * The declarations of every corpus constant a body names, in declaration order.
 * @param {string} body
 * @param {{elideOver?: number}} options
 * @returns {string[]}
 */
function constantsFor(body, options) {
	return Object.entries(PROGRAM_CONSTANTS)
		.filter(([name]) => new RegExp('\\b' + name + '\\b').test(body))
		.map(([name, value]) => 'const ' + name + ' = ' + literal(value, options))
}

/**
 * One program as a compilable module: the frame, the constants whole, the body.
 *
 * `elideOver: Infinity` because this text is compiled rather than read. The page's copy
 * elides a data URL at 60 characters, which is right for a reader and would be a lie here:
 * the bytes fed to the bundler have to be the bytes the library was given.
 * @param {Program} program
 * @param {string} subject
 * @returns {string}
 */
export function programModule(program, subject) {
	const frame = FRAME[subject]
	if (!frame) throw new Error('no bundle frame registered for subject "' + subject + '"')
	const body = functionBody(programArm(program, subject))
	const constants = constantsFor(body, { elideOver: Number.POSITIVE_INFINITY })
	return [
		frame.import,
		'',
		...constants,
		...(constants.length > 0 ? [''] : []),
		frame.construct,
		'',
		body,
		'',
		EXPORT_LINE,
		'',
	].join('\n')
}

/**
 * The frame with the program cut out of it, for the page to state once.
 *
 * Assembled from the same two strings {@link programModule} compiles, and carried into the
 * snapshot, so the frame a reader is shown cannot drift from the frame that was measured.
 * @param {string} subject
 * @returns {string}
 */
export function programFrame(subject) {
	const frame = FRAME[subject]
	if (!frame) throw new Error('no bundle frame registered for subject "' + subject + '"')
	return [frame.import, '', frame.construct, '// the program goes here', '', EXPORT_LINE].join('\n')
}

/**
 * The same program as the syntax page prints it: constants elided, no frame.
 * @param {Program} program
 * @param {string} subject
 * @returns {string}
 */
export function programSource(program, subject) {
	return renderSource(programArm(program, subject), PROGRAM_CONSTANTS)
}
