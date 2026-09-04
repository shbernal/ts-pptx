/**
 * The construct-coverage corpus: one deck intent per probe, expressed twice.
 *
 * Every row of the comparison is produced by running both libraries and reading the bytes
 * they emit. Nothing here is an opinion about either project -- a probe states an intent
 * ("a slide that enters with a push transition"), each `build` expresses that intent in its
 * own library's idiom, and `scripts/comparison/measure.mjs` reports which of four things
 * happened. Transcribing one library's call into the other is how a comparison becomes
 * rigged, so the two arms of a probe deliberately do not have to look alike.
 *
 * Four outcomes, per probe per library:
 *
 *   emitted   the construct is in the named part
 *   absent    an API exists, the output does not carry the construct
 *   no-api    nothing in the public surface expresses the intent (`build` is null)
 *   error     the build threw; the message is recorded
 *
 * `no-api` is the only one of the four that is a *claim* rather than a reading, so the
 * harness checks it: the library's shipped bundle is grepped for `construct`, and a hit
 * fails the run unless this file acknowledges it in `sightings`. A token can appear in a
 * bundle without an API behind it -- `a:gradFill` is in both libraries' baked-in theme XML
 * -- and an acknowledgement is how that stays visible instead of being quietly waved
 * through. Anything in `sightings` reaches the snapshot and is printed beside its row.
 *
 * The corpus has to be able to lose. It carries a probe neither library can do, and the
 * harness reports the "upstream is ahead" set separately so an empty one is stated in words
 * rather than left to be inferred from its absence.
 */
import path from 'node:path'
import { ROOT } from '../script-utils.mjs'

/** A 1x1 transparent PNG. Inline, so an image probe reads nothing off disk. */
const PNG_1PX_B64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
/** ts-pptx takes a whole `data:` URL. */
const PNG_1PX_URL = 'data:image/png;base64,' + PNG_1PX_B64
/** pptxgenjs takes the MIME and payload without the `data:` scheme, as its own docs show. */
const PNG_1PX_BARE = 'image/png;base64,' + PNG_1PX_B64

/** An opaque OLE payload. The library never parses it; what is under test is the packaging. */
const OLE_BLOB_B64 = Buffer.from('ts-pptx comparison probe payload').toString('base64')

/** Raw OMML for one inline run. `latexToOmml` is a converter, not the thing being measured. */
const OMML_INLINE = '<m:oMath><m:r><m:t>n-1</m:t></m:r></m:oMath>'

/** A real glTF binary, so the deck a probe leaves behind is one PowerPoint opens. */
const CUBE_GLB = path.join(ROOT, 'demos', 'common', 'media', 'cube.glb')
/** A real font file, for the same reason: an embedded face PowerPoint will not bind is not evidence. */
const SILKSCREEN_TTF = path.join(ROOT, 'test', 'read', 'fixtures', 'fonts', 'Silkscreen-Regular.ttf')

const SLIDE1 = 'ppt/slides/slide1.xml'
const PRESENTATION = 'ppt/presentation.xml'

/** Chart data both libraries accept unchanged: the shape is one they share by descent. */
const BAR_DATA = [{ name: 'Revenue', labels: ['Q1', 'Q2', 'Q3'], values: [12, 19, 7] }]

/**
 * One deck intent, measured against both libraries.
 * @typedef {object} Probe
 * @property {string} id - kebab-case, stable; the table row key
 * @property {string} label - how the row reads on the page
 * @property {string} group - construct family, for grouping the table
 * @property {string} construct - the token asserted present in `part`
 * @property {string} part - package part the construct must appear in
 * @property {Record<string, ((pres: any) => unknown) | null>} build - per subject, or `null` for no API
 * @property {Record<string, string>} [sightings] - per subject, why `construct` is in its bundle with no API behind it
 */

/** @type {Probe[]} */
export const PROBES = [
	// -- Shared baseline ----------------------------------------------------------
	// Both libraries are expected to emit every one of these. They are the control
	// group: a comparison whose corpus only holds things one side cannot do proves
	// only that the corpus was chosen. They are also the harness's own tripwire --
	// see the gate in `measure.mjs`.
	{
		id: 'text-run',
		label: 'Text run',
		group: 'shared',
		construct: '<a:t>',
		part: SLIDE1,
		build: {
			'ts-pptx': (pres) => {
				pres.addSlide().addText('probe', { x: 1, y: 1, w: 4, h: 1 })
			},
			pptxgenjs: (pres) => {
				pres.addSlide().addText('probe', { x: 1, y: 1, w: 4, h: 1 })
			},
		},
	},
	{
		id: 'table',
		label: 'Table',
		group: 'shared',
		construct: '<a:tbl>',
		part: SLIDE1,
		build: {
			'ts-pptx': (pres) => {
				pres.addSlide().addTable(
					[
						[{ text: 'Region' }, { text: 'Units' }],
						[{ text: 'North' }, { text: '41' }],
					],
					{ x: 1, y: 1, w: 6 }
				)
			},
			pptxgenjs: (pres) => {
				pres.addSlide().addTable(
					[
						[{ text: 'Region' }, { text: 'Units' }],
						[{ text: 'North' }, { text: '41' }],
					],
					{ x: 1, y: 1, w: 6 }
				)
			},
		},
	},
	{
		id: 'image',
		label: 'Raster image',
		group: 'shared',
		construct: '<p:pic>',
		part: SLIDE1,
		build: {
			'ts-pptx': (pres) => {
				pres.addSlide().addImage({ data: PNG_1PX_URL, x: 1, y: 1, w: 2, h: 2 })
			},
			pptxgenjs: (pres) => {
				pres.addSlide().addImage({ data: PNG_1PX_BARE, x: 1, y: 1, w: 2, h: 2 })
			},
		},
	},
	{
		id: 'bar-chart',
		label: 'Bar chart',
		group: 'shared',
		construct: '<c:barChart>',
		part: 'ppt/charts/chart1.xml',
		build: {
			// The signatures diverged: ts-pptx puts `type` in the options object, upstream takes
			// it as the first argument. Same intent, each library called the way it asks to be.
			'ts-pptx': (pres) => {
				pres.addSlide().addChart(BAR_DATA, { type: 'bar', x: 1, y: 1, w: 6, h: 4 })
			},
			pptxgenjs: (pres) => {
				pres.addSlide().addChart('bar', BAR_DATA, { x: 1, y: 1, w: 6, h: 4 })
			},
		},
	},
	{
		id: 'hyperlink',
		label: 'External hyperlink',
		group: 'shared',
		construct: '<a:hlinkClick',
		part: SLIDE1,
		build: {
			'ts-pptx': (pres) => {
				pres.addSlide().addText('docs', { x: 1, y: 1, w: 4, h: 1, hyperlink: { url: 'https://example.com/' } })
			},
			pptxgenjs: (pres) => {
				pres.addSlide().addText('docs', { x: 1, y: 1, w: 4, h: 1, hyperlink: { url: 'https://example.com/' } })
			},
		},
	},
	{
		id: 'slide-master',
		label: 'User-defined slide master',
		group: 'shared',
		// The placeholder reference a slide inherits when it is filed under a defined master:
		// evidence the master was applied, not merely that a master part exists -- every deck
		// has one of those whether or not the caller asked for it.
		construct: '<p:ph',
		part: SLIDE1,
		build: {
			'ts-pptx': (pres) => {
				pres.defineSlideMaster({
					title: 'PROBE_MASTER',
					objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 1, y: 1, w: 8, h: 1 } } }],
				})
				pres.addSlide({ masterName: 'PROBE_MASTER' }).addText('probe', { placeholder: 'title' })
			},
			pptxgenjs: (pres) => {
				pres.defineSlideMaster({
					title: 'PROBE_MASTER',
					objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 1, y: 1, w: 8, h: 1 } } }],
				})
				pres.addSlide({ masterName: 'PROBE_MASTER' }).addText('probe', { placeholder: 'title' })
			},
		},
	},
	{
		id: 'sections',
		label: 'Sections',
		group: 'shared',
		construct: '<p14:sectionLst',
		part: PRESENTATION,
		build: {
			'ts-pptx': (pres) => {
				pres.addSection({ title: 'Findings' })
				pres.addSlide({ sectionTitle: 'Findings' }).addText('probe', { x: 1, y: 1, w: 4, h: 1 })
			},
			pptxgenjs: (pres) => {
				pres.addSection({ title: 'Findings' })
				pres.addSlide({ sectionTitle: 'Findings' }).addText('probe', { x: 1, y: 1, w: 4, h: 1 })
			},
		},
	},
	{
		id: 'speaker-notes',
		label: 'Speaker notes',
		group: 'shared',
		// The note's own text, not `<a:t>`: a notes slide carries a slide-number field whose run
		// would satisfy the looser token whether or not the note itself made it across.
		construct: 'probe note',
		part: 'ppt/notesSlides/notesSlide1.xml',
		build: {
			'ts-pptx': (pres) => {
				pres.addSlide().addNotes('probe note')
			},
			pptxgenjs: (pres) => {
				pres.addSlide().addNotes('probe note')
			},
		},
	},
	{
		id: 'preset-shape',
		label: 'Preset-geometry shape',
		group: 'shared',
		construct: '<a:prstGeom',
		part: SLIDE1,
		build: {
			'ts-pptx': (pres) => {
				pres.addSlide().addShape('roundRect', { x: 1, y: 1, w: 3, h: 2, fill: { color: '4472C4' } })
			},
			pptxgenjs: (pres) => {
				pres.addSlide().addShape('roundRect', { x: 1, y: 1, w: 3, h: 2, fill: { color: '4472C4' } })
			},
		},
	},
	{
		id: 'slide-background',
		label: 'Slide background colour',
		group: 'shared',
		construct: '<p:bg>',
		part: SLIDE1,
		build: {
			'ts-pptx': (pres) => {
				const slide = pres.addSlide()
				slide.background = { color: 'F2F2F2' }
				slide.addText('probe', { x: 1, y: 1, w: 4, h: 1 })
			},
			pptxgenjs: (pres) => {
				const slide = pres.addSlide()
				slide.background = { color: 'F2F2F2' }
				slide.addText('probe', { x: 1, y: 1, w: 4, h: 1 })
			},
		},
	},

	// -- Motion -------------------------------------------------------------------
	{
		id: 'slide-transition',
		label: 'Slide transition',
		group: 'motion',
		construct: '<p:transition',
		part: SLIDE1,
		build: {
			'ts-pptx': (pres) => {
				const slide = pres.addSlide()
				slide.transition = { type: 'push', speed: 'slow', variant: { dir: 'd' } }
				slide.addText('probe', { x: 1, y: 1, w: 4, h: 1 })
			},
			pptxgenjs: null,
		},
	},
	{
		id: 'build-animation',
		label: 'Build animation on a shape',
		group: 'motion',
		construct: '<p:timing>',
		part: SLIDE1,
		build: {
			'ts-pptx': (pres) => {
				const slide = pres.addSlide()
				slide.addText('probe', { x: 1, y: 1, w: 4, h: 1, objectName: 'Headline' })
				slide.addAnimation({ preset: 'fadeIn', objectName: 'Headline' })
			},
			pptxgenjs: null,
		},
	},

	// -- Embedding ----------------------------------------------------------------
	{
		id: 'ole-object',
		label: 'Embedded OLE object',
		group: 'embedding',
		construct: '<p:oleObj',
		part: SLIDE1,
		build: {
			'ts-pptx': (pres) => {
				pres.addSlide().addOleObject({ data: OLE_BLOB_B64, extn: 'bin', x: 1, y: 1, w: 4, h: 3 })
			},
			pptxgenjs: null,
		},
	},
	{
		id: 'model-3d',
		label: '3D model',
		group: 'embedding',
		construct: 'am3d:model3d',
		part: SLIDE1,
		build: {
			'ts-pptx': (pres) => {
				pres.addSlide().addModel3d({ path: CUBE_GLB, meterPerModelUnit: 0.5, x: 1, y: 1, w: 4, h: 3 })
			},
			pptxgenjs: null,
		},
	},
	{
		id: 'embedded-font',
		label: 'Embedded font face',
		group: 'embedding',
		construct: '<p:embeddedFontLst>',
		part: PRESENTATION,
		build: {
			'ts-pptx': async (pres) => {
				await pres.embedFont({ path: SILKSCREEN_TTF, typeface: 'Silkscreen' })
				pres.addSlide().addText('probe', { x: 1, y: 1, w: 4, h: 1, fontFace: 'Silkscreen' })
			},
			pptxgenjs: null,
		},
	},

	// -- Shapes, text and fills ---------------------------------------------------
	{
		id: 'connector',
		label: 'Connector between shapes',
		group: 'shapes',
		construct: '<p:cxnSp>',
		part: SLIDE1,
		build: {
			'ts-pptx': (pres) => {
				const slide = pres.addSlide()
				slide.addText('A', { objectName: 'BoxA', x: 1, y: 1, w: 1.5, h: 1 })
				slide.addText('B', { objectName: 'BoxB', x: 5, y: 1, w: 1.5, h: 1 })
				slide.addConnector({ type: 'elbow', x1: 2.5, y1: 1.5, x2: 5, y2: 1.5, startShape: 'BoxA', endShape: 'BoxB' })
			},
			pptxgenjs: null,
		},
	},
	{
		id: 'inline-math',
		label: 'Inline equation',
		group: 'text',
		construct: '<a14:m',
		part: SLIDE1,
		build: {
			'ts-pptx': (pres) => {
				pres.addSlide().addText([{ text: 'for all ' }, { math: OMML_INLINE, inline: true }, { text: ' terms' }], {
					x: 1,
					y: 1,
					w: 6,
					h: 1,
				})
			},
			pptxgenjs: null,
		},
	},
	{
		id: 'gradient-fill',
		label: 'Gradient shape fill',
		group: 'fills',
		construct: '<a:gradFill',
		part: SLIDE1,
		build: {
			'ts-pptx': (pres) => {
				pres.addSlide().addShape('rect', {
					x: 1,
					y: 1,
					w: 4,
					h: 2,
					fill: {
						gradient: {
							kind: 'linear',
							angle: 45,
							stops: [
								{ position: 0, color: '4472C4' },
								{ position: 100, color: 'ED7D31' },
							],
						},
					},
				})
			},
			pptxgenjs: null,
		},
		sightings: {
			// Both libraries ship the same baked-in Office theme, whose `fmtScheme` fill styles are
			// gradients. The token is in the bundle; no public option reaches it.
			pptxgenjs: 'appears only inside the bundled Office theme XML, which no API parameterises',
		},
	},
	{
		id: 'table-cell-bevel',
		label: '3D bevel on a table cell',
		group: 'tables',
		construct: '<a:cell3D',
		part: SLIDE1,
		build: {
			'ts-pptx': (pres) => {
				pres
					.addSlide()
					.addTable([[{ text: 'raised', options: { cell3D: { bevel: 'circle', width: 6, height: 6 } } }]], {
						x: 1,
						y: 1,
						w: 4,
					})
			},
			pptxgenjs: null,
		},
	},

	// -- Charts and navigation ----------------------------------------------------
	{
		id: 'chartex-funnel',
		label: 'Funnel chart (chartEx)',
		group: 'charts',
		construct: '<cx:chart>',
		part: 'ppt/charts/chartEx1.xml',
		build: {
			'ts-pptx': (pres) => {
				pres.addSlide().addChart([{ name: 'Stage', labels: ['Lead', 'Trial', 'Won'], values: [120, 48, 17] }], {
					type: 'funnel',
					x: 1,
					y: 1,
					w: 6,
					h: 4,
				})
			},
			pptxgenjs: null,
		},
	},
	{
		id: 'slide-zoom',
		label: 'Slide Zoom tile',
		group: 'navigation',
		construct: 'pslz:sldZm',
		part: 'ppt/slides/slide2.xml',
		build: {
			'ts-pptx': (pres) => {
				pres.addSlide().addText('target', { x: 1, y: 1, w: 4, h: 1 })
				pres.addSlide().addSlideZoom({ target: 1, x: 1, y: 1, w: 3, h: 1.7 })
			},
			pptxgenjs: null,
		},
	},

	// -- Where neither side reaches -----------------------------------------------
	{
		id: 'smartart',
		label: 'SmartArt diagram (write side)',
		group: 'diagrams',
		// ts-pptx reads diagrams (`src/read/api/diagram.ts`) and writes none; upstream does
		// neither. So this row is a shared gap rather than a difference, which is the kind of
		// fact a comparison that only printed its own wins would leave out.
		construct: '<dgm:relIds',
		part: SLIDE1,
		build: { 'ts-pptx': null, pptxgenjs: null },
	},
]

/** Every library the corpus measures, in the order the table shows them. */
export const SUBJECTS = ['ts-pptx', 'pptxgenjs']

/**
 * One probe by id, or a throw naming the ones that exist. Backs `--probe`.
 * @param {string} id
 * @returns {Probe}
 */
export function probeById(id) {
	const found = PROBES.find((probe) => probe.id === id)
	if (!found) throw new Error('no probe with id "' + id + '"; have: ' + PROBES.map((p) => p.id).join(', '))
	return found
}
