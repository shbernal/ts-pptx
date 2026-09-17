import { defineRegressionSuite, build, readEntry, assert, assertIncludes, captureDiagnostics } from '../../helpers.js'

// Every definer spelled its own default for an omitted `x`/`y`/`w`/`h`, and they disagreed on what
// "omitted" meant. `w: 0, h: 0` gave media a 2in square, a chart half the slide and an image its
// natural size, while a shape, an OLE object or a zoom got nothing, and none of them said a word.
// A `NaN` took the default wherever the definer tested with `||`. Now `0` is a stated extent on
// every one of them, kept and reported, and `NaN` reaches the coordinate converter.

const PNG =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
const SERIES = [{ name: 'S', labels: ['A', 'B'], values: [1, 2] }]

/** Each definer that takes a frame, handed `frame` on a fresh slide. */
const DEFINERS = {
	addShape: (slide, frame) => slide.addShape('rect', frame),
	addText: (slide, frame) => slide.addText('x', frame),
	addImage: (slide, frame) => slide.addImage({ data: PNG, ...frame }),
	addMedia: (slide, frame) => slide.addMedia({ type: 'video', data: 'video/mp4;base64,AAAA', ...frame }),
	addOleObject: (slide, frame) => slide.addOleObject({ data: 'AAECAwQFBgc=', ...frame }),
	addModel3d: (slide, frame) => slide.addModel3d({ data: 'Z2xURgIAAAA=', ...frame }),
	addChart: (slide, frame) => slide.addChart(SERIES, { type: 'bar', ...frame }),
}

/** The slide's shapes, after the spTree's own zero-size group properties. */
async function shapesXml(zip) {
	const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
	return xml.slice(xml.indexOf('</p:grpSpPr>'))
}

defineRegressionSuite('Authored frame defaults', [
	{
		name: 'a zero w and h stay zero on every definer, and are reported',
		fn: async () => {
			for (const [api, define] of Object.entries(DEFINERS)) {
				const { result: xml, codes } = await captureDiagnostics(async () => {
					const { zip } = await build((p) => define(p.addSlide(), { x: 1, y: 1, w: 0, h: 0 }))
					return shapesXml(zip)
				})
				assert(/<a:ext cx="0" cy="0"\/>/.test(xml), `${api}: expected a zero extent; got: ${xml.slice(0, 600)}`)
				assert(codes.includes('frame/zero-extent'), `${api}: expected frame/zero-extent; got ${codes.join(', ')}`)
			}
		},
	},
	{
		name: 'a slide zoom with no geometry is reported rather than silently drawn at zero size',
		fn: async () => {
			const { codes, messages } = await captureDiagnostics(() =>
				build((p) => {
					const host = p.addSlide()
					p.addSlide()
					host.addSlideZoom(/** @type {any} */ ({ target: 2 }))
				})
			)
			assert(codes.includes('frame/zero-extent'), `expected frame/zero-extent; got ${codes.join(', ')}`)
			// A zoom is the one definer whose default extent is zero, so a caller who stated no frame
			// used to be told "w is 0 and h is 0" -- our default quoted back as if they had written it.
			const text = (messages ?? []).join(' | ')
			assertIncludes(text, 'no w or h was given', `the message names what was missing; got ${text}`)
		},
	},
	{
		// The other half of the same distinction: a zero the caller did write is still quoted.
		name: 'a zero the caller stated is quoted back, beside an axis they left out',
		fn: async () => {
			const { messages } = await captureDiagnostics(() =>
				build((p) => {
					const host = p.addSlide()
					p.addSlide()
					host.addSlideZoom(/** @type {any} */ ({ target: 2, w: 0, h: 2 }))
				})
			)
			const text = (messages ?? []).join(' | ')
			assertIncludes(text, 'w is 0', `the stated zero is named; got ${text}`)
			assert(!text.includes('no default size'), `and it is not blamed on the default; got ${text}`)
		},
	},
	{
		// A line is drawn with no height or no width on purpose.
		name: 'a line with no height or no width is not reported',
		fn: async () => {
			const { codes } = await captureDiagnostics(() =>
				build((p) => {
					const slide = p.addSlide()
					slide.addShape('line', { x: 1, y: 1, w: 4, h: 0, line: { color: '000000' } })
					slide.addShape('line', { x: 1, y: 1, w: 0, h: 3, line: { color: '000000' } })
					slide.addText('', { shape: 'line', x: 1, y: 2, w: 4, h: 0 })
				})
			)
			assert(!codes.includes('frame/zero-extent'), `a line should not warn; got ${codes.join(', ')}`)
		},
	},
	{
		name: 'a NaN extent is refused rather than replaced by the default',
		fn: async () => {
			for (const [api, define] of Object.entries(DEFINERS)) {
				let code = null
				try {
					await build((p) => define(p.addSlide(), { x: 1, y: 1, w: Number.NaN, h: 1 }))
				} catch (err) {
					code = err?.code ?? null
				}
				assert(code === 'coord/non-finite', `${api}: expected coord/non-finite; got ${code}`)
			}
		},
	},
])
