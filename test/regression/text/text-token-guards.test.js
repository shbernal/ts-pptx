import { defineRegressionSuite, build, readEntry, assert, captureDiagnostics } from '../../helpers.js'

// Enumerated text tokens a caller supplies reach an attribute only when the schema has them. A token
// outside its `ST_` type used to be written as given, which is a part PowerPoint reports as needing
// repair, and nothing said so.

const SLIDE_XML = 'ppt/slides/slide1.xml'

/** Build a one-slide deck, returning its slide XML and the diagnostic codes raised. */
async function slideWith(build_) {
	const { result: xml, codes } = await captureDiagnostics(async () => {
		const { zip } = await build(build_)
		return readEntry(zip, SLIDE_XML)
	})
	return { xml, codes }
}

defineRegressionSuite('Text token guards', [
	{
		name: 'an unknown strike, caps or underline style warns and is left out',
		fn: async () => {
			const { xml, codes } = await slideWith((p) => {
				p.addSlide().addText(
					[
						{
							text: 'run',
							options: /** @type {any} */ ({
								strike: 'bogusStrike',
								caps: 'bogusCaps',
								underline: { style: 'bogusU' },
							}),
						},
					],
					{ x: 1, y: 1, w: 4, h: 1 }
				)
			})
			assert(!/bogus/.test(xml), `no bad token reaches the part; got: ${xml}`)
			for (const code of ['text/invalid-strike', 'text/invalid-caps', 'text/invalid-underline']) {
				assert(codes.includes(code), `${code} is raised; got ${JSON.stringify(codes)}`)
			}
		},
	},
	{
		name: 'valid strike, caps and underline tokens are written as given',
		fn: async () => {
			const { xml, codes } = await slideWith((p) => {
				p.addSlide().addText(
					[{ text: 'run', options: { strike: 'dblStrike', caps: 'small', underline: { style: 'wavyDbl' } } }],
					{ x: 1, y: 1, w: 4, h: 1 }
				)
			})
			assert(/strike="dblStrike" cap="small" u="wavyDbl"/.test(xml), `the tokens are written; got: ${xml}`)
			assert(codes.length === 0, `and nothing is reported; got ${JSON.stringify(codes)}`)
		},
	},
	{
		name: 'an unknown tab stop alignment warns and the stop takes the schema default',
		fn: async () => {
			const { xml, codes } = await slideWith((p) => {
				p.addSlide().addText('a\tb', {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					tabStops: [
						{ position: 1, alignment: /** @type {any} */ ('bogusAlign') },
						{ position: 2, alignment: 'dec' },
					],
				})
			})
			assert(xml.includes('<a:tab pos="914400"/>'), `the bad alignment is left out; got: ${xml}`)
			assert(xml.includes('<a:tab pos="1828800" algn="dec"/>'), `a valid one is kept; got: ${xml}`)
			assert(codes.includes('text/invalid-tab-alignment'), `and the caller is told; got ${JSON.stringify(codes)}`)
		},
	},
	{
		// `type` is required on `a:buAutoNum`, so an unknown scheme cannot simply be left out.
		name: 'an unknown numberType warns and the list counts in arabicPeriod',
		fn: async () => {
			const { xml, codes } = await slideWith((p) => {
				p.addSlide().addText('item', {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					bullet: { type: 'number', numberType: /** @type {any} */ ('bogusNum') },
				})
			})
			assert(xml.includes('<a:buAutoNum type="arabicPeriod" startAt="1"/>'), `it counts in arabicPeriod; got: ${xml}`)
			assert(codes.includes('bullet/invalid-number-type'), `and the caller is told; got ${JSON.stringify(codes)}`)
		},
	},
	{
		name: 'a numberType beyond the sixteen Latin schemes is written as given',
		fn: async () => {
			const { xml, codes } = await slideWith((p) => {
				p.addSlide().addText('item', {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					bullet: { type: 'number', numberType: 'circleNumDbPlain' },
				})
			})
			assert(xml.includes('<a:buAutoNum type="circleNumDbPlain" startAt="1"/>'), `the scheme is written; got: ${xml}`)
			assert(codes.length === 0, `and nothing is reported; got ${JSON.stringify(codes)}`)
		},
	},
])
