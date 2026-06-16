import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// genXmlPlaceholder previously re-looked-up the already-mapped OOXML value
// (e.g. 'pic', 'tbl') in PLACEHOLDER_TYPE_MAP, whose keys are the friendly
// input names ('image', 'table') - so the `type` attribute was silently
// dropped for image and table placeholders. The mapping now accepts either
// the friendly key or the OOXML value and emits the OOXML value.
defineRegressionSuite('Placeholder type attribute', 'genXmlPlaceholder-type-map', [
	{
		name: 'image + table placeholders (OOXML-value form) emit type="pic"/"tbl"',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'PH_TYPE_MASTER',
					objects: [
						{ placeholder: { options: { name: 'pic1', type: 'pic', x: 0.5, y: 0.5, w: 4, h: 3 }, text: '' } },
						{ placeholder: { options: { name: 'tbl1', type: 'tbl', x: 5, y: 0.5, w: 4, h: 3 }, text: '' } },
					],
				})
				p.addSlide({ masterName: 'PH_TYPE_MASTER' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<p:ph[^>]*type="pic"/.test(xml), 'expected <p:ph type="pic"/>; got: ' + xml)
			assert(/<p:ph[^>]*type="tbl"/.test(xml), 'expected <p:ph type="tbl"/>; got: ' + xml)
		},
	},
	{
		name: 'image + table placeholders (friendly-key form) emit type="pic"/"tbl"',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'PH_TYPE_MASTER_KEYS',
					objects: [
						{ placeholder: { options: { name: 'pic1', type: 'image', x: 0.5, y: 0.5, w: 4, h: 3 }, text: '' } },
						{ placeholder: { options: { name: 'tbl1', type: 'table', x: 5, y: 0.5, w: 4, h: 3 }, text: '' } },
					],
				})
				p.addSlide({ masterName: 'PH_TYPE_MASTER_KEYS' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<p:ph[^>]*type="pic"/.test(xml), 'expected <p:ph type="pic"/>; got: ' + xml)
			assert(/<p:ph[^>]*type="tbl"/.test(xml), 'expected <p:ph type="tbl"/>; got: ' + xml)
		},
	},
	{
		name: 'unknown placeholder type emits no type attribute',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('x', { x: 1, y: 1, w: 3, h: 1, placeholder: 'bogusType' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(!/type="bogusType"/.test(xml), 'unknown placeholder type must not emit a type attribute; got: ' + xml)
		},
	},
])
