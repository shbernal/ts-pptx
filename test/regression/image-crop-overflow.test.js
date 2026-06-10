import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// 1x1 transparent PNG
const PNG_DATA =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

async function expectCropError(sizingOpts, expectedFragment) {
	let err
	try {
		await build((p) => {
			const s = p.addSlide()
			s.addImage({ data: PNG_DATA, x: 1, y: 1, w: 4, h: 3, sizing: sizingOpts })
		})
	} catch (e) {
		err = e
	}
	assert(err, 'expected build to throw for out-of-bounds crop')
	const msg = String(err?.message || err)
	assert(msg.includes(expectedFragment), `expected error to include "${expectedFragment}"; got: ${msg}`)
}

defineRegressionSuite('Image crop overflow', [
	{
		name: 'crop window overflowing right edge throws with right-edge description',
		fn: async () => {
			// x=2 + w=3 = 5 > image w=4 → right overflow
			await expectCropError({ type: 'crop', x: 2, y: 0, w: 3, h: 3 }, 'past right edge')
		},
	},
	{
		name: 'crop window overflowing bottom edge throws with bottom-edge description',
		fn: async () => {
			// y=1 + h=3 = 4 > image h=3 → bottom overflow
			await expectCropError({ type: 'crop', x: 0, y: 1, w: 4, h: 3 }, 'past bottom edge')
		},
	},
	{
		name: 'negative x throws with left-edge description',
		fn: async () => {
			await expectCropError({ type: 'crop', x: -0.5, y: 0, w: 3, h: 3 }, 'past left edge')
		},
	},
	{
		name: 'negative y throws with top-edge description',
		fn: async () => {
			await expectCropError({ type: 'crop', x: 0, y: -0.5, w: 4, h: 2 }, 'past top edge')
		},
	},
	{
		name: 'in-bounds crop at exact image edge does not throw',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				// x=0, y=0, w=4, h=3 exactly fills the 4x3 image — r=0, b=0
				s.addImage({ data: PNG_DATA, x: 1, y: 1, w: 4, h: 3, sizing: { type: 'crop', x: 0, y: 0, w: 4, h: 3 } })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:srcRect l="0" r="0" t="0" b="0"\/>/.test(xml), 'expected zero-margin srcRect; got: ' + xml)
		},
	},
	{
		name: 'in-bounds crop with interior window emits correct srcRect percentages',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				// 4x3 image, crop window x=0.5, y=0.5, w=3, h=2
				// l=0.5/4=12.5% → 12500, r=(4-3.5)/4=12.5% → 12500
				// t=0.5/3≈16.7% → 16667, b=(3-2.5)/3≈16.7% → 16667
				s.addImage({ data: PNG_DATA, x: 1, y: 1, w: 4, h: 3, sizing: { type: 'crop', x: 0.5, y: 0.5, w: 3, h: 2 } })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				/<a:srcRect l="12500" r="12500" t="16667" b="16667"\/>/.test(xml),
				'expected srcRect l=12500 r=12500 t=16667 b=16667; got: ' + xml
			)
		},
	},
])
