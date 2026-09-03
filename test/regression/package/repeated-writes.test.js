import TsPptx, { ShapeType } from '../../../dist/node.js'
import JSZip from 'jszip'
import { defineRegressionSuite, assert } from '../../helpers.js'

async function buildOnce(pres) {
	const buf = await pres.toBytes()
	const zip = await JSZip.loadAsync(buf)
	const entry = zip.file('ppt/slides/slide1.xml')
	if (!entry) throw new Error('slide1.xml missing')
	return entry.async('string')
}

defineRegressionSuite('Repeated presentation writes [legacy bug-04]', [
	{
		name: 'two writes on same Presentation produce identical slide1.xml (text/shape branch)',
		fn: async () => {
			const pres = new TsPptx()
			const slide = pres.addSlide()
			/** The caller's own `ShadowProps`; the writer normalizes a COPY and leaves this one alone.
			 * @type {import('../../../dist/node.js').ShadowProps & { _alpha?: number }} */
			const shadow = { type: 'outer', blur: 6, offset: 2, color: '000000', transparency: 85 }
			slide.addShape(ShapeType.rect, { x: 1, y: 1, w: 4, h: 2, shadow })

			const xml1 = await buildOnce(pres)
			const xml2 = await buildOnce(pres)
			assert(
				xml1 === xml2,
				'expected slide1.xml byte-equal across two writes; differ.\nfirst:\n' + xml1 + '\nsecond:\n' + xml2
			)
		},
	},
	{
		// The normalizer is pure, so the caller's object comes back exactly as written — the derived
		// `_alpha` lives on the definer's copy, which is what the emitter reads.
		name: 'user shadow object is not mutated across two writes',
		fn: async () => {
			const pres = new TsPptx()
			const slide = pres.addSlide()
			/** The caller's own `ShadowProps`; the writer normalizes a COPY and leaves this one alone.
			 * @type {import('../../../dist/node.js').ShadowProps & { _alpha?: number }} */
			const shadow = { type: 'outer', blur: 6, offset: 2, color: '000000', transparency: 85 }
			slide.addShape(ShapeType.rect, { x: 1, y: 1, w: 4, h: 2, shadow })

			await buildOnce(pres)
			await buildOnce(pres)

			assert(
				JSON.stringify(shadow) ===
					JSON.stringify({ type: 'outer', blur: 6, offset: 2, color: '000000', transparency: 85 }),
				'expected the caller shadow object untouched; got ' + JSON.stringify(shadow)
			)
		},
	},
	{
		name: 'two writes on same Presentation produce identical slide1.xml (image branch)',
		fn: async () => {
			const pres = new TsPptx()
			const slide = pres.addSlide()
			// 1x1 transparent PNG, base64
			const png =
				'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkAAIAAAoAAv/lxKUAAAAASUVORK5CYII='
			/** The caller's own `ShadowProps`; the writer normalizes a COPY and leaves this one alone.
			 * @type {import('../../../dist/node.js').ShadowProps & { _alpha?: number }} */
			const shadow = { type: 'outer', blur: 6, offset: 2, color: '000000', transparency: 85 }
			slide.addImage({ data: png, x: 1, y: 1, w: 1, h: 1, shadow })

			const xml1 = await buildOnce(pres)
			const xml2 = await buildOnce(pres)
			assert(
				xml1 === xml2,
				'expected image-branch slide1.xml byte-equal across two writes; differ.\nfirst:\n' + xml1 + '\nsecond:\n' + xml2
			)

			assert(shadow.blur === 6, 'expected image-branch shadow.blur to remain 6 (pt); got ' + shadow.blur)
			assert(
				shadow._alpha === undefined,
				'expected the image branch to leave the caller shadow untouched; got ' + JSON.stringify(shadow)
			)
		},
	},
	{
		name: 'regression - single write still emits sane shadow EMU values',
		fn: async () => {
			const pres = new TsPptx()
			const slide = pres.addSlide()
			slide.addShape(ShapeType.rect, {
				x: 1,
				y: 1,
				w: 4,
				h: 2,
				shadow: { type: 'outer', blur: 6, offset: 2, color: '000000', transparency: 85 },
			})
			const xml = await buildOnce(pres)
			assert(xml.indexOf('<a:effectLst>') !== -1, 'expected <a:effectLst> in single-write output')
			// blurRad=ptsToEmuLenient(6)=76200; dist=ptsToEmuLenient(2)=25400; dir=Math.round(270*60000)=16200000;
			// transparency 85 -> opacity 0.15 -> alpha=Math.round(0.15*100000)=15000
			assert(xml.indexOf('blurRad="76200"') !== -1, 'expected blurRad="76200"; got: ' + xml)
			assert(xml.indexOf('dist="25400"') !== -1, 'expected dist="25400"; got: ' + xml)
			assert(xml.indexOf('dir="16200000"') !== -1, 'expected dir="16200000"; got: ' + xml)
			assert(xml.indexOf('<a:alpha val="15000"/>') !== -1, 'expected <a:alpha val="15000"/>; got: ' + xml)
			// Belt-and-braces: no absurdly large EMU values that signal double-conversion
			assert(xml.indexOf('blurRad="967740000"') === -1, 'unexpected double-converted blurRad in output: ' + xml)
		},
	},
])
