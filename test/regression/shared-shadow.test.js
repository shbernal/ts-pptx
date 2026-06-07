import JSZip from 'jszip'
import PptxGenJS from '../../dist/node.js'
import { defineRegressionSuite, assert } from '../helpers.js'

async function buildSlideXml(pres) {
	const buf = await pres.stream()
	const zip = await JSZip.loadAsync(buf)
	const entry = zip.file('ppt/slides/slide1.xml')
	if (!entry) throw new Error('slide1.xml missing')
	return entry.async('string')
}

// Extract each <p:sp>...</p:sp> block and return only the <a:effectLst>...</a:effectLst>
// substring inside, so we can compare shadow XML between shapes regardless of position.
function extractEffectLsts(xml) {
	const blocks = []
	const re = /<p:sp>[\s\S]*?<\/p:sp>/g
	let m
	while ((m = re.exec(xml)) !== null) {
		const sp = m[0]
		const eMatch = sp.match(/<a:effectLst>[\s\S]*?<\/a:effectLst>/)
		blocks.push(eMatch ? eMatch[0] : '')
	}
	return blocks
}

defineRegressionSuite('Shared shadow options', 'legacy bug-05', [
	{
		name: 'two addShape calls sharing one shadow object emit identical effectLst',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			const shadow = { type: 'outer', blur: 6, offset: 2, color: '000000', opacity: 0.15 }
			slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 1, w: 2, h: 1, shadow })
			slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 3, w: 2, h: 1, shadow })

			const xml = await buildSlideXml(pres)
			const effects = extractEffectLsts(xml)
			assert(effects.length === 2, 'expected 2 <p:sp> blocks; got ' + effects.length + '\n' + xml)
			assert(effects[0] !== '', 'expected first shape to have <a:effectLst>; got: ' + xml)
			assert(
				effects[0] === effects[1],
				'expected shared shadow to produce identical effectLst on both shapes.\n' +
					'first:  ' +
					effects[0] +
					'\nsecond: ' +
					effects[1]
			)
			// And the values must be sane EMUs, not double-converted.
			assert(
				effects[0].indexOf('blurRad="76200"') !== -1,
				'expected blurRad="76200" in shared effectLst; got: ' + effects[0]
			)
			assert(
				effects[0].indexOf('blurRad="967740000"') === -1,
				'unexpected double-converted blurRad in shared effectLst: ' + effects[0]
			)
		},
	},
	{
		name: 'shared shadow object literal not mutated after build',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			const shadow = { type: 'outer', blur: 6, offset: 2, color: '000000', opacity: 0.15 }
			slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 1, w: 2, h: 1, shadow })
			slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 3, w: 2, h: 1, shadow })
			await buildSlideXml(pres)

			assert(shadow.blur === 6, 'shadow.blur changed: ' + shadow.blur)
			assert(shadow.offset === 2, 'shadow.offset changed: ' + shadow.offset)
			assert(shadow.opacity === 0.15, 'shadow.opacity changed: ' + shadow.opacity)
			assert(shadow.color === '000000', 'shadow.color changed: ' + shadow.color)
			assert(shadow.angle === undefined, 'shadow.angle changed: ' + shadow.angle)
		},
	},
	{
		name: 'shared shadow across image+shape emits same EMU values (no double-conversion)',
		fn: async () => {
			const pres = new PptxGenJS()
			const slide = pres.addSlide()
			const png =
				'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkAAIAAAoAAv/lxKUAAAAASUVORK5CYII='
			const shadow = { type: 'outer', blur: 6, offset: 2, color: '000000', opacity: 0.15 }
			slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 1, w: 2, h: 1, shadow })
			slide.addImage({ data: png, x: 1, y: 3, w: 1, h: 1, shadow })

			const xml = await buildSlideXml(pres)
			const matches = xml.match(/<a:effectLst>[\s\S]*?<\/a:effectLst>/g) || []
			assert(
				matches.length === 2,
				'expected 2 effectLst blocks (one shape, one image); got ' + matches.length + '\n' + xml
			)
			// Whitespace differs between shape and image emission templates (cosmetic), but
			// the EMU values must match — that is what this regression guards against. Normalise spaces
			// then compare.
			const norm = (s) => s.replace(/\s+/g, ' ').replace(/ ?\/>/g, '/>').replace(/> </g, '><').trim()
			const a = norm(matches[0])
			const b = norm(matches[1])
			assert(
				a === b,
				'expected shape and image effectLst to carry the same EMU values when sharing one shadow object.\n' +
					'shape: ' +
					a +
					'\nimage: ' +
					b
			)
			// Belt-and-braces: both must use the original (pt-derived) EMU values, not double-converted.
			assert(a.indexOf('blurRad="76200"') !== -1, 'expected blurRad="76200" in shape effectLst; got: ' + a)
			assert(b.indexOf('blurRad="76200"') !== -1, 'expected blurRad="76200" in image effectLst; got: ' + b)
			assert(
				a.indexOf('blurRad="967740000"') === -1 && b.indexOf('blurRad="967740000"') === -1,
				'unexpected double-converted blurRad: shape=' + a + ' image=' + b
			)
		},
	},
])
