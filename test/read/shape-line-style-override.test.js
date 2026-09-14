// Read-model and preserve-import coverage for an outline PowerPoint layers over a shape style, on
// shape-line-style-override.pptx. Four rectangles carry one preset shape style, whose
// `p:style/a:lnRef idx="2"` names accent2 shaded to 15%:
//
//   StyleOnly   nothing else, so no spPr/a:ln
//   WeightOnly  Line.Weight = 6pt, saved as <a:ln w="76200"/>
//   DashOnly    Line.DashStyle = dash, saved as <a:ln><a:prstDash val="dash"/></a:ln>
//   ColorOnly   Line.ForeColor = red, saved as <a:ln><a:solidFill><a:srgbClr val="FF0000"/>…
//
// Exported to PNG, WeightOnly's and DashOnly's outlines paint 622C0F, the same colour as
// StyleOnly's. An a:ln that states no fill is layered over the style line, not put in its place.

import { describe, test } from 'vitest'

import { assert, assertEqual, partXml } from '../helpers.js'
import { openFixture } from './corpus.js'

/** The theme's accent2, which the lnRef names. */
const STYLE_HEX = 'E97132'
/** accent2 after the lnRef's 15% shade: what the style outline paints. */
const STYLE_PAINTED = '622C0F'
/** The style line's fill as the preserve import bakes it. */
const STYLE_FILL = `<a:solidFill><a:srgbClr val="${STYLE_HEX}"><a:shade val="15000"/></a:srgbClr></a:solidFill>`
/** `a:lnStyleLst` entry 2's attributes: the ones no own `a:ln` in the fixture states. */
const STYLE_ATTRS = 'cap="flat" cmpd="sng" algn="ctr"'

function shapeNamed(slide, name) {
	const shape = slide.shapes.find((s) => s.name === name)
	assert(shape, `expected a shape named ${name}`)
	return shape
}

/** The `a:ln` in the `p:spPr` of the shape named `name`, from serialized slide XML. */
function ownLine(xml, name) {
	const sp = xml.split('<p:sp>').find((chunk) => chunk.includes(`name="${name}"`))
	assert(sp, `expected a shape named ${name} in the slide XML`)
	const spPr = sp.slice(sp.indexOf('<p:spPr>'), sp.indexOf('</p:spPr>'))
	return (spPr.match(/<a:ln[ >][\s\S]*?<\/a:ln>/) ?? [null])[0]
}

describe('Shape.resolvedLine under a p:style lnRef (shape-line-style-override.pptx)', () => {
	test('a shape with no a:ln resolves the style line colour', async () => {
		const line = shapeNamed((await openFixture('shape-line-style-override')).slides[0], 'StyleOnly').resolvedLine
		assert(line, 'the lnRef resolves a colour')
		assertEqual(line.hex, STYLE_HEX, 'lnRef accent2')
		assertEqual(line.effectiveHex, STYLE_PAINTED, 'accent2 shaded to 15%')
	})

	test('an a:ln that states only a width or a dash keeps the style line colour', async () => {
		const slide = (await openFixture('shape-line-style-override')).slides[0]
		assertEqual(shapeNamed(slide, 'WeightOnly').lineWidthPt, 6, 'WeightOnly states its own 6pt width')
		assertEqual(shapeNamed(slide, 'DashOnly').lineDash, 'dash', 'DashOnly states its own dash')
		for (const name of ['WeightOnly', 'DashOnly']) {
			const shape = shapeNamed(slide, name)
			assertEqual(shape.lineColor, null, `${name} states no line colour of its own`)
			assert(shape.resolvedLine, `${name} resolves a line colour`)
			assertEqual(shape.resolvedLine.effectiveHex, STYLE_PAINTED, `${name} paints the style colour`)
		}
	})

	test("an a:ln's own fill wins over the style line", async () => {
		const line = shapeNamed((await openFixture('shape-line-style-override')).slides[0], 'ColorOnly').resolvedLine
		assert(line, 'the own fill resolves')
		assertEqual(line.effectiveHex, 'FF0000', 'the own red, not the style colour')
	})
})

describe("importSlide({ theme: 'preserve' }) bakes a partial a:ln (shape-line-style-override.pptx)", () => {
	test('the style line fills in whatever each own a:ln leaves unstated', async () => {
		const target = await openFixture('empty')
		const source = await openFixture('shape-line-style-override')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		assert(!/<a:lnRef idx="[1-9]/.test(xml), 'every lnRef is neutralized')
		const solid = '<a:prstDash val="solid"/><a:miter lim="800000"/>'
		assertEqual(ownLine(xml, 'StyleOnly'), `<a:ln w="19050" ${STYLE_ATTRS}>${STYLE_FILL}${solid}</a:ln>`)
		assertEqual(
			ownLine(xml, 'WeightOnly'),
			`<a:ln w="76200" ${STYLE_ATTRS}>${STYLE_FILL}${solid}</a:ln>`,
			'the own width stays and the style supplies the rest'
		)
		assertEqual(
			ownLine(xml, 'DashOnly'),
			`<a:ln w="19050" ${STYLE_ATTRS}>${STYLE_FILL}<a:prstDash val="dash"/><a:miter lim="800000"/></a:ln>`,
			'the own dash stays and the style supplies the width, colour and join'
		)
		assertEqual(
			ownLine(xml, 'ColorOnly'),
			`<a:ln w="19050" ${STYLE_ATTRS}><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>${solid}</a:ln>`,
			'the own colour stays'
		)
	})
})
