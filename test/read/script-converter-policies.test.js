// The converter's cross-mapper policies, at the sites that used to skip them.
//
// `src/script/from-read/values.ts` declares the policies and explains each one; four were
// honoured by the shape mapper and skipped by the table, chart or gradient mapper. The round
// trip cannot see a skipped one: it excludes exactly the *declared* losses, and an undeclared
// one is invisible when both IRs come from the same reader. Nor can the byte-identity gate --
// no showcase deck carries a `dk1` cell fill or a table inside a degenerate group.
//
// So each case here builds the construct, converts it, and asserts what the IR says.

import { describe, test } from 'vitest'
import JSZip from 'jszip'
import { Presentation } from '../../dist/read.js'
import { readModelToIr } from '../../dist/script.js'
import { assert, assertEqual } from '../helpers.js'
import { authorRead } from './authored.js'

/** Apply `rewrite` to every slide part of `buf`, reload, and convert. */
async function irWithSlideXml(buf, rewrite) {
	const zip = await JSZip.loadAsync(buf)
	for (const name of Object.keys(zip.files)) {
		if (!/^ppt\/slides\/slide\d+\.xml$/.test(name)) continue
		zip.file(name, rewrite(await zip.file(name).async('string')))
	}
	const reopened = await Presentation.load(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
	return readModelToIr(reopened)
}

/** Every note construct the IR recorded. */
const constructs = (ir) => ir.fidelity.map((note) => note.construct)

/** The first `addTable` call on slide 1. */
const tableCall = (ir) => ir.slides[0].calls.find((call) => call.method === 'addTable')

/** A one-cell table whose fill, background and border are each a different mapped token. */
function tableDeck() {
	return authorRead((pres) => {
		pres.addSlide().addTable([[{ text: 'A', options: { fill: { color: 'accent1' } } }]], {
			x: 0.5,
			y: 0.5,
			w: 4,
			h: 1,
			fill: { color: 'accent2' },
			border: { type: 'solid', color: 'accent3', width: 1 },
		})
	})
}

describe('an unwritable scheme token is baked and noted, not passed through raw', () => {
	// `dk1` is one of the seven `ST_SchemeColorVal` values the write path's `clrMap` does not
	// carry. Passed through, the generated script warns `color/invalid-value` and paints the
	// default text colour -- a silently different deck, with no note to say so.
	test("a cell's fill token", async () => {
		const { buf } = await tableDeck()
		const ir = await irWithSlideXml(buf, (xml) => xml.replaceAll('val="accent1"', 'val="dk1"'))
		const cell = tableCall(ir).args[0][0][0]
		assert(cell.options.fill.color !== 'dk1', `the raw token must not reach the script; got ${cell.options.fill.color}`)
		assert(
			constructs(ir).includes('table.cell.fill.schemeToken'),
			'and the bake is noted; got ' + JSON.stringify(constructs(ir))
		)
	})

	test("the table's own background token", async () => {
		// The write API emits no `a:tblPr/a:solidFill` -- a table background is read-only today --
		// so the token is injected into the part the reader will see.
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addTable([[{ text: 'A' }]], { x: 0.5, y: 0.5, w: 4, h: 1 })
		})
		const ir = await irWithSlideXml(buf, (xml) =>
			xml.replace('<a:tblPr/>', '<a:tblPr><a:solidFill><a:schemeClr val="lt2"/></a:solidFill></a:tblPr>')
		)
		assert(
			constructs(ir).includes('table.fill.schemeToken'),
			'the table background bake is noted; got ' + JSON.stringify(constructs(ir))
		)
		const fill = tableCall(ir).args[1].tableFill
		assert(fill.color !== 'lt2', `the raw token must not reach the script; got ${fill.color}`)
	})

	test("a cell border's token", async () => {
		const { buf } = await tableDeck()
		const ir = await irWithSlideXml(buf, (xml) => xml.replaceAll('val="accent3"', 'val="hlink"'))
		assert(
			constructs(ir).includes('table.cell.borders.schemeToken'),
			'the border bake is noted; got ' + JSON.stringify(constructs(ir))
		)
	})

	test('a gradient stop token', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addShape('rect', {
				x: 1,
				y: 1,
				w: 2,
				h: 2,
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 0,
						stops: [
							{ position: 0, color: 'accent1' },
							{ position: 100, color: 'FFFFFF' },
						],
					},
				},
			})
		})
		const ir = await irWithSlideXml(buf, (xml) => xml.replaceAll('val="accent1"', 'val="folHlink"'))
		assert(
			constructs(ir).includes('fill.gradient.schemeToken'),
			'a stop that had to be baked says so; got ' + JSON.stringify(constructs(ir))
		)
	})

	test("a shape outline's token", async () => {
		// The outline kept its own copy of the colour ladder, and the copy baked an unwritable token
		// to hex without the note every other surface records.
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addShape('rect', { x: 1, y: 1, w: 2, h: 2, line: { color: 'accent4', width: 2 } })
		})
		const ir = await irWithSlideXml(buf, (xml) => xml.replaceAll('val="accent4"', 'val="dk1"'))
		assert(
			constructs(ir).includes('line.schemeToken'),
			'the outline bake is noted; got ' + JSON.stringify(constructs(ir))
		)
		const shape = ir.slides[0].calls.find((call) => call.method === 'addShape')
		assert(shape, 'the IR carries the shape')
		const options = shape.args.find((arg) => arg && typeof arg === 'object' && 'line' in arg)
		const color = /** @type {any} */ (options)?.line?.color
		assert(color !== 'dk1' && /^[0-9A-F]{6}$/.test(String(color)), `the token is baked to hex; got ${color}`)
	})

	/** A shape with an outer shadow in `color`, authored and read back. */
	const shadowDeck = (color) =>
		authorRead((pres) => {
			pres.addSlide().addShape('rect', {
				x: 1,
				y: 1,
				w: 2,
				h: 2,
				shadow: { type: 'outer', color, blur: 3, offset: 2, angle: 45 },
			})
		})
	/** The `shadow` or `glow` option of slide 1's first shape call. */
	const effectOf = (ir, key) =>
		ir.slides[0].calls.flatMap((call) => call.args).find((arg) => arg && typeof arg === 'object' && key in arg)?.[key]

	test("a shadow's writable token stays a token", async () => {
		// The shadow writer takes a scheme token (`createColorElement` writes `a:schemeClr`), so a
		// writable one has no reason to be baked. It used to be, and the copy stopped tracking its theme.
		const { presentation } = await shadowDeck('accent5')
		const shadow = effectOf(readModelToIr(presentation), 'shadow')
		assertEqual(shadow?.color, 'accent5', 'the shadow keeps its token')
	})

	test("a shadow's unwritable token is baked and noted", async () => {
		const { buf } = await shadowDeck('accent5')
		const ir = await irWithSlideXml(buf, (xml) => xml.replaceAll('val="accent5"', 'val="dk1"'))
		assert(
			constructs(ir).includes('shadow.schemeToken'),
			'the shadow bake is noted; got ' + JSON.stringify(constructs(ir))
		)
		const color = effectOf(ir, 'shadow')?.color
		assert(/^[0-9A-F]{6}$/.test(String(color)), `the token is baked to hex; got ${color}`)
	})

	test("a glow's unwritable token is baked and noted", async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addShape('rect', { x: 1, y: 1, w: 2, h: 2, fill: { color: 'FFFFFF' } })
		})
		const glow = '<a:effectLst><a:glow rad="50800"><a:schemeClr val="dk2"/></a:glow></a:effectLst></p:spPr>'
		let injected = 0
		const ir = await irWithSlideXml(buf, (xml) =>
			xml.replace('</p:spPr>', () => {
				injected++
				return glow
			})
		)
		assertEqual(injected, 1, 'the glow is injected into the shape')
		assert(constructs(ir).includes('glow.schemeToken'), 'the glow bake is noted; got ' + JSON.stringify(constructs(ir)))
		const color = effectOf(ir, 'glow')?.color
		assert(/^[0-9A-F]{6}$/.test(String(color)), `the token is baked to hex; got ${color}`)
	})

	test('a writable token still passes through as a token', async () => {
		// The point of preferring the token is that the copy keeps tracking its theme, so the
		// guard has to prove the bake is the exception rather than the rule.
		const { presentation } = await tableDeck()
		const ir = readModelToIr(presentation)
		assertEqual(tableCall(ir).args[0][0][0].options.fill.color, 'accent1', 'a mapped token survives')
		assert(!constructs(ir).includes('table.cell.fill.schemeToken'), 'and nothing is noted')
	})
})

describe('a graphic frame with no absolute frame is mapped like a shape', () => {
	// `absoluteFrame` is null whenever an enclosing group lacks a usable transform. The shape
	// mapper falls back to `resolvedFrame` and records `shape.frameInherited`, saying so;
	// graphic frames had their own mapper that returned `{}` with no fallback and no note, so a
	// table or chart in such a group was an *undeclared* loss. Both go through one mapper now.
	//
	// What the group mapper then does with the call is a separate matter -- a table is not a
	// `GroupChildProps` variant, so it is dropped and noted as `group.child` either way. The
	// note below is the observable difference this fix makes.
	test('a table inside a group with a degenerate child extent is noted', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addTable([[{ text: 'A' }]], { x: 1, y: 1, w: 4, h: 1 })
		})
		// A group whose `a:chExt` is zero: the child coordinate space has no scale, so nothing
		// downstream can compose a slide-absolute frame out of it.
		const ir = await irWithSlideXml(buf, (xml) => {
			const frame = /<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/.exec(xml)[0]
			const group =
				'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="99" name="Grp"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
				'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/>' +
				'<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
				frame +
				'</p:grpSp>'
			return xml.replace(frame, group)
		})
		assert(
			constructs(ir).includes('shape.frameInherited'),
			'the frame whose position had to be baked in says so; got ' + JSON.stringify(constructs(ir))
		)
	})

	test("a custom shape's points are scaled onto the box it is placed at", async () => {
		// A 2in triangle whose path viewport is doubled, so each path unit is half an EMU. Scaling
		// took `absoluteFrame`, which the degenerate group makes null, and fell back to a scale of 1:
		// the shape was placed at its resolved box and its points printed in raw path units.
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addShape('custGeom', {
				x: 1,
				y: 1,
				w: 2,
				h: 2,
				points: [
					{ x: 0, y: 0 },
					{ x: 2, y: 0 },
					{ x: 1, y: 2, close: true },
				],
			})
		})
		const ir = await irWithSlideXml(buf, (xml) => {
			const shape = /<p:sp>[\s\S]*?<\/p:sp>/.exec(xml)[0]
			const doubled = shape.replace(/<a:path w="(\d+)" h="(\d+)"/, (_, w, h) => `<a:path w="${w * 2}" h="${h * 2}"`)
			const group =
				'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="99" name="Grp"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
				'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/>' +
				'<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
				doubled +
				'</p:grpSp>'
			return xml.replace(shape, group)
		})
		assert(constructs(ir).includes('shape.frameInherited'), 'the shape took its resolved box')
		const call = JSON.stringify(ir.slides[0].calls)
		assert(call.includes('"x":"914400emu"'), 'the far corner lands at half the path width; got ' + call)
		assert(!call.includes('"x":"1828800emu"'), 'no point is printed in raw path units; got ' + call)
	})
})

describe('a shape nothing can place is dropped with a note', () => {
	// `absoluteFrame` and `resolvedFrame` are both null for a shape with no `a:xfrm` that is not a
	// placeholder. The connector mapper returned `null` there with no note, and every other kind was
	// emitted with no geometry, which the writer places at its default position.

	/** Delete the own `a:xfrm` of every shape and connector on the slide, and count them. */
	async function withoutFrames(buf) {
		let stripped = 0
		const ir = await irWithSlideXml(buf, (xml) =>
			xml.replace(/(<p:spPr>)\s*<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>/g, (_, open) => {
				stripped++
				return open
			})
		)
		return { ir, stripped }
	}

	test('an auto shape', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addShape('rect', { x: 1, y: 1, w: 2, h: 2, fill: { color: 'FF0000' } })
		})
		const { ir, stripped } = await withoutFrames(buf)
		assertEqual(stripped, 1, 'the shape loses its transform')
		assert(!ir.slides[0].calls.some((call) => call.method === 'addShape'), 'the unplaceable shape is not emitted')
		assert(
			constructs(ir).includes('shape.frameUnresolved'),
			'and its omission is noted; got ' + JSON.stringify(constructs(ir))
		)
	})

	test('a connector', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addConnector({ type: 'straight', x1: 1, y1: 1, x2: 4, y2: 3 })
		})
		const { ir, stripped } = await withoutFrames(buf)
		assertEqual(stripped, 1, 'the connector loses its transform')
		assert(!ir.slides[0].calls.some((call) => call.method === 'addConnector'), 'the connector is not emitted')
		assert(
			constructs(ir).includes('shape.frameUnresolved'),
			'and its omission is noted; got ' + JSON.stringify(constructs(ir))
		)
	})
})

describe('a fully opaque source emits no transparency key', () => {
	// `alphaToTransparency` documents that fully opaque is `undefined`, not `0`, because the
	// write path emits no `a:alphaModFix` for a zero transparency. One of five callers
	// implemented that; the rest passed the value through, and it stayed invisible only because
	// the canonicaliser drops `transparency: 0` as an implied default -- masking it rather than
	// agreeing with it.
	test('an explicit alpha of 100% on a shape fill', async () => {
		// The transparency only reaches the IR down the resolved-colour leg, which needs a fill
		// the reader cannot report as a literal: an unwritable scheme token is one.
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addShape('rect', { x: 1, y: 1, w: 2, h: 2, fill: { color: 'accent1', transparency: 40 } })
		})
		const ir = await irWithSlideXml(buf, (xml) =>
			xml.replaceAll('val="accent1"', 'val="dk1"').replaceAll('<a:alpha val="60000"/>', '<a:alpha val="100000"/>')
		)
		const shape = ir.slides[0].calls.find((call) => call.method === 'addShape')
		const fill = /** @type {Record<string, unknown>} */ (/** @type {Record<string, unknown>} */ (shape.args[1]).fill)
		// `dk1` resolves through the theme's colour map to the dark-1 slot, not to the caller's hex.
		assert(fill.color !== 'dk1', `the token was baked; got ${String(fill.color)}`)
		assertEqual(fill.transparency, undefined, 'and fully opaque states nothing')
	})
})

describe('a percentage keeps the precision the source wrote', () => {
	// The writer takes thousandths of a percent and the reader divides them back out, so rounding to
	// whole percent was a loss the converter added on its own. The re-read deck agreed with the
	// rounded value, so the round trip could not see it.
	test('a gradient stop position and transparency', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addShape('rect', {
				x: 1,
				y: 1,
				w: 2,
				h: 2,
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 0,
						stops: [
							{ position: 0, color: 'FF0000' },
							{ position: 33.333, color: '00FF00', transparency: 12.5 },
							{ position: 100, color: '0000FF' },
						],
					},
				},
			})
		})
		const shape = readModelToIr(presentation).slides[0].calls.find((call) => call.method === 'addShape')
		const stops = /** @type {any} */ (shape).args[1].fill.gradient.stops
		assertEqual(stops[1].position, 33.333, 'the stop sits where the source put it')
		assertEqual(stops[1].transparency, 12.5, 'at the transparency the source gave it')
	})

	test("a slide background's transparency, down to a hair under opaque", async () => {
		// The background kept its own copy of the conversion, with whole-percent rounding and a
		// zero rule of its own: an alpha of 0.996 stated nothing there and `0` everywhere else.
		const { buf, presentation } = await authorRead((pres) => {
			pres.addSlide().background = { color: 'C00000', transparency: 12.5 }
		})
		assertEqual(
			readModelToIr(presentation).slides[0].background?.transparency,
			12.5,
			'a fractional transparency survives'
		)
		let rewritten = 0
		const nearlyOpaque = await irWithSlideXml(buf, (xml) =>
			xml.replace('<a:alpha val="87500"/>', () => {
				rewritten++
				return '<a:alpha val="99600"/>'
			})
		)
		assertEqual(rewritten, 1, 'the background alpha is rewritten')
		assertEqual(nearlyOpaque.slides[0].background?.transparency, 0.4, 'and 0.4% is stated, not dropped')
	})
})

describe('an `xsd:boolean` attribute is parsed, not compared to `1`', () => {
	// `p:cNvSpPr/@txBox` is the sole discriminator between a text box and an auto shape, and
	// `xsd:boolean` admits `true` as well as `1`. Both producers this repo can author with --
	// its own write path and PowerPoint -- emit `1`, so every fixture agrees with a bare
	// `=== '1'` test and none of them can catch it. A foreign deck that spells it `true` was
	// converted with each of its text boxes turned into an auto shape: different autofit, wrap
	// and resize rules, no note, and a round trip that compares clean because both sides read
	// it the same wrong way.
	test('a text box whose `txBox` is spelled `true`', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addText('boxed', { x: 1, y: 1, w: 3, h: 1, isTextBox: true })
		})
		const ir = await irWithSlideXml(buf, (xml) => xml.replaceAll('txBox="1"', 'txBox="true"'))
		const text = ir.slides[0].calls.find((call) => call.method === 'addText')
		assertEqual(
			/** @type {Record<string, unknown>} */ (text.args[1]).isTextBox,
			true,
			'the other lexical form of the same boolean means the same thing'
		)
	})
})

describe('a fill transparency is read on every surface that can carry one', () => {
	// `a:alphaModFix` was read by the shape's copy of the fill ladder and by neither of the
	// table's two, so a table or a cell whose `a:solidFill` carried an alpha lost it -- with
	// nothing declaring the loss, because the round trip cannot see a key neither side produces.
	// Same staging as the shape case above: the transparency only reaches the IR down the
	// resolved-colour leg, so the token has to be one the reader cannot write back.
	test("a cell's fill", async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addTable([[{ text: 'A', options: { fill: { color: 'accent1', transparency: 40 } } }]], {
				x: 0.5,
				y: 0.5,
				w: 4,
				h: 1,
			})
		})
		const ir = await irWithSlideXml(buf, (xml) => xml.replaceAll('val="accent1"', 'val="dk1"'))
		const cell = tableCall(ir).args[0][0][0]
		assertEqual(cell.options.fill.transparency, 40, 'the cell keeps the alpha its source stated')
	})

	test("the table's own background", async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addTable([[{ text: 'A' }]], { x: 0.5, y: 0.5, w: 4, h: 1 })
		})
		const ir = await irWithSlideXml(buf, (xml) =>
			xml.replace(
				'<a:tblPr/>',
				'<a:tblPr><a:solidFill><a:schemeClr val="lt2"><a:alpha val="60000"/></a:schemeClr></a:solidFill></a:tblPr>'
			)
		)
		assertEqual(tableCall(ir).args[1].tableFill.transparency, 40, 'and so does the table behind it')
	})
})
