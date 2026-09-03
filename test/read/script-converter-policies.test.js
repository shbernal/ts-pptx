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
