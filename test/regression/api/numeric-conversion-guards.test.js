import TsPptx, { InvalidOptionError, ShapeType } from '../../../dist/node.js'
import { defineRegressionSuite, build, readEntry, assert, assertEqual, assertIncludes } from '../../helpers.js'

// A converter that accepts garbage emits it: `Math.round(NaN * 100)` is `NaN` and
// `Math.round(Infinity * 60000)` is `Infinity`, and both serialize straight into an attribute
// that PowerPoint then reports as needing repair. Every case below hands a converter a value it
// cannot represent and asserts it refuses, plus the one case that is representable and was being
// mangled: an angle past a full turn.

async function caught(fn) {
	try {
		await fn()
		return null
	} catch (err) {
		return err
	}
}

/** The `err.code` of whatever building this deck throws, or `null` if it built. */
async function codeFrom(buildFn) {
	const err = await caught(() => build(buildFn))
	if (err === null) return null
	assert(err instanceof InvalidOptionError, 'expected an InvalidOptionError; got: ' + String(err))
	return err.code
}

/** The first shape's `<a:xfrm>` opening tag on slide 1. */
async function xfrmFor(buildFn) {
	const { zip } = await build(buildFn)
	const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
	const tag = /<a:xfrm[^>]*>/.exec(xml.slice(xml.indexOf('<p:sp>')))
	assert(tag, 'expected an <a:xfrm> on the shape; got: ' + xml.slice(0, 400))
	return tag[0]
}

const BOX = { x: 1, y: 1, w: 4, h: 2 }

/** A minimal `custGeom` triangle -- `<a:ahLst>` is only emitted for a freeform path. */
function custGeomWith(extra) {
	return {
		x: 1,
		y: 1,
		w: 2,
		h: 2,
		points: [
			{ x: 0, y: 0 },
			{ x: 2, y: 0 },
			{ x: 1, y: 2, close: true },
		],
		...extra,
	}
}

defineRegressionSuite('Numeric conversion guards', [
	{
		// `ptToHundredths` was the one converter in `units.ts` with no finiteness guard. An
		// `Infinity` reached it through every truthiness-guarded caller, came back as `Infinity`,
		// and was then reported as out of range and emitted as `sz="Infinity"`.
		name: 'a non-finite point measure is refused instead of reaching the attribute',
		fn: async () => {
			assertEqual(
				await codeFrom((p) => p.addSlide().addText('x', { ...BOX, fontSize: Infinity })),
				'coord/non-finite',
				'infinite font size'
			)
			assertEqual(
				await codeFrom((p) => p.addSlide().addText('x', { ...BOX, charSpacing: Infinity })),
				'coord/non-finite',
				'infinite character spacing'
			)
			assertEqual(
				await codeFrom((p) => p.addSlide().addText('x', { ...BOX, lineSpacing: Infinity })),
				'coord/non-finite',
				'infinite line spacing'
			)
			assertEqual(
				await codeFrom((p) => p.addSlide().addText('x', { ...BOX, paraSpaceBefore: Infinity })),
				'coord/non-finite',
				'infinite paragraph spacing'
			)
		},
	},
	{
		// `NaN` never reaches the converter on these paths -- every caller guards on truthiness,
		// and `NaN` is falsy, so the option reads as absent and the attribute is simply omitted.
		// Pinned because it is the reason the guard above is written against `Infinity`.
		name: 'a NaN fontSize reads as absent, leaving no sz at all',
		fn: async () => {
			const { zip } = await build((p) => p.addSlide().addText('x', { ...BOX, fontSize: NaN }))
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(!/sz="(NaN|Infinity)"/.test(xml), 'no degenerate sz reaches the part; got: ' + xml.slice(0, 400))
		},
	},
	{
		name: 'a finite out-of-range fontSize still clamps and warns, as before',
		fn: async () => {
			const { zip } = await build((p) => p.addSlide().addText('x', { ...BOX, fontSize: 99999 }))
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertIncludes(xml, 'sz="400000"', 'clamped to the top of ST_TextFontSize, not refused')
		},
	},
	{
		// `Math.round(Infinity * 60000)` is `Infinity`, which serialized as `rot="Infinity"`.
		name: 'a non-finite rotate is refused instead of reaching rot=',
		fn: async () => {
			assertEqual(
				await codeFrom((p) => p.addSlide().addShape('rect', { ...BOX, rotate: Infinity })),
				'coord/non-finite',
				'Infinite rotation'
			)
		},
	},
	{
		// The old reduction subtracted a single turn: 800 became 440 degrees, 370 became 10, and a
		// negative was never touched at all -- neither modular nor a pass-through.
		name: 'a rotation past a full turn reduces modularly, and one within a turn is untouched',
		fn: async () => {
			assertIncludes(
				await xfrmFor((p) => p.addSlide().addShape('rect', { ...BOX, rotate: 800 })),
				'rot="4800000"',
				'800 degrees is 80 degrees, i.e. 4800000 sixtieths-of-a-thousandth'
			)
			assertIncludes(
				await xfrmFor((p) => p.addSlide().addShape('rect', { ...BOX, rotate: -400 })),
				'rot="-2400000"',
				'-400 degrees is -40 degrees'
			)
			assertIncludes(
				await xfrmFor((p) => p.addSlide().addShape('rect', { ...BOX, rotate: -45 })),
				'rot="-2700000"',
				'a negative rotation within a turn keeps its sign; both spellings are valid ST_Angle'
			)
		},
	},
	{
		// A polar adjust handle's range is not modular: `maxAng: 540` used to become 180, silently
		// collapsing the handle's travel to a third of what was asked for.
		name: 'an adjust-handle angle past a full turn is emitted as given, not wrapped',
		fn: async () => {
			const { zip } = await build((p) =>
				p
					.addSlide()
					.addShape(
						ShapeType.custGeom,
						custGeomWith({ adjustHandles: [{ x: 0, y: 0, gdRefAng: 'a1', minAng: 0, maxAng: 540 }] })
					)
			)
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertIncludes(xml, 'maxAng="32400000"', '540 degrees, unreduced')
		},
	},
	{
		name: 'a non-finite adjust-handle angle is refused',
		fn: async () => {
			assertEqual(
				await codeFrom((p) =>
					p
						.addSlide()
						.addShape(
							ShapeType.custGeom,
							custGeomWith({ adjustHandles: [{ x: 0, y: 0, gdRefAng: 'a1', minAng: 0, maxAng: NaN }] })
						)
				),
				'coord/non-finite',
				'NaN maxAng'
			)
		},
	},
	{
		// Charts route their label rotations through the non-modular converter for the same reason:
		// a caller who writes -45 gets -45 back, and Infinity is refused rather than serialized.
		name: 'a non-finite chart label rotation is refused',
		fn: async () => {
			const data = [{ name: 'R', labels: ['A', 'B'], values: [1, 2] }]
			const err = await caught(() =>
				build((p) =>
					p.addSlide().addChart(data, { ...BOX, type: 'bar', catAxisLabelRotate: Infinity, catAxisLabelFontSize: 12 })
				)
			)
			assert(err instanceof InvalidOptionError, 'expected an InvalidOptionError; got: ' + String(err))
			assertEqual(err.code, 'coord/non-finite')
		},
	},
	{
		// The lenient EMU converter keeps its contract: a value it cannot read collapses the
		// feature to zero rather than taking the deck down. That is the deliberate other half of
		// the split, and the reason both converters exist.
		name: 'an unreadable line width still collapses to zero rather than throwing',
		fn: async () => {
			const pres = new TsPptx()
			pres.addSlide().addShape('rect', { ...BOX, line: { color: 'FF0000', width: /** @type {any} */ ('wide') } })
			const buf = await pres.toBytes()
			assert(buf.byteLength > 0, 'the deck still builds')
		},
	},
])
