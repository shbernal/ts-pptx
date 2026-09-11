import TsPptx, { InvalidOptionError, ShapeType } from '../../../dist/node.js'
import {
	defineRegressionSuite,
	build,
	readEntry,
	assert,
	assertEqual,
	assertIncludes,
	captureDiagnostics,
} from '../../helpers.js'

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
		// Each of these was tested for truthiness before its converter ran. `NaN` is falsy, so it
		// was dropped or replaced by the default without a word, while `Infinity` on the same
		// option threw. Only `0` and an absent value still mean "not stated".
		name: 'a NaN rotation, line width, margin or fill transparency is refused rather than dropped',
		fn: async () => {
			const cases = [
				{
					label: 'shape rotate',
					code: 'coord/non-finite',
					buildFn: (p) => p.addSlide().addShape('rect', { ...BOX, rotate: NaN }),
				},
				{
					label: 'shape line width',
					code: 'coord/non-finite',
					buildFn: (p) => p.addSlide().addShape('rect', { ...BOX, line: { color: 'FF0000', width: NaN } }),
				},
				{
					label: 'text line width',
					code: 'coord/non-finite',
					buildFn: (p) => p.addSlide().addText('x', { ...BOX, line: { color: 'FF0000', width: NaN } }),
				},
				{
					label: 'margin component',
					code: 'coord/non-finite',
					buildFn: (p) => p.addSlide().addText('x', { ...BOX, margin: [NaN, 0, 0, 0] }),
				},
				{
					label: 'fill transparency',
					code: 'percent/non-finite',
					buildFn: (p) => p.addSlide().addShape('rect', { ...BOX, fill: { color: 'FF0000', transparency: NaN } }),
				},
			]
			for (const { label, code, buildFn } of cases) assertEqual(await codeFrom(buildFn), code, label)
		},
	},
	{
		name: 'a NaN chart rotation, border width or marker line size is refused rather than dropped',
		fn: async () => {
			const data = [{ name: 'R', labels: ['A', 'B'], values: [1, 2] }]
			const options = {
				titleRotate: { showTitle: true, title: 'T', titleRotate: NaN },
				catAxisLabelRotate: { catAxisLabelRotate: NaN },
				valAxisLabelRotate: { valAxisLabelRotate: NaN },
				'chartArea border width': { chartArea: { border: { color: '000000', width: NaN } } },
				lineDataSymbolLineSize: { lineDataSymbolLineSize: NaN },
			}
			for (const [label, extra] of Object.entries(options)) {
				const code = await codeFrom((p) => p.addSlide().addChart(data, { ...BOX, type: 'line', ...extra }))
				assertEqual(code, 'coord/non-finite', label)
			}
		},
	},
	{
		// `0` is how those options say "not stated", so it still emits what leaving them out emits.
		name: 'a zero rotation or line width still reads as not stated',
		fn: async () => {
			const xfrm = await xfrmFor((p) => p.addSlide().addShape('rect', { ...BOX, rotate: 0 }))
			assert(!/\brot=/.test(xfrm), 'rotate: 0 emits no rot; got: ' + xfrm)
			const { zip } = await build((p) => p.addSlide().addShape('rect', { ...BOX, line: { color: 'FF0000', width: 0 } }))
			assertIncludes(await readEntry(zip, 'ppt/slides/slide1.xml'), '<a:ln w="12700"', 'width: 0 takes the 1pt default')
		},
	},
	{
		// A line width is a ranged number like any other: `Infinity` clamps to the top of
		// ST_LineWidth and warns. It used to collapse to `w="0"` without a word.
		name: 'an infinite line width clamps to the top of ST_LineWidth and warns',
		fn: async () => {
			const { result, codes } = await captureDiagnostics(() =>
				build((p) => p.addSlide().addShape('rect', { ...BOX, line: { color: 'FF0000', width: Infinity } }))
			)
			assertIncludes(await readEntry(result.zip, 'ppt/slides/slide1.xml'), '<a:ln w="20116800"', 'clamped to 1584pt')
			assert(codes.includes('line/width-out-of-range'), 'the clamp is reported; got ' + codes.join(', '))
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
