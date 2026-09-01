import { defineRegressionSuite, build, readEntry, assert, assertEqual, captureDiagnostics } from '../../helpers.js'

// `shadow.transparency` and `shadow.angle` were the last two options outside the rule
// `docs/diagnostics.md` states for an out-of-range number: clamp to the nearest bound and warn,
// throw when the value is not a number at all.
//
// Both broke it in the way that rule exists to rule out. `transparency: 120` warned and then
// left `_alpha` unset, so the caller got a warning *and* a shadow at the 75% default opacity
// they never asked for. `angle` was guarded by `if (corrected.angle)`, so a `NaN` was falsy,
// fell past the range check entirely, and reached the emitter as `Math.round(NaN * 60000)` --
// `dir="NaN"` in the package, which is the degenerate output the guard was there to prevent.

const BOX = { x: 1, y: 1, w: 2, h: 1 }
const SHADOW = { type: 'outer', blur: 6, offset: 2, color: '000000' }

async function shadowXml(shadow) {
	const { zip } = await build((p) => p.addSlide().addShape('rect', { ...BOX, shadow }))
	const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
	const match = /<a:outerShdw[^>]*>[\s\S]*?<\/a:outerShdw>/.exec(xml)
	assert(match, 'expected an <a:outerShdw> on the shape; got:\n' + xml)
	return match[0]
}

async function codeThrownBy(shadow) {
	try {
		await build((p) => p.addSlide().addShape('rect', { ...BOX, shadow }))
	} catch (err) {
		return err?.code ?? null
	}
	return null
}

defineRegressionSuite('Shadow out-of-range policy', [
	{
		name: 'an out-of-range transparency clamps to the bound and paints there',
		fn: async () => {
			const { result: shdw, codes } = await captureDiagnostics(() => shadowXml({ ...SHADOW, transparency: 120 }))
			// 100% transparent is alpha 0. Before the fix `_alpha` stayed unset and the emitter
			// fell through to its 0.75 default, i.e. val="75000".
			assert(shdw.includes('<a:alpha val="0"/>'), 'transparency 120 should paint at 100; got: ' + shdw)
			assert(codes.includes('shadow/transparency-out-of-range'), 'the clamp is reported; got ' + codes.join(', '))
		},
	},
	{
		name: 'an in-range transparency is unaffected',
		fn: async () => {
			const { result: shdw, codes } = await captureDiagnostics(() => shadowXml({ ...SHADOW, transparency: 40 }))
			assert(shdw.includes('<a:alpha val="60000"/>'), 'transparency 40 is 60% opaque; got: ' + shdw)
			assertEqual(codes.length, 0, 'an in-range value warns about nothing')
		},
	},
	{
		name: 'a transparency that is not a number is refused rather than warned about',
		fn: async () => {
			assertEqual(await codeThrownBy({ ...SHADOW, transparency: NaN }), 'percent/non-finite', 'NaN transparency')
		},
	},
	{
		name: 'an out-of-range angle clamps to the nearest bound, not to a text shadow default',
		fn: async () => {
			const { result: shdw, codes } = await captureDiagnostics(() => shadowXml({ ...SHADOW, angle: 400 }))
			// 359 degrees in 60000ths. It used to become 270 -- `DEF_TEXT_SHADOW.angle`, which is
			// neither the nearest bound nor the `@default 0` the option documents.
			assert(shdw.includes('dir="21540000"'), 'angle 400 should clamp to 359; got: ' + shdw)
			assert(codes.includes('shadow/angle-out-of-range'), 'the clamp is reported; got ' + codes.join(', '))
		},
	},
	{
		name: 'a negative angle clamps to zero',
		fn: async () => {
			const { result: shdw, codes } = await captureDiagnostics(() => shadowXml({ ...SHADOW, angle: -45 }))
			assert(shdw.includes('dir="0"'), 'angle -45 should clamp to 0; got: ' + shdw)
			assert(codes.includes('shadow/angle-out-of-range'), 'the clamp is reported; got ' + codes.join(', '))
		},
	},
	{
		name: 'an angle of zero is a value, not an absence',
		fn: async () => {
			const { result: shdw, codes } = await captureDiagnostics(() => shadowXml({ ...SHADOW, angle: 0 }))
			assert(shdw.includes('dir="0"'), 'angle 0 emits dir="0"; got: ' + shdw)
			assertEqual(codes.length, 0, 'zero is in range and warns about nothing')
		},
	},
	{
		name: 'an angle that is not a number is refused rather than reaching the attribute',
		fn: async () => {
			assertEqual(await codeThrownBy({ ...SHADOW, angle: NaN }), 'shadow/angle-non-finite', 'NaN angle')
		},
	},
	{
		name: 'no shadow input reaches the package as the literal NaN',
		fn: async () => {
			// The guard that matters: whatever a bad input does, it never becomes an attribute
			// value. `blur` and `offset` stay deliberately lenient about a value that is not a
			// number at all (it collapses the feature to 0, silently) -- the clamp below is only
			// about numbers that are out of `blurRad`/`dist`'s unsigned range.
			const { result: shdw } = await captureDiagnostics(() =>
				shadowXml({ ...SHADOW, blur: NaN, offset: Infinity, angle: Infinity, transparency: Infinity })
			)
			assert(!shdw.includes('NaN'), 'no attribute may carry NaN; got: ' + shdw)
			assert(shdw.includes('blurRad="0"'), 'a non-finite blur collapses the feature; got: ' + shdw)
			assert(shdw.includes('dir="21540000"'), 'an infinite angle clamps like any other; got: ' + shdw)
		},
	},
	{
		// `blur` and `offset` are ST_PositiveCoordinate, so a negative one is not a smaller
		// shadow -- it is a value PowerPoint reports as needing repair. Their JSDoc used to
		// promise `0-100` and `0-200` and enforce neither; those two numbers are the limits of
		// PowerPoint's own spinners, not of the format, so the schema bound is what moved.
		name: 'a negative blur or offset clamps to zero and says so',
		fn: async () => {
			const { result: shdw, codes } = await captureDiagnostics(() => shadowXml({ ...SHADOW, blur: -6, offset: -2 }))
			assert(shdw.includes('blurRad="0"'), 'a negative blur clamps to 0; got: ' + shdw)
			assert(shdw.includes('dist="0"'), 'a negative offset clamps to 0; got: ' + shdw)
			assert(codes.includes('shadow/blur-out-of-range'), 'the blur clamp is reported; got ' + codes.join(', '))
			assert(codes.includes('shadow/offset-out-of-range'), 'the offset clamp is reported; got ' + codes.join(', '))
		},
	},
	{
		name: 'a blur past the PowerPoint spinner is emitted, not clamped',
		fn: async () => {
			// The sensitivity check for the bound above: if the doc'd `0-100` had been enforced
			// instead of the schema's range, this would paint at 100pt and warn.
			const { result: shdw, codes } = await captureDiagnostics(() => shadowXml({ ...SHADOW, blur: 150, offset: 300 }))
			assert(shdw.includes('blurRad="1905000"'), 'blur 150pt is legal OOXML; got: ' + shdw)
			assert(shdw.includes('dist="3810000"'), 'offset 300pt is legal OOXML; got: ' + shdw)
			assertEqual(codes.length, 0, 'a value the format accepts warns about nothing')
		},
	},
])
