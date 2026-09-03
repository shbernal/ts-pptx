import { ChartType } from '../../../dist/node.js'
import { defineRegressionSuite, build, readEntry, assert, captureDiagnostics } from '../../helpers.js'

const SLIDE_XML = 'ppt/slides/slide1.xml'
const MASTER_XML = 'ppt/slideMasters/slideMaster1.xml'

/** The `ppt/charts/chartN.xml` a one-chart deck emits. */
async function chartXml(zip) {
	return readEntry(zip, 'ppt/charts/chart1.xml')
}

// Values whose `ST_` type has a range, reaching an attribute that had no guard on the way in.
// Each of these had a sibling site that DID guard the same value, which is why none of them was
// caught by a reader: the library already knew the rule and applied it somewhere else.
defineRegressionSuite('Schema value guards', [
	{
		// `ST_TextFontSize` is 100..400000. `clampFontSizeSz` bounds it; six `sz` sites called
		// `ptToHundredths` directly, so one chart could carry both readings of the same number.
		name: 'every chart font size is bounded by ST_TextFontSize, not just the data-label one',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addChart([{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }], {
						type: ChartType.bar,
						x: 1,
						y: 1,
						w: 6,
						h: 3,
						showLegend: true,
						legendFontSize: 5000,
						catAxisLabelFontSize: 5000,
						valAxisLabelFontSize: 0.4,
						dataTableFontSize: 5000,
						showTitle: true,
						title: 'over',
						titleFontSize: 5000,
					})
				})
				return chartXml(zip)
			})
			const sizes = [...xml.matchAll(/sz="(-?\d+)"/g)].map((m) => Number(m[1]))
			assert(sizes.length > 0, 'expected some sz attributes')
			assert(
				sizes.every((n) => n >= 100 && n <= 400000),
				'every sz is inside ST_TextFontSize; got ' + JSON.stringify(sizes)
			)
			assert(codes.includes('font/size-out-of-range'), 'and the caller is told; got ' + JSON.stringify(codes))
		},
	},
	{
		// The master's text styles state their measures in inches, so they reached `inch2Emu` and
		// never the paragraph clamps: `marL` came out at ~18000x the ST_TextMargin ceiling.
		name: 'master textStyles margins and sizes are bounded by their own ST_ types',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.defineSlideMaster({
						title: 'GUARDS',
						textStyles: { body: [{ fontSize: 99999, marginLeft: 1e6, indent: -1e6 }] },
					})
					p.addSlide({ masterTitle: 'GUARDS' }).addText('x', { x: 1, y: 1, w: 2, h: 1 })
				})
				return readEntry(zip, MASTER_XML)
			})
			for (const [, value] of xml.matchAll(/sz="(-?\d+)"/g)) {
				assert(Number(value) >= 100 && Number(value) <= 400000, `sz ${value} is outside ST_TextFontSize`)
			}
			for (const [, value] of xml.matchAll(/marL="(-?\d+)"/g)) {
				assert(Number(value) >= 0 && Number(value) <= 51206400, `marL ${value} is outside ST_TextMargin`)
			}
			for (const [, value] of xml.matchAll(/indent="(-?\d+)"/g)) {
				assert(Number(value) >= -51206400 && Number(value) <= 51206400, `indent ${value} is outside ST_TextIndent`)
			}
			assert(codes.includes('font/size-out-of-range'), 'the size is reported; got ' + JSON.stringify(codes))
			assert(
				codes.includes('text/paragraph-margin-out-of-range'),
				'the margin is reported; got ' + JSON.stringify(codes)
			)
		},
	},
	{
		// `indentLevel` was validated for `a:p/@lvl` and then multiplied into the bullet arm's
		// default `marL` regardless, so the warning said the value was ignored while it reached
		// the package at ~6700x the ST_TextMargin ceiling.
		name: 'a rejected indentLevel is rejected for the bullet margin too',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addText('deep', { x: 1, y: 1, w: 4, h: 1, bullet: true, indentLevel: 1e6 })
				})
				return readEntry(zip, SLIDE_XML)
			})
			assert(!xml.includes('lvl="1000000"'), `the level itself is dropped; got: ${xml}`)
			assert(xml.includes('marL="342900"'), `the margin falls back to level 0; got: ${xml}`)
			assert(codes.includes('text/invalid-indent-level'), 'and the caller is told; got ' + JSON.stringify(codes))
		},
	},
	{
		name: 'a paraMarginLeft past ST_TextMargin is clamped, as it always was for the explicit form',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addText('wide', { x: 1, y: 1, w: 4, h: 1, bullet: true, paraMarginLeft: 1e6 })
				})
				return readEntry(zip, SLIDE_XML)
			})
			assert(xml.includes('marL="51206400"'), `clamped to the ST_TextMargin ceiling; got: ${xml}`)
			assert(
				codes.includes('text/paragraph-margin-out-of-range'),
				'and the caller is told; got ' + JSON.stringify(codes)
			)
		},
	},
	{
		// `BorderProps.dashType` and `ShapeLineProps.dashType` are the same type reaching the same
		// attribute; only the table border checked it.
		name: 'a shape line dash outside ST_PresetLineDashVal is reported and falls back',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addShape('rect', { x: 1, y: 1, w: 2, h: 1, line: { color: 'FF0000', dashType: 'bogusDash' } })
				})
				return readEntry(zip, SLIDE_XML)
			})
			assert(!xml.includes('bogusDash'), `the bad value must not reach the part; got: ${xml}`)
			assert(xml.includes('<a:prstDash val="solid"/>'), `it falls back to solid; got: ${xml}`)
			assert(codes.includes('border/invalid-dash-type'), 'and the caller is told; got ' + JSON.stringify(codes))
		},
	},
	{
		// `ST_LineEndType` had no runtime tuple at all, so the two arrow options went in verbatim.
		name: 'an arrowhead outside ST_LineEndType is reported and the attribute is omitted',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addShape('line', {
						x: 1,
						y: 1,
						w: 2,
						h: 1,
						line: { color: 'FF0000', beginArrowType: 'wedge', endArrowType: 'arrow' },
					})
				})
				return readEntry(zip, SLIDE_XML)
			})
			assert(!xml.includes('wedge'), `the bad value must not reach the part; got: ${xml}`)
			assert(!xml.includes('<a:headEnd'), `the unresolvable end is omitted; got: ${xml}`)
			assert(xml.includes('<a:tailEnd type="arrow"/>'), `the legal one still emits; got: ${xml}`)
			assert(codes.includes('line/invalid-arrow-type'), 'and the caller is told; got ' + JSON.stringify(codes))
		},
	},
	{
		// `textWarp` was typed as a free-form string and emitted verbatim into `a:prstTxWarp/@prst`.
		name: 'a text warp outside ST_TextShapeType is reported and the child is omitted',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addText('warped', { x: 1, y: 1, w: 4, h: 1, textWarp: 'textLoopTheLoop' })
				})
				return readEntry(zip, SLIDE_XML)
			})
			assert(!xml.includes('textLoopTheLoop'), `the bad value must not reach the part; got: ${xml}`)
			assert(!xml.includes('<a:prstTxWarp'), `the whole child is omitted; got: ${xml}`)
			assert(codes.includes('text/invalid-warp'), 'and the caller is told; got ' + JSON.stringify(codes))
		},
	},
	{
		// `a:glow/@rad` is ST_PositiveCoordinate, the same type as the shadow's `blurRad` beside it
		// — which was converted through the clamping helper while the glow was a bare multiply.
		name: 'a negative or non-finite glow size is clamped rather than written',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addText('glowy', {
						x: 1,
						y: 1,
						w: 4,
						h: 1,
						glow: { size: -5, color: 'FFFF00', opacity: 0.5 },
					})
				})
				return readEntry(zip, SLIDE_XML)
			})
			assert(xml.includes('<a:glow rad="0">'), `a negative radius clamps to zero; got: ${xml}`)
			assert(codes.includes('glow/size-out-of-range'), 'and the caller is told; got ' + JSON.stringify(codes))
		},
	},
	{
		name: 'a NaN glow size collapses to zero rather than writing rad="NaN"',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('glowy', { x: 1, y: 1, w: 4, h: 1, glow: { size: Number.NaN, color: 'FFFF00' } })
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assert(!xml.includes('NaN'), `no NaN may reach the attribute; got: ${xml}`)
		},
	},
	{
		// The chart definer hand-rolled ~20 of these as silent `Array.includes` tests.
		name: 'an unrecognized chart enum option is reported before it falls back',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(async () => {
				const { zip } = await build((p) => {
					p.addSlide().addChart([{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }], {
						type: ChartType.bar,
						x: 1,
						y: 1,
						w: 6,
						h: 3,
						showLegend: true,
						legendPos: 'middle',
					})
				})
				return chartXml(zip)
			})
			assert(!xml.includes('middle'), `the bad value must not reach the part; got: ${xml}`)
			assert(xml.includes('<c:legendPos val="r"/>'), `it falls back to the default; got: ${xml}`)
			assert(codes.includes('chart/invalid-option-value'), 'and the caller is told; got ' + JSON.stringify(codes))
		},
	},
])
