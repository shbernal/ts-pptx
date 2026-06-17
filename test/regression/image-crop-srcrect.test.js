import { defineRegressionSuite, build, readEntry, selfClosingTags, xmlAttributes, assert } from '../helpers.js'

// `crop: { l, t, r, b }` emits an explicit OOXML <a:srcRect> (percentage edge insets) verbatim.
// Regression guard for two things that the schema fixture alone cannot catch:
//   1. the option must survive addImage's option assembly (it was once dropped by the allowlist
//      in gen-objects.ts, so no srcRect was emitted at all — yet the package stayed schema-valid);
//   2. the inset percentages must serialize in 1000ths of a percent (ST_Percentage: 100% = 100000).

const PNG_1x1 =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

async function srcRectFor(opts) {
	const { zip } = await build((p) => {
		const s = p.addSlide()
		s.addImage(opts)
	})
	const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
	const tag = selfClosingTags(xml, 'a:srcRect').find((t) => /[ltrb]="/.test(t))
	assert(tag, 'expected a populated a:srcRect element; got: ' + xml)
	const a = xmlAttributes(tag)
	return { l: +(a.l ?? 0), t: +(a.t ?? 0), r: +(a.r ?? 0), b: +(a.b ?? 0) }
}

defineRegressionSuite('Image explicit crop (srcRect percentage insets)', [
	{
		// The core regression: the crop option must reach the renderer and emit a srcRect at all.
		name: 'crop reaches the renderer and emits a populated srcRect',
		fn: async () => {
			const r = await srcRectFor({ data: PNG_1x1, x: 1, y: 1, w: 2, h: 2, crop: { l: 0, t: 0, r: 50, b: 50 } })
			assert(
				r.l === 0 && r.t === 0 && r.r === 50000 && r.b === 50000,
				`top-left quadrant: expected l=0 t=0 r=50000 b=50000; got ${JSON.stringify(r)}`
			)
		},
	},
	{
		// Omitted edges default to 0; non-round percentages round to the nearest 1000th.
		name: 'omitted edges default to 0 and fractional percents round to 1000ths',
		fn: async () => {
			const r = await srcRectFor({ data: PNG_1x1, x: 1, y: 1, w: 2, h: 2, crop: { l: 12.5, t: 33.333 } })
			assert(
				r.l === 12500 && r.t === 33333 && r.r === 0 && r.b === 0,
				`expected l=12500 t=33333 r=0 b=0; got ${JSON.stringify(r)}`
			)
		},
	},
	{
		// crop wins over sizing (mutually exclusive) and warns.
		name: 'crop overrides sizing and warns',
		fn: async () => {
			const warnings = []
			const orig = console.warn
			console.warn = (m) => warnings.push(String(m))
			try {
				const r = await srcRectFor({
					data: PNG_1x1,
					x: 1,
					y: 1,
					w: 2,
					h: 2,
					crop: { l: 25, r: 25 },
					sizing: { type: 'cover', w: 2, h: 2 },
				})
				assert(r.l === 25000 && r.r === 25000, `expected crop srcRect (l=r=25000); got ${JSON.stringify(r)}`)
			} finally {
				console.warn = orig
			}
			assert(
				warnings.some((w) => w.includes('mutually exclusive')),
				`expected a mutual-exclusion warning; got: ${JSON.stringify(warnings)}`
			)
		},
	},
	{
		// Out-of-range / overlapping insets fail loudly rather than emitting a degenerate srcRect.
		name: 'invalid insets throw (out of range, overlap, NaN)',
		fn: async () => {
			const cases = [
				{ crop: { l: -5 }, why: 'negative' },
				{ crop: { r: 150 }, why: 'over 100' },
				{ crop: { l: 60, r: 60 }, why: 'l+r >= 100' },
				{ crop: { t: 50, b: 50 }, why: 't+b >= 100' },
				{ crop: { l: NaN }, why: 'NaN' },
			]
			for (const c of cases) {
				let threw = false
				try {
					await srcRectFor({ data: PNG_1x1, x: 1, y: 1, w: 2, h: 2, crop: c.crop })
				} catch {
					threw = true
				}
				assert(threw, `expected crop ${c.why} (${JSON.stringify(c.crop)}) to throw`)
			}
		},
	},
])
