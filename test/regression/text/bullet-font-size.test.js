import { setDiagnosticHandler, defineRegressionSuite, build, readEntry, assert } from '../../helpers.js'

// Custom bullet glyph font and size. Authored decks emit
// `<a:buFont typeface="Wingdings"/>` for symbol bullets and `<a:buSzPct/>` values
// other than 100% for resized glyphs; neither was previously controllable.

async function getPPr(zip) {
	const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
	const m = xml.match(/<a:pPr[^>]*\/?>(?:[\s\S]*?<\/a:pPr>)?/)
	if (!m) throw new Error('no <a:pPr> found in slide1.xml; xml=' + xml)
	return { xml, ppr: m[0] }
}

defineRegressionSuite('Bullet glyph font and size', [
	{
		name: 'bullet.fontFace emits <a:buFont/> between <a:buSzPct/> and <a:buChar/>',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('item', {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					bullet: { characterCode: 'F0E0', fontFace: 'Wingdings', size: 100 },
				})
			})
			const { ppr } = await getPPr(zip)
			assert(
				/<a:buSzPct val="100000"\/><a:buFont typeface="Wingdings"\/><a:buChar char="&#xF0E0;"\/>/.test(ppr),
				'expected buSzPct → buFont(Wingdings) → buChar ordering; got: ' + ppr
			)
		},
	},
	{
		name: 'bullet.size emits scaled <a:buSzPct/> (thousandths of a percent)',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('item', { x: 1, y: 1, w: 4, h: 1, bullet: { size: 80 } })
			})
			const { ppr } = await getPPr(zip)
			assert(/<a:buSzPct val="80000"\/>/.test(ppr), 'expected <a:buSzPct val="80000"/>; got: ' + ppr)
		},
	},
	{
		name: 'out-of-range bullet.size warns and clamps to the nearest bound',
		fn: async () => {
			const warnings = []
			setDiagnosticHandler((d) => warnings.push(d.message))
			let ppr
			try {
				const { zip } = await build((p) => {
					const s = p.addSlide()
					s.addText('item', { x: 1, y: 1, w: 4, h: 1, bullet: { size: 500 } })
				})
				ppr = (await getPPr(zip)).ppr
			} finally {
				setDiagnosticHandler(null)
			}
			// 25-400% is ST_TextBulletSizePercent's range, and a finite value outside it has a
			// nearest legal neighbour, so it moves there and says so. Emitting nothing instead
			// resized the glyph to whatever the list style inherits, which is a discarded
			// request reported as a warning (docs/diagnostics.md, "Warn or throw?").
			assert(
				/<a:buSzPct val="400000"\/>/.test(ppr),
				'expected <a:buSzPct val="400000"/> for a clamped size; got: ' + ppr
			)
			assert(
				warnings.some((w) => w.includes('bullet.size')),
				'expected a diagnostic mentioning bullet.size; got: ' + JSON.stringify(warnings)
			)
		},
	},
	{
		name: 'bullet.size below the range clamps up to 25%',
		fn: async () => {
			setDiagnosticHandler(() => {})
			let ppr
			try {
				const { zip } = await build((p) => {
					p.addSlide().addText('item', { x: 1, y: 1, w: 4, h: 1, bullet: { size: 1 } })
				})
				ppr = (await getPPr(zip)).ppr
			} finally {
				setDiagnosticHandler(null)
			}
			assert(/<a:buSzPct val="25000"\/>/.test(ppr), 'expected <a:buSzPct val="25000"/>; got: ' + ppr)
		},
	},
	{
		name: 'a bullet.size that is not a number throws rather than emitting val="NaN"',
		fn: async () => {
			let err
			try {
				await build((p) => {
					p.addSlide().addText('item', { x: 1, y: 1, w: 4, h: 1, bullet: { size: Number.NaN } })
				})
			} catch (e) {
				err = e
			}
			assert(err && err.code === 'percent/non-finite', 'expected percent/non-finite; got: ' + String(err && err.code))
		},
	},
	{
		name: 'bullet.fontFace applies to numbered bullets in place of the +mj-lt default',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('item', { x: 1, y: 1, w: 4, h: 1, bullet: { type: 'number', fontFace: 'Arial' } })
			})
			const { ppr } = await getPPr(zip)
			assert(
				/<a:buFont typeface="Arial"\/><a:buAutoNum/.test(ppr),
				'expected buFont(Arial) before buAutoNum; got: ' + ppr
			)
			assert(!/\+mj-lt/.test(ppr), 'expected custom font to replace +mj-lt default; got: ' + ppr)
		},
	},
	{
		name: 'object bullet without fontFace/size emits neither buSzPct nor buFont',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('item', { x: 1, y: 1, w: 4, h: 1, bullet: { characterCode: '25BA' } })
			})
			const { ppr } = await getPPr(zip)
			assert(/<a:buChar char="&#x25BA;"\/>/.test(ppr), 'expected the glyph; got: ' + ppr)
			// Both are omitted rather than defaulted, so an unstyled bullet inherits its size
			// and face from the list style instead of being pinned to 100% of the body font.
			assert(!/<a:buSzPct/.test(ppr), 'expected no <a:buSzPct/> when size is not set; got: ' + ppr)
			assert(!/<a:buFont/.test(ppr), 'expected no <a:buFont/> when fontFace is not set; got: ' + ppr)
		},
	},
])
