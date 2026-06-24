import { defineRegressionSuite, build, readEntry, assert, assertIncludes, assertNotIncludes } from '../helpers.js'

// upstream-issue-1360: defineSlideMaster({ textStyles }) configures the shared slide master's
// per-level <p:txStyles> (titleStyle / bodyStyle / otherStyle). Previously the block was a fixed
// Office-default literal with no API to set nested bullet character, size, or color.

async function masterXml(zip) {
	return readEntry(zip, 'ppt/slideMasters/slideMaster1.xml')
}

defineRegressionSuite('Master text styles (#1360)', [
	{
		// Regression guard: a deck that does NOT set textStyles must keep the exact built-in default
		// <p:txStyles> (byte-for-byte literal), so existing decks are unaffected.
		name: 'default master txStyles unchanged when textStyles is unset',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({ title: 'PLAIN_MASTER' })
				p.addSlide({ masterName: 'PLAIN_MASTER' })
			})
			const xml = await masterXml(zip)
			// Spot-check the verbatim Office defaults remain.
			assertIncludes(xml, '<p:titleStyle>', 'titleStyle present')
			assertIncludes(xml, '<a:lvl1pPr algn="ctr" defTabSz="914400"', 'default title lvl1 verbatim')
			assertIncludes(xml, '<a:defRPr sz="4400" kern="1200">', 'default title size 44pt')
			assertIncludes(xml, '<a:lvl1pPr marL="342900" indent="-342900" algn="l"', 'default body lvl1 verbatim')
			assertIncludes(xml, '<a:buChar char="•"/>', 'default body lvl1 bullet glyph')
			assertIncludes(xml, '<a:defPPr><a:defRPr lang="en-US"/></a:defPPr>', 'otherStyle defPPr preserved')
		},
	},
	{
		name: 'body level overrides emit configured size, color, and bullet',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'CONFIGURED_MASTER',
					textStyles: {
						body: [
							{
								fontSize: 24,
								color: 'C00000',
								bold: true,
								italic: true,
								bullet: { characterCode: '25AA', fontFace: 'Arial' },
							},
						],
					},
				})
				p.addSlide({ masterName: 'CONFIGURED_MASTER' })
			})
			const xml = await masterXml(zip)
			assertIncludes(xml, 'sz="2400" b="1" i="1"', 'lvl1 24pt bold italic')
			assertIncludes(xml, '<a:srgbClr val="C00000"/>', 'lvl1 red color')
			assertIncludes(xml, '<a:buChar char="&#x25AA;"/>', 'lvl1 custom bullet char')
			// Unconfigured deeper levels keep their defaults (e.g. lvl3 24pt default glyph •).
			assertIncludes(xml, '<a:lvl3pPr marL="1143000" indent="-228600"', 'lvl3 default geometry retained')
		},
	},
	{
		name: 'bullet:false suppresses the level bullet (a:buNone)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'NOBULLET_MASTER',
					textStyles: { body: [{ bullet: false }] },
				})
				p.addSlide({ masterName: 'NOBULLET_MASTER' })
			})
			const xml = await masterXml(zip)
			const lvl1 = xml.slice(xml.indexOf('<a:lvl1pPr marL="342900"'), xml.indexOf('<a:lvl2pPr'))
			assertIncludes(lvl1, '<a:buNone/>', 'lvl1 bullet suppressed')
			assertNotIncludes(lvl1, '<a:buChar', 'lvl1 has no buChar')
		},
	},
	{
		// The single shared master means textStyles is deck-wide; a later defineSlideMaster call's
		// group replaces the earlier one (last-wins per group).
		name: 'textStyles is deck-wide, last call wins per group',
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster({ title: 'M1', textStyles: { title: { fontSize: 30 }, body: [{ fontSize: 18 }] } })
				p.defineSlideMaster({ title: 'M2', textStyles: { body: [{ fontSize: 22 }] } })
				p.addSlide({ masterName: 'M1' })
			})
			const xml = await masterXml(zip)
			const bodyStyle = xml.slice(xml.indexOf('<p:bodyStyle>'), xml.indexOf('</p:bodyStyle>'))
			// title from M1 retained (30pt -> sz=3000); body replaced by M2 (22pt -> sz=2200).
			assertIncludes(
				xml.slice(xml.indexOf('<p:titleStyle>'), xml.indexOf('</p:titleStyle>')),
				'sz="3000"',
				'title from M1 retained'
			)
			assertIncludes(bodyStyle, 'sz="2200"', 'body lvl1 from M2 wins')
			assertNotIncludes(bodyStyle, 'sz="1800"', 'M1 body 18pt no longer present in bodyStyle')
		},
	},
	{
		name: 'levels beyond 9 are ignored with a warning',
		fn: async () => {
			const originalWarn = console.warn
			const warnings = []
			console.warn = (msg) => warnings.push(String(msg))
			try {
				const { zip } = await build((p) => {
					p.defineSlideMaster({
						title: 'OVERFLOW_MASTER',
						textStyles: { body: Array.from({ length: 11 }, () => ({ fontSize: 12 })) },
					})
					p.addSlide({ masterName: 'OVERFLOW_MASTER' })
				})
				const xml = await masterXml(zip)
				assertIncludes(xml, '<a:lvl9pPr', 'lvl9 still emitted')
				assert(!xml.includes('<a:lvl10pPr'), 'no lvl10 emitted')
				assert(
					warnings.some((w) => w.includes('only the first 9 are used')),
					'warned about level overflow: ' + warnings.join('|')
				)
			} finally {
				console.warn = originalWarn
			}
		},
	},
])
