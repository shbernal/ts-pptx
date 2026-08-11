// Fixture-driven edges for src/read/api/chrome.ts that an authored deck cannot show.
//
// chrome-read.test.js measures the write→read round-trip of the chrome model, but
// the writer authors a deliberately plain master: no `p:cSld/p:bg` of its own, and
// every placeholder carrying explicit geometry. Real PowerPoint decks are the other
// way round — the master owns the background and layout placeholders routinely omit
// `a:xfrm` to inherit the master's. So `SlideMaster.background` and the
// inherited-geometry reads are asserted here against genuine PowerPoint output.
//
// Fixtures used: mixed.pptx (a master with an explicit `p:bgPr` solid fill),
// theme-colors.pptx (the Ion theme — a master `p:bgRef`, plus layouts that define no
// background of their own).
//
// ---------------------------------------------------------------------------
// Why chrome.ts branch coverage stops around 64%
// ---------------------------------------------------------------------------
// It is the lowest branch number on the read side, and that is deliberate — see
// docs/testing.md "Branches that are not worth covering". Every branch still
// uncovered here is the false arm of a guard that a schema-valid package cannot
// take. Four groups, all verified against the ECMA-376 content models:
//
//   1. `part.dom.documentElement` null checks (`Theme.name`, `SlideLayout.type`,
//      `SlideMaster.colorMap`, `#themeElements`, both `#cSld`). A part whose XML
//      has no root element never reaches a getter — the load fails first.
//   2. Required children read as optional. `a:themeElements` (CT_OfficeStyleSheet),
//      `a:clrScheme` and `a:fontScheme` (CT_BaseStyles), `p:cSld` (CT_SlideMaster /
//      CT_SlideLayout), `p:spTree` (CT_CommonSlideData), `p:nvSpPr` (CT_Shape),
//      `p:cNvPr` and `p:nvPr` (CT_ShapeNonVisual) are all `minOccurs="1"`. Same for
//      the `?? null` slot fallbacks in `colorScheme`/`color`/`colorMap`: a
//      `a:clrScheme` must carry all 12 children and a `p:clrMap` all 12 attributes.
//   3. Re-null-checks of a `p:ph` that `placeholderShapes()` already matched
//      (`Placeholder.type`, `.idx`). Unreachable by construction — the filter is
//      the only way a Placeholder gets built.
//   4. Relationships assumed missing or dangling: a master with no `theme` rel, a
//      layout with no `slideMaster` rel, a `p:sldLayoutId` with no `r:id` or an
//      `r:id` resolving to no part. Such a package fails OPC validation.
//
// A scan of every fixture deck finds zero instances of any of them, and zero of
// the tolerance fallbacks either (an unnamed `p:cSld` or `p:cNvPr` → `''`). Reaching
// them means hand-building a broken package, which asserts nothing about how the
// reader handles PowerPoint's output: the metric moves and the guarantee does not.
//
// Two branches were *not* in that category, because the schema does allow the
// input — `p:sldLayoutIdLst` is `minOccurs="0"` on CT_SlideMaster and `p:txBody` is
// `minOccurs="0"` on CT_Shape. So `SlideMaster.layouts` → `[]` and
// `Placeholder.textFrame` → `null` are real contracts on decks PowerPoint can write,
// uncovered only because no fixture happens to be shaped that way. Both are asserted
// in the last describe below, against an authored deck patched into that shape — and
// each patched package is run past the schema validator, which is what separates
// these from the four groups above: the input is legal, so the contract is real.

import path from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { authorRead, schemaErrors, validatorInstalled } from './authored.js'
import { openFixture } from './corpus.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

async function open(name, ext = 'pptx') {
	return openFixture(`${name}.${ext}`)
}

/**
 * Author a one-slide deck and rewrite its `slideMaster1.xml` with `mutate`, leaving
 * the rest of the package — theme, layout, slide, rels — the writer's own. Returns
 * the patched bytes so the caller can both read them back and schema-validate them;
 * validating is the point here, since the claim under test is that the shape is
 * legal input rather than a hand-broken package.
 */
async function patchedMaster(mutate) {
	const { buf } = await authorRead((pres) => {
		pres.addSlide()
	})
	const zip = await JSZip.loadAsync(buf)
	const partName = 'ppt/slideMasters/slideMaster1.xml'
	const xml = await zip.file(partName).async('string')
	const patched = mutate(xml)
	assert(patched !== xml, 'the patch actually changed the master XML')
	zip.file(partName, patched)
	return zip.generateAsync({ type: 'uint8array' })
}

describe('SlideMaster.background — the master tier of the chain', () => {
	test("an explicit p:bgPr on the master reads as the master's own solid fill", async () => {
		const bg = (await open('mixed')).masters()[0].background
		assert(bg?.type === 'solid', "mixed.pptx's master authors a solid p:bgPr")
		assertEqual(bg.source, 'master', 'the source records the master tier, not slide/layout')
		assertEqual(bg.color?.effectiveHex, 'FFFFFF', 'the fill colour resolves')
	})

	test('a p:bgRef on the master keeps its idx and resolves through the theme', async () => {
		const bg = (await open('theme-colors')).masters()[0].background
		assert(bg?.type === 'themeRef', 'the Ion master authors a theme-indexed background')
		assertEqual(bg.source, 'master', 'sourced from the master')
		assertEqual(bg.idx, 1003, 'the raw bgRef index')
		assertEqual(bg.color?.effectiveHex, '1E5155', "the bgRef's own schemeClr resolves through the Ion palette")
	})

	test('a master defining no background of its own reads null, not a none-fill', async () => {
		// The writer authors the background onto the layout, leaving the master's
		// `p:cSld` without a `p:bg` — so the getter must distinguish "defines none"
		// (null) from "explicitly transparent" (`type: 'none'`).
		const { presentation } = await authorRead((pres) => {
			pres.defineSlideMaster({ title: 'BRANDED', background: { color: 'F1F2F3' } })
			pres.addSlide({ masterTitle: 'BRANDED' })
		})
		const master = presentation.masters()[0]
		assertEqual(master.background, null, 'no p:bg on the master → null')
		// …and the layout that does define one still reads it, so the null above is a
		// real absence rather than a broken lookup.
		const branded = master.layouts.find((l) => l.name === 'BRANDED')
		assertEqual(branded?.background?.type, 'solid', 'the layout tier still resolves')
	})
})

describe('SlideLayout — inherited tiers', () => {
	test('a layout defining no background of its own reads null', async () => {
		const layout = (await open('theme-colors')).masters()[0].layouts[0]
		assertEqual(layout.name, 'Title Slide', 'the first Ion layout')
		assertEqual(layout.background, null, 'PowerPoint layouts usually inherit the master background')
	})

	test('a layout resolves its theme through its master', async () => {
		const layout = (await open('theme-colors')).masters()[0].layouts[0]
		const theme = layout.theme
		assert(theme, 'layout.theme walks layout → master → theme')
		assertEqual(theme.name, 'Ion', 'the fixture theme')
		assertEqual(theme.partName, layout.master?.theme?.partName, 'the shorthand reaches the same part as the long walk')
	})

	test("the import-only layout @type reads PowerPoint's value", async () => {
		// chrome-read.test.js pins this as null on an authored deck (the writer emits
		// none); the imported side is where a real value shows up.
		const layout = (await open('theme-colors')).masters()[0].layouts[0]
		assertEqual(layout.type, 'title', 'an imported layout carries its @type')
	})
})

describe('Placeholder geometry — inherited rather than own', () => {
	test('a layout placeholder with no own a:xfrm reads null geometry on all four axes', async () => {
		const layouts = (await open('theme-colors')).masters()[0].layouts
		const inherited = layouts.flatMap((l) => l.placeholders).find((ph) => ph.type === 'dt')
		assert(inherited, 'the Ion layouts carry a date placeholder')
		// It still identifies itself — only the geometry is absent, because PowerPoint
		// lets it inherit the master placeholder's box.
		assertEqual(inherited.idx, '10', 'the p:ph idx still reads')
		assert(inherited.name.length > 0, 'the shape name still reads')
		assert(typeof inherited.id === 'number', 'the drawing id still reads')
		assertEqual(inherited.left, null, 'no own a:xfrm → null left')
		assertEqual(inherited.top, null, 'null top')
		assertEqual(inherited.width, null, 'null width')
		assertEqual(inherited.height, null, 'null height')
	})
})

describe('optional children the fixtures never carry', () => {
	// The two branches the header calls out: schema-legal input that no committed
	// deck happens to be shaped like. Each is patched into an authored master and
	// validated, so the assertion pins a contract rather than a guard's false arm.

	// `p:sldLayoutIdLst` is minOccurs="0" on CT_SlideMaster: a master may list no
	// layouts at all. The layout part and its relationship stay in the package —
	// only the listing goes — since an unreferenced part is still legal, and that
	// keeps the empty read attributable to the missing list rather than to a
	// half-dismantled package.
	const withoutLayoutIdLst = () => patchedMaster((xml) => xml.replace(/<p:sldLayoutIdLst>.*?<\/p:sldLayoutIdLst>/s, ''))

	test('a master with no p:sldLayoutIdLst reads an empty layout list', async () => {
		const presentation = await Presentation.load(await withoutLayoutIdLst())
		const master = presentation.masters()[0]
		assert(master, 'the master itself still loads')
		assertEqual(master.layouts.length, 0, 'no p:sldLayoutIdLst → [], not a throw')
		// The layout part is still there and still bound to the slide, so the empty
		// list above is the master declining to enumerate — not the layout going missing.
		assertEqual(presentation.slides[0].layout?.name, 'DEFAULT', "the slide's own layout rel is unaffected")
	})

	test.skipIf(!validatorInstalled)('…and that master is schema-valid', async () => {
		assertEqual((await schemaErrors(await withoutLayoutIdLst())).length, 0, 'p:sldLayoutIdLst is genuinely optional')
	})

	// `p:txBody` is minOccurs="0" on CT_Shape. Two placeholders are spliced in, one
	// with a text body and one without, so the null reads as a real absence rather
	// than a lookup that fails for every shape.
	const PH_WITHOUT_TXBODY =
		'<p:sp><p:nvSpPr><p:cNvPr id="9" name="Bodyless Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
		'<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>' +
		'<p:spPr><a:xfrm><a:off x="838200" y="1825625"/><a:ext cx="10515600" cy="4351338"/></a:xfrm></p:spPr></p:sp>'
	const PH_WITH_TXBODY =
		'<p:sp><p:nvSpPr><p:cNvPr id="10" name="Prompted Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
		'<p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr>' +
		'<p:spPr><a:xfrm><a:off x="838200" y="1825625"/><a:ext cx="10515600" cy="4351338"/></a:xfrm></p:spPr>' +
		'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Master prompt</a:t></a:r></a:p></p:txBody></p:sp>'

	const withBodylessPlaceholder = () =>
		patchedMaster((xml) => xml.replace('</p:spTree>', `${PH_WITHOUT_TXBODY}${PH_WITH_TXBODY}</p:spTree>`))

	test('a placeholder with no p:txBody reads a null textFrame', async () => {
		const master = (await Presentation.load(await withBodylessPlaceholder())).masters()[0]
		const [bodyless, prompted] = master.placeholders
		assertEqual(master.placeholders.length, 2, 'both spliced placeholders are matched by the p:ph filter')
		// It is a fully-formed placeholder otherwise — only the text body is absent.
		assertEqual(bodyless.name, 'Bodyless Placeholder', 'the shape name still reads')
		assertEqual(bodyless.type, 'body', 'the p:ph type still reads')
		assertEqual(bodyless.idx, '1', 'the p:ph idx still reads')
		assertEqual(bodyless.left, 838200, 'the geometry still reads')
		assertEqual(bodyless.textFrame, null, 'no p:txBody → null')
		// The sibling proves the null is an absence, not a broken lookup.
		assertEqual(prompted.textFrame?.text, 'Master prompt', 'the placeholder that does carry one still reads it')
	})

	test.skipIf(!validatorInstalled)('…and that master is schema-valid', async () => {
		assertEqual((await schemaErrors(await withBodylessPlaceholder())).length, 0, 'p:txBody is genuinely optional')
	})
})
