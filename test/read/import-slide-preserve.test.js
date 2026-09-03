// Tests for `Presentation.importSlide(source, index, { theme: 'preserve' })`.
//
// Contract under test: `preserve` flattens the imported slide's *source* theme
// into the slide XML (scheme colours + style-matrix fills baked to literal
// srgbClr, with colour transforms carried through) and binds the slide to this
// deck's existing master/layout — so the output is a single-theme file whose
// imported slides render with their original colours and no longer depend on
// which theme they resolve against. The default (no option) is unchanged: the
// whole source theme subgraph is copied across.
//
// ---------------------------------------------------------------------------
// Why src/read/oxml/theme.ts branch coverage stops around 94%
// ---------------------------------------------------------------------------
// This file (with import-shape.test.js for the two `flattenShape`-only passes and
// import-slide-restyle.test.js for `remapLiteralColors`) is where theme.ts's
// branches are pinned. Every arm still uncovered is the arm a schema-valid package
// cannot take — see docs/testing.md "Branches that are not worth covering", and
// chrome-read-edges.test.js for the same treatment of chrome.ts. Four groups, each
// verified against the ECMA-376 content models rather than assumed:
//
//   1. Required attributes and children read as optional. `@val` on `a:srgbClr`,
//      `a:sysClr` and `a:schemeClr`; `@idx` on `a:CT_StyleMatrixReference`
//      (`use="required"`, `xsd:unsignedInt`, so it is never absent, empty, or
//      non-numeric); all 12 attributes of `a:CT_ColorMapping`; the single colour
//      child of `a:CT_Color`; `a:bodyPr` on `a:CT_TextBody`; `p:nvPr` on
//      `p:CT_ShapeNonVisual`; and all four of `a:lnRef`/`a:fillRef`/`a:effectRef`/
//      `a:fontRef` on `a:CT_ShapeStyle` — `materializeStyleRefs` is the only caller
//      that reaches those three guards, and it only ever passes a `p:style`'s own
//      children.
//   2. Re-checks of something the caller already established. `p:nvPr` looked up
//      again on a shape whose `p:nvSpPr` just matched; the parent of a `p:style`
//      found by descendant search; `resolveThemeFont(null, …)`, whose three callers
//      each guard the typeface first (a type-narrowing candidate, not a test).
//   3. Guards mutually exclusive with their caller. `applyInheritedBackground`
//      returns early when the slide owns a `p:bg` — but `#importSlidePreserve`
//      supplies `ctx.inheritedBackground` only when it does *not* (presentation.ts
//      `#effectiveBackground`), and `p:cSld` is `minOccurs="1"`, so the two
//      conditions cannot both hold. Likewise `placeholderInheritedListStyles`'
//      layout tier: `resolveSlideThemeParts` reaches the master *through* the
//      layout, so a null `layoutRoot` always implies a null `masterRoot`.
//      (The reverse — a layout that resolves to no master — is legal and is
//      covered; see below.)
//   4. Content models with no room for another case. The `a:effectStyle` child
//      dispatch has an implicit else that needs a fourth child type, but
//      `a:CT_EffectStyleItem` is exactly `effectLst`/`scene3d`/`sp3d`. And
//      `childElements` skipping a non-element node needs a text or comment node
//      among a colour element's children, i.e. pretty-printed XML; the fixtures are
//      minified, and covering it would assert nothing about deck semantics.
//
// What is *not* in those groups is asserted here, because the schema permits the
// input and only the fixtures happened not to carry it. Chief among them: both
// links a slide needs to reach its style chain are listed as relationships a part
// is *permitted* to have, not required to (ECMA-376 Part 1 §13.3.8 slide →
// slideLayout, §13.3.9 slideLayout → slideMaster). A slide missing the first
// resolves to neither root and every placeholder pass no-ops; a layout missing the
// second leaves the chain with a first tier and no second. Both packages validate,
// so the `!layoutRoot && !masterRoot` family, once written off as "structurally
// unreachable", is ordinary legal input. The same goes for a placeholder with no
// `p:txBody` and a master with no `p:txStyles` (both `minOccurs="0"`), a style-matrix
// `@idx` of 0 or past the end of its list, a theme stating its `clrScheme` in
// `a:prstClr`, and a `clrScheme`/`clrMap` that collides two slots or two tokens.
//
// Every one of those sources is spliced into a PowerPoint-authored deck and run
// past the schema validator, both before and after import — which is what makes
// "PowerPoint could write this" a check rather than an assertion. A splice the
// validator rejects belongs in the four groups above, not in a test.

import { readFile } from 'node:fs/promises'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual, partXml } from '../helpers.js'
import { validateBuf, validatorInstalled } from '../validator.js'
import { fixturePath, openFixture } from './corpus.js'
import { resolveSingle } from './opc.js'

const SLIDE_LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'

/** The serialized XML of a part, by partname, from saved package bytes. */
function countParts(opc, re) {
	return [...opc.parts.keys()].filter((n) => re.test(n)).length
}

/** A 1×1 transparent PNG, for a synthetic decoration picture. */
const PNG_1x1 = Buffer.from(
	'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
	'base64'
)

/**
 * mixed.pptx with a `p:pic` decoration spliced onto its slideMaster1 shape tree,
 * pointing at a fresh media part — so `carryMasterGraphics` has a picture (with a
 * relationship to rewrite) to carry. Returns the rebuilt package bytes.
 */
async function deckWithMasterPicture() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const pic =
		'<p:pic><p:nvPicPr><p:cNvPr id="987" name="CarryLogo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
		'<p:blipFill><a:blip r:embed="rId999"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
		'<p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="500" cy="500"/></a:xfrm>' +
		'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
	const master = (await zip.file('ppt/slideMasters/slideMaster1.xml').async('string')).replace(
		'</p:grpSpPr>',
		`</p:grpSpPr>${pic}`
	)
	zip.file('ppt/slideMasters/slideMaster1.xml', master)

	const rels = (await zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels').async('string')).replace(
		'</Relationships>',
		'<Relationship Id="rId999" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/carrytest.png"/></Relationships>'
	)
	zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', rels)

	zip.file('ppt/media/carrytest.png', PNG_1x1)
	const ct = (await zip.file('[Content_Types].xml').async('string')).replace(
		'</Types>',
		'<Override PartName="/ppt/media/carrytest.png" ContentType="image/png"/></Types>'
	)
	zip.file('[Content_Types].xml', ct)

	return zip.generateAsync({ type: 'uint8array' })
}

/**
 * mixed.pptx with slideLayout1's ctrTitle `a:xfrm` removed, so an imported
 * slide1 ctrTitle (which carries no own geometry) must inherit its geometry from
 * the slideMaster title placeholder instead of the layout. Returns package bytes.
 */
async function deckMixedNoLayoutCtrTitleXfrm() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const layout = (await zip.file('ppt/slideLayouts/slideLayout1.xml').async('string')).replace(
		'<a:xfrm><a:off x="990600" y="1828800"/><a:ext cx="7772400" cy="1143000"/></a:xfrm>',
		''
	)
	zip.file('ppt/slideLayouts/slideLayout1.xml', layout)
	return zip.generateAsync({ type: 'uint8array' })
}

/** mixed.pptx with an explicit `sz` on slide1's first ctrTitle run. Returns package bytes. */
async function deckMixedWithExplicitTitleSize() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const slide = (await zip.file('ppt/slides/slide1.xml').async('string')).replace(
		'<a:rPr lang="fr-FR" dirty="0"/><a:t>Data </a:t>',
		'<a:rPr lang="fr-FR" sz="4444" dirty="0"/><a:t>Data </a:t>'
	)
	zip.file('ppt/slides/slide1.xml', slide)
	return zip.generateAsync({ type: 'uint8array' })
}

/** empty.pptx with its master `p:bgRef` switched to an idx below 1000 (the regular `fillStyleLst`, not `bgFillStyleLst`). Returns package bytes. */
async function deckEmptyBgRefFillStyleLst() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('empty')))
	const master = (await zip.file('ppt/slideMasters/slideMaster1.xml').async('string')).replace(
		'<p:bgRef idx="1001">',
		'<p:bgRef idx="1">'
	)
	zip.file('ppt/slideMasters/slideMaster1.xml', master)
	return zip.generateAsync({ type: 'uint8array' })
}

/** empty.pptx with its master `p:bgRef` idx zeroed out, so materializeBackground falls back to a:noFill. Returns package bytes. */
async function deckEmptyBgRefIdxZero() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('empty')))
	const master = (await zip.file('ppt/slideMasters/slideMaster1.xml').async('string')).replace(
		'<p:bgRef idx="1001">',
		'<p:bgRef idx="0">'
	)
	zip.file('ppt/slideMasters/slideMaster1.xml', master)
	return zip.generateAsync({ type: 'uint8array' })
}

/**
 * mixed.pptx with slide1's `slideLayout` relationship removed. ECMA-376 Part 1
 * §13.3.8 lists that relationship among the ones a Slide part is *permitted* to
 * have, not required to — so the slide resolves to no layout, and (because the
 * master is reached through the layout) to no master either. Returns package bytes.
 */
async function deckMixedSlideNoLayoutRel() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const rels = (await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string')).replace(
		/<Relationship[^>]*slideLayout[^>]*\/>/,
		''
	)
	zip.file('ppt/slides/_rels/slide1.xml.rels', rels)
	return zip.generateAsync({ type: 'uint8array' })
}

/**
 * mixed.pptx with every `a:clrScheme` slot restated as `a:prstClr` — a legal
 * `a:CT_Color` child that names no 6-hex RGB. The master `p:clrMap` still resolves
 * a token to a slot; that slot simply has no literal behind it. Returns package bytes.
 */
async function deckMixedUnreadableColorScheme() {
	const slots = ['dk1', 'lt1', 'dk2', 'lt2']
		.concat([1, 2, 3, 4, 5, 6].map((n) => `accent${n}`))
		.concat(['hlink', 'folHlink'])
	// `a:scrgbClr` is the one `a:EG_ColorChoice` member this reader deliberately does not
	// resolve: its channels are percentages of a colour space the schema does not pin down,
	// so linear-light and sRGB-encoded readings differ by a gamma curve and neither is
	// derivable. That makes it the honest way to build "a slot with nothing literal behind
	// it". It used to be `a:prstClr val="black"`, which stopped working as a stand-in the
	// day the preset table landed and that name started resolving to `000000`.
	const scheme = `<a:clrScheme name="Preset">${slots
		.map((s) => `<a:${s}><a:scrgbClr r="0" g="0" b="0"/></a:${s}>`)
		.join('')}</a:clrScheme>`
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const theme = (await zip.file('ppt/theme/theme1.xml').async('string')).replace(
		/<a:clrScheme[\s\S]*?<\/a:clrScheme>/,
		scheme
	)
	zip.file('ppt/theme/theme1.xml', theme)
	return zip.generateAsync({ type: 'uint8array' })
}

/**
 * empty.pptx with its master `p:bgRef` set to exactly 1000 — the first index of
 * the `bgFillStyleLst` range, which offsets by 1000 and so selects entry 0 of a
 * 1-based list. Legal `unsignedInt`; resolves to nothing. Returns package bytes.
 */
async function deckEmptyBgRefIdx1000() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('empty')))
	const master = (await zip.file('ppt/slideMasters/slideMaster1.xml').async('string')).replace(
		'<p:bgRef idx="1001">',
		'<p:bgRef idx="1000">'
	)
	zip.file('ppt/slideMasters/slideMaster1.xml', master)
	return zip.generateAsync({ type: 'uint8array' })
}

/** mixed.pptx with every slide5 `a:lnRef` zeroed — the "no line from the style matrix" ref PowerPoint writes constantly. Returns package bytes. */
async function deckMixedLnRefIdxZero() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const slide = (await zip.file('ppt/slides/slide5.xml').async('string')).replace(
		/<a:lnRef idx="\d+"/g,
		'<a:lnRef idx="0"'
	)
	zip.file('ppt/slides/slide5.xml', slide)
	return zip.generateAsync({ type: 'uint8array' })
}

/**
 * mixed.pptx with every slide5 `a:lnRef`/`a:effectRef` pointed past the end of the
 * theme's three-entry `lnStyleLst`/`effectStyleLst`. `@idx` is an unbounded
 * `unsignedInt`, so this is legal input that simply resolves to nothing.
 * Returns package bytes.
 */
async function deckMixedStyleRefOutOfRange() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const slide = (await zip.file('ppt/slides/slide5.xml').async('string'))
		.replace(/<a:lnRef idx="\d+"/g, '<a:lnRef idx="9"')
		.replace(/<a:effectRef idx="\d+"/g, '<a:effectRef idx="9"')
	zip.file('ppt/slides/slide5.xml', slide)
	return zip.generateAsync({ type: 'uint8array' })
}

/**
 * mixed.pptx with slide1's ctrTitle given a `phClr` fill of its own. `phClr` is a
 * member of `ST_SchemeColorVal` like any other token, but it names "the colour
 * this style-matrix entry is being filled with" — outside a `fmtScheme` entry
 * there is nothing to substitute. Returns package bytes.
 */
async function deckMixedStraySchemePhClr() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const slide = (await zip.file('ppt/slides/slide1.xml').async('string')).replace(
		'<p:spPr/>',
		'<p:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></p:spPr>'
	)
	zip.file('ppt/slides/slide1.xml', slide)
	return zip.generateAsync({ type: 'uint8array' })
}

/**
 * mixed.pptx with theme1's first `fillStyleLst` entry turned into a two-stop
 * gradient whose second stop is a *non*-`phClr` scheme colour. The stock Office
 * themes happen never to do this, but nothing in `a:CT_StyleMatrix` forbids it —
 * a style-matrix entry may pin part of its fill to a fixed scheme slot.
 *
 * Every slide5 shape carrying `a:fillRef idx="1"` also carries its own `spPr`
 * fill, which wins outright, so one of them has that fill stripped — otherwise the
 * entry is never materialized onto a shape at all. Returns package bytes.
 */
async function deckMixedFillStyleFixedSchemeStop() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const theme = (await zip.file('ppt/theme/theme1.xml').async('string')).replace(
		'<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>',
		'<a:fillStyleLst><a:gradFill rotWithShape="1"><a:gsLst>' +
			'<a:gs pos="0"><a:schemeClr val="phClr"/></a:gs>' +
			'<a:gs pos="100000"><a:schemeClr val="tx1"/></a:gs>' +
			'</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
	)
	zip.file('ppt/theme/theme1.xml', theme)

	const slide = (await zip.file('ppt/slides/slide5.xml').async('string')).replace(
		'<a:off x="566555" y="2009998"/><a:ext cx="2160240" cy="1152128"/></a:xfrm>' +
			'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
			'<a:solidFill><a:schemeClr val="tx2"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr></a:solidFill>',
		'<a:off x="566555" y="2009998"/><a:ext cx="2160240" cy="1152128"/></a:xfrm>' +
			'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
	)
	zip.file('ppt/slides/slide5.xml', slide)

	return zip.generateAsync({ type: 'uint8array' })
}

/** mixed.pptx with slide1 given its own literal p:bg, so it is not the one inherited from the master. Returns package bytes. */
async function deckMixedSlideOwnBackground() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const slide = (await zip.file('ppt/slides/slide1.xml').async('string')).replace(
		'<p:cSld><p:spTree>',
		'<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></p:bgPr></p:bg><p:spTree>'
	)
	zip.file('ppt/slides/slide1.xml', slide)
	return zip.generateAsync({ type: 'uint8array' })
}

/** mixed.pptx with slideMaster1's p:bg stripped, so no slide/layout/master in the chain defines a background. Returns package bytes. */
async function deckMixedNoBackgroundAnywhere() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const master = (await zip.file('ppt/slideMasters/slideMaster1.xml').async('string')).replace(
		'<p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>',
		''
	)
	zip.file('ppt/slideMasters/slideMaster1.xml', master)
	return zip.generateAsync({ type: 'uint8array' })
}

/** mixed.pptx with slide5's first p:style effectRef pointed at fmtScheme effectStyleLst entry 3 (effectLst + scene3d + sp3d), instead of the unresolved idx=0. Returns package bytes. */
async function deckMixedEffectRefMaterialized() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const slide = (await zip.file('ppt/slides/slide5.xml').async('string')).replace(
		'<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>',
		'<a:effectRef idx="3"><a:schemeClr val="accent1"/></a:effectRef>'
	)
	zip.file('ppt/slides/slide5.xml', slide)
	return zip.generateAsync({ type: 'uint8array' })
}

/**
 * mixed.pptx with slide1's subTitle placeholder stripped of its `p:txBody`
 * (`minOccurs="0"` on `p:CT_Shape`, so a picture/chart/table placeholder that
 * holds no text is ordinary PowerPoint output). Returns package bytes.
 */
async function deckMixedPlaceholderNoTxBody() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const slide = (await zip.file('ppt/slides/slide1.xml').async('string')).replace(
		'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody>',
		''
	)
	zip.file('ppt/slides/slide1.xml', slide)
	return zip.generateAsync({ type: 'uint8array' })
}

/**
 * mixed.pptx with slideLayout1's ctrTitle placeholder stripped of its `p:txBody`,
 * so the layout tier of the style chain contributes no `a:lstStyle` at all and
 * resolution falls through to the master. Returns package bytes.
 */
async function deckMixedLayoutPlaceholderNoTxBody() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const layout = (await zip.file('ppt/slideLayouts/slideLayout1.xml').async('string')).replace(
		'<p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr/></a:lvl1pPr></a:lstStyle>' +
			'<a:p><a:r><a:rPr lang="fr-FR"/><a:t>Cliquez pour modifier le style du titre du masque</a:t></a:r></a:p></p:txBody>',
		''
	)
	zip.file('ppt/slideLayouts/slideLayout1.xml', layout)
	return zip.generateAsync({ type: 'uint8array' })
}

/** mixed.pptx with a second paragraph appended to slide1's ctrTitle, at the same (default) level as the first. Returns package bytes. */
async function deckMixedTwoParagraphsSameLevel() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const slide = (await zip.file('ppt/slides/slide1.xml').async('string')).replace(
		'<a:endParaRPr lang="en-US" dirty="0"/></a:p>',
		'<a:endParaRPr lang="en-US" dirty="0"/></a:p><a:p><a:r><a:rPr lang="fr-FR"/><a:t>Second line</a:t></a:r></a:p>'
	)
	zip.file('ppt/slides/slide1.xml', slide)
	return zip.generateAsync({ type: 'uint8array' })
}

/**
 * mixed.pptx with slide1's ctrTitle fixing its own run properties at the two
 * levels below the run itself: paragraph 1 gets an `a:pPr/a:defRPr` carrying both
 * a `sz` and a `solidFill`, and the text body's `a:lstStyle` gets a level-1
 * `a:defRPr sz`. A second paragraph is appended that sets no `a:pPr`, so it
 * resolves against the `a:lstStyle` tier alone. Returns package bytes.
 */
async function deckMixedSlideFixesRunProps() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const slide = (await zip.file('ppt/slides/slide1.xml').async('string'))
		.replace(
			'<p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/>',
			'<p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle><a:lvl1pPr><a:defRPr sz="1111"/></a:lvl1pPr></a:lstStyle>'
		)
		.replace(
			'<a:p><a:r><a:rPr lang="fr-FR" dirty="0"/><a:t>Data </a:t>',
			'<a:p><a:pPr><a:defRPr sz="2222"><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:defRPr></a:pPr>' +
				'<a:r><a:rPr lang="fr-FR" dirty="0"/><a:t>Data </a:t>'
		)
		.replace(
			'<a:endParaRPr lang="en-US" dirty="0"/></a:p>',
			'<a:endParaRPr lang="en-US" dirty="0"/></a:p><a:p><a:r><a:rPr lang="fr-FR"/><a:t>Second line</a:t></a:r></a:p>'
		)
	zip.file('ppt/slides/slide1.xml', slide)
	return zip.generateAsync({ type: 'uint8array' })
}

/**
 * mixed.pptx cut down to a source whose placeholder style chain supplies nothing:
 * the master's `p:txStyles` removed (`minOccurs="0"` on `p:CT_SlideMaster`),
 * slideLayout1's ctrTitle `a:lstStyle` emptied, and slide1 given an extra picture
 * placeholder whose `type`/`idx` match no layout or master placeholder. Returns
 * package bytes.
 */
async function deckMixedInheritsNothing() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('mixed')))
	const master = (await zip.file('ppt/slideMasters/slideMaster1.xml').async('string')).replace(
		/<p:txStyles>[\s\S]*<\/p:txStyles>/,
		''
	)
	zip.file('ppt/slideMasters/slideMaster1.xml', master)

	const layout = (await zip.file('ppt/slideLayouts/slideLayout1.xml').async('string')).replace(
		'<p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr/></a:lvl1pPr></a:lstStyle>',
		'<p:txBody><a:bodyPr/><a:lstStyle/>'
	)
	zip.file('ppt/slideLayouts/slideLayout1.xml', layout)

	const orphan =
		'<p:sp><p:nvSpPr><p:cNvPr id="77" name="Picture Placeholder 7"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
		'<p:nvPr><p:ph type="pic" idx="7"/></p:nvPr></p:nvSpPr><p:spPr/>' +
		'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Picture</a:t></a:r></a:p></p:txBody></p:sp>'
	const slide = (await zip.file('ppt/slides/slide1.xml').async('string')).replace('</p:spTree>', `${orphan}</p:spTree>`)
	zip.file('ppt/slides/slide1.xml', slide)

	return zip.generateAsync({ type: 'uint8array' })
}

/** layout-placeholder-bodypr.pptx with its body placeholder's single paragraph replaced by five, at explicit lvl 0..4. Returns package bytes. */
async function deckMultiLevelBody() {
	const zip = await JSZip.loadAsync(await readFile(fixturePath('layout-placeholder-bodypr')))
	const levels = [0, 1, 2, 3, 4]
		.map((lvl) => (lvl === 0 ? '' : `<a:pPr lvl="${lvl}"/>`))
		.map((pPr, lvl) => `<a:p>${pPr}<a:r><a:rPr lang="en-US"/><a:t>L${lvl}</a:t></a:r></a:p>`)
		.join('')
	const slide = (await zip.file('ppt/slides/slide1.xml').async('string')).replace(
		'<a:p><a:r><a:rPr lang="en-US"/><a:t>Middle-anchored body</a:t></a:r></a:p>',
		levels
	)
	zip.file('ppt/slides/slide1.xml', slide)
	return zip.generateAsync({ type: 'uint8array' })
}

/** The index of the first `mixed` slide that uses scheme colours + a p:style. */
const THEMED_SLIDE_INDEX = 4 // slide5: 69 schemeClr, 18 p:style (Fusion theme: accent1=00E4A8, dk2=333399)

describe("Presentation.importSlide({ theme: 'preserve' })", () => {
	test('flattens a non-default PowerPoint theme source into literals when importing into a default-theme deck', async () => {
		const target = await openFixture('empty')
		const source = await openFixture('multi-theme')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		assert(!/schemeClr/.test(xml), 'no a:schemeClr token remains in the flattened slide')
		assert(/<a:srgbClr val="B01513"/.test(xml), 'source Ion accent1 flattened to its literal RGB')
		assert(/<a:srgbClr val="EA6312"/.test(xml), 'source Ion accent2 flattened to its literal RGB')
		assert(/<a:srgbClr val="E6B729"/.test(xml), 'source Ion accent3 line colour flattened to its literal RGB')
		assert(/<a:srgbClr val="54849A"/.test(xml), 'source Ion accent5 run colour flattened to its literal RGB')
		assert(
			/<a:srgbClr val="EA6312"><a:lumMod val="65000"\/><a:lumOff val="35000"\/><\/a:srgbClr>/.test(xml),
			'colour transforms are carried onto the flattened literal'
		)
		assert(/<a:fillRef idx="0"\/>/.test(xml), 'style fillRef neutralized after materialization')
		assert(/<a:lnRef idx="0"\/>/.test(xml), 'style lnRef neutralized after materialization')
	})

	test('flattens scheme colours to the resolved source RGB, carrying transforms', async () => {
		// Import mixed→mixed (equal slide size); slide5 binds to theme1 "Fusion".
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		assert(!/schemeClr/.test(xml), 'no a:schemeClr token remains in the flattened slide')
		assert(!/phClr/.test(xml), 'no phClr placeholder leaked from the style matrix')
		// accent1 → 00E4A8; tx2 → (clrMap tx2=dk2) → dk2 = 333399.
		assert(xml.includes('val="00E4A8"'), 'accent1 resolved to its Fusion RGB (00E4A8)')
		assert(xml.includes('val="333399"'), 'tx2 mapped through clrMap to dk2 RGB (333399)')
		// A scheme colour with a child transform keeps the transform on the literal.
		assert(/<a:srgbClr val="[0-9A-Fa-f]{6}"><a:lumMod/.test(xml), 'lumMod transform carried onto the resolved srgbClr')
	})

	test('materializes p:style fill/line/effect into spPr and neutralizes the refs', async () => {
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		// Every fill/line/effect ref is neutralized to idx="0" with no colour child…
		assert(/<a:fillRef idx="0"\/>/.test(xml), 'fillRef neutralized to idx="0" with no colour')
		assert(/<a:lnRef idx="0"\/>/.test(xml), 'lnRef neutralized to idx="0" with no colour')
		// …but the fontRef is left intact so its font re-binds to the destination theme.
		assert(/<a:fontRef idx="(major|minor|none)"/.test(xml), 'fontRef is preserved for the destination theme')
		// The shapes that were styled by reference now carry an explicit spPr fill.
		assert(
			/<p:spPr[ >][\s\S]*?<a:solidFill>/.test(xml) || /<a:gradFill/.test(xml),
			'an explicit spPr fill was materialized'
		)
	})

	test("carries the slide's effective background from the source master onto the slide", async () => {
		// mixed slides define no own p:bg; the master does (bgPr/solidFill schemeClr
		// bg1 → clrMap bg1=lt1 → lt1 = FFFFFF). Rebinding to the destination master
		// would otherwise drop it, so preserve must bake it onto the slide.
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		const bg = (xml.match(/<p:bg>[\s\S]*?<\/p:bg>/) ?? [''])[0]
		assert(bg, 'the imported slide carries an explicit p:cSld/p:bg')
		assert(!/schemeClr/.test(bg), 'the carried background holds no scheme colour')
		assert(/<a:srgbClr val="FFFFFF"\/>/.test(bg), 'the background is the resolved literal (bg1 → lt1 = FFFFFF)')
		// p:bg must sit before p:spTree inside p:cSld.
		assert(
			/<p:cSld[^>]*>\s*<p:bg>/.test(xml) || /<p:bg>[\s\S]*?<\/p:bg>\s*<p:spTree/.test(xml),
			'p:bg precedes p:spTree'
		)
	})

	test('materializes an inherited bgRef background into an explicit fill', async () => {
		// The empty deck's master uses a theme-indexed background (bgRef idx="1001").
		// preserve must resolve it through fmtScheme into a literal bgPr fill.
		const target = await openFixture('empty')
		const source = await openFixture('empty')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		const bg = (xml.match(/<p:bg>[\s\S]*?<\/p:bg>/) ?? [''])[0]
		assert(bg, 'the imported slide carries an explicit background')
		assert(!/bgRef/.test(bg), 'the theme-indexed bgRef was resolved away')
		assert(!/schemeClr/.test(bg) && !/phClr/.test(bg), 'the background fill is a literal, no scheme/placeholder colour')
		assert(/<p:bgPr>[\s\S]*<a:srgbClr /.test(bg), 'bgRef became a bgPr with a literal srgbClr fill')
	})

	test("bakes a placeholder run's inherited colour onto the run so a rebind cannot change it", async () => {
		// slide1's ctrTitle runs carry no own colour; the white/dark they render comes
		// from the source master titleStyle (defRPr/solidFill schemeClr tx2 → clrMap
		// tx2=dk2 → 333399). Rebinding to the destination master would drop that
		// inheritance, so preserve must write the resolved colour explicitly.
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'preserve' }) // slide1: ctrTitle
		const xml = await partXml(await target.save(), imported.partName)

		const sp = (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?ctrTitle[\s\S]*?<\/p:sp>/) ?? [''])[0]
		assert(sp, 'the imported slide still has its ctrTitle placeholder')
		const runs = [...sp.matchAll(/<a:r>[\s\S]*?<\/a:r>/g)].map((m) => m[0])
		assert(runs.length > 0, 'ctrTitle has runs')
		for (const run of runs) {
			assert(
				/<a:rPr[\s\S]*?<a:solidFill><a:srgbClr val="333399"\/><\/a:solidFill>/.test(run),
				'each ctrTitle run carries the resolved master titleStyle colour (333399) explicitly'
			)
			assert(!/schemeClr/.test(run), 'the baked run colour is a literal, not a scheme token')
		}
	})

	test('attaches to the destination master/layout without importing a new theme', async () => {
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const themesBefore = countParts(target.opc, /\/theme\/theme\d+\.xml$/)
		const mastersBefore = countParts(target.opc, /\/slideMasters\/slideMaster\d+\.xml$/)

		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'preserve' })
		const reopened = await Presentation.load(await target.save())
		const opc = reopened.opc

		assertEqual(countParts(opc, /\/theme\/theme\d+\.xml$/), themesBefore, 'preserve adds no new theme part')
		assertEqual(
			countParts(opc, /\/slideMasters\/slideMaster\d+\.xml$/),
			mastersBefore,
			'preserve adds no new master part'
		)

		// The imported slide binds to a layout that already exists in the target deck.
		const last = reopened.slides[reopened.slides.length - 1]
		const layout = resolveSingle(opc, last.partName, SLIDE_LAYOUT_REL)
		assert(layout && opc.part(layout), `imported slide binds to an existing destination layout (${layout})`)

		// No dangling internal relationships anywhere in the package.
		for (const partName of opc.parts.keys()) {
			if (partName.endsWith('.rels')) continue
			for (const rel of opc.relationshipsFor(partName)) {
				if (rel.targetMode === 'External') continue
				const t = opc.relationshipsFor(partName).resolveTarget(rel.id)
				assert(opc.part(t), `${partName} → ${rel.id} resolves to an existing part (${t})`)
			}
		}
		assertEqual(imported.index, reopened.slides.length - 1, 'imported slide is last')
	})

	test('preserve carries the slide media across (image slide stays intact)', async () => {
		// The image fixture's slide carries a picture; import it with preserve.
		const source = await openFixture('image')
		const picSlide = source.slides.findIndex((s) => s.shapes.some((sh) => sh.shapeType === 'picture'))
		assert(picSlide >= 0, 'image fixture has a slide with a picture')

		const target = await openFixture('image')
		const imported = target.importSlide(source, picSlide, { theme: 'preserve' })
		const reopened = await Presentation.load(await target.save())
		const last = reopened.slides[reopened.slides.length - 1]
		const pic = last.shapes.find((s) => s.shapeType === 'picture')
		assert(pic, 'imported slide still has its picture')
		assert(
			pic.imagePartName && reopened.opc.part(pic.imagePartName),
			`the picture's media part survives (${pic.imagePartName})`
		)
		assertEqual(imported.partName, last.partName, 'imported slide is the appended one')
	})

	test('carryMasterGraphics bakes source master/layout decorations onto the slide, behind its content', async () => {
		// mixed's slideMaster1 carries non-placeholder rectangles ("Rectangle 2".."8");
		// slide1's layout1 carries more (groups + rectangles). preserve drops the source
		// master, so by default those decorations vanish; carryMasterGraphics bakes them on.
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'preserve', carryMasterGraphics: true }) // slide1: ctrTitle
		const xml = await partXml(await target.save(), imported.partName)

		assert(xml.includes('name="Rectangle 2"'), 'a source-master decoration was baked onto the slide')
		// Decorations sit ahead of the slide's own content (document order == z-order).
		assert(
			xml.indexOf('name="Rectangle 2"') < xml.indexOf('ctrTitle'),
			'carried decoration precedes the slide placeholder'
		)
		// Flatten still ran over the carried shapes: no scheme token survives.
		assert(!/schemeClr/.test(xml), 'carried decorations were flattened to literal colours')
	})

	test('carryMasterGraphics carries no placeholder shapes from the master/layout', async () => {
		// The master/layout placeholders are inherited via the placeholder mechanism, not
		// baked as decorations; carry must add only non-placeholder shapes.
		const target = await openFixture('mixed')
		const plain = target.importSlide(await openFixture('mixed'), 0, { theme: 'preserve' })
		const withGfx = target.importSlide(await openFixture('mixed'), 0, { theme: 'preserve', carryMasterGraphics: true })
		const bytes = await target.save()
		const countPh = (xml) => (xml.match(/<p:ph[ />]/g) ?? []).length

		const plainPh = countPh(await partXml(bytes, plain.partName))
		const gfxPh = countPh(await partXml(bytes, withGfx.partName))
		assertEqual(gfxPh, plainPh, 'carry added no extra placeholder shapes')
	})

	test('without carryMasterGraphics, master/layout decorations are not carried (default)', async () => {
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'preserve' }) // no carry flag
		const xml = await partXml(await target.save(), imported.partName)
		assert(!xml.includes('name="Rectangle 2"'), 'no source-master decoration leaks in without the flag')
	})

	test('carryMasterGraphics still attaches to the destination master without importing a theme', async () => {
		const target = await openFixture('mixed')
		const themesBefore = countParts(target.opc, /\/theme\/theme\d+\.xml$/)
		const mastersBefore = countParts(target.opc, /\/slideMasters\/slideMaster\d+\.xml$/)

		target.importSlide(await openFixture('mixed'), 0, { theme: 'preserve', carryMasterGraphics: true })
		const reopened = await Presentation.load(await target.save())
		const opc = reopened.opc
		assertEqual(countParts(opc, /\/theme\/theme\d+\.xml$/), themesBefore, 'carry adds no new theme part')
		assertEqual(countParts(opc, /\/slideMasters\/slideMaster\d+\.xml$/), mastersBefore, 'carry adds no new master part')

		// No dangling internal relationships anywhere in the package.
		for (const partName of opc.parts.keys()) {
			if (partName.endsWith('.rels')) continue
			for (const rel of opc.relationshipsFor(partName)) {
				if (rel.targetMode === 'External') continue
				const t = opc.relationshipsFor(partName).resolveTarget(rel.id)
				assert(opc.part(t), `${partName} → ${rel.id} resolves to an existing part (${t})`)
			}
		}
	})

	test('carryMasterGraphics copies a decoration picture media and rewrites its relationship', async () => {
		// mixed's master has no picture decoration, so splice one (a p:pic on the master
		// spTree referencing a fresh media part) into a source deck, then carry it across.
		const source = await Presentation.load(await deckWithMasterPicture())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'preserve', carryMasterGraphics: true })
		const reopened = await Presentation.load(await target.save())

		const last = reopened.slides[reopened.slides.length - 1]
		assertEqual(imported.partName, last.partName, 'imported slide is the appended one')
		const pic = last.shapes.find((s) => s.shapeType === 'picture')
		assert(pic, 'the carried master picture lands as a slide picture')
		assert(
			pic.imagePartName && reopened.opc.part(pic.imagePartName),
			`its media part was copied across (${pic.imagePartName})`
		)

		const xml = await partXml(await reopened.save(), last.partName)
		assert(xml.includes('name="CarryLogo"'), 'the carried picture is the one we spliced onto the master')
		assert(!/r:embed="rId999"/.test(xml), 'the source rel id was rewritten to a fresh slide-local id')
	})

	test('bakes placeholder geometry inherited from the source layout onto the shape', async () => {
		// slide1's ctrTitle carries no own a:xfrm; its position/size are inherited from
		// slideLayout1's ctrTitle (off 990600,1828800; ext 7772400,1143000). Rebinding to
		// the destination master would drop that, so preserve must bake the layout xfrm on.
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'preserve' }) // slide1: ctrTitle, no own xfrm
		const xml = await partXml(await target.save(), imported.partName)

		const sp = (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?ctrTitle[\s\S]*?<\/p:sp>/) ?? [''])[0]
		assert(sp, 'the imported slide still has its ctrTitle placeholder')
		assert(
			/<a:xfrm><a:off x="990600" y="1828800"\/><a:ext cx="7772400" cy="1143000"\/><\/a:xfrm>/.test(sp),
			'ctrTitle carries the source layout1 geometry explicitly'
		)
	})

	test('bakes placeholder geometry from the source master when the layout lacks it', async () => {
		// With slideLayout1's ctrTitle xfrm stripped, the ctrTitle's geometry falls through
		// to slideMaster1's title placeholder (off 1023938,131763; ext 7793037,1143000).
		const source = await Presentation.load(await deckMixedNoLayoutCtrTitleXfrm())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		const sp = (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?ctrTitle[\s\S]*?<\/p:sp>/) ?? [''])[0]
		assert(
			/<a:off x="1023938" y="131763"\/><a:ext cx="7793037" cy="1143000"\/>/.test(sp),
			'ctrTitle inherits the master title geometry when the layout defines none'
		)
	})

	test('leaves a placeholder that already has its own a:xfrm untouched', async () => {
		// slide2's title placeholder defines its own xfrm (off 1115616,3200); explicit
		// geometry is not inherited, so preserve must not overwrite it.
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const imported = target.importSlide(source, 1, { theme: 'preserve' }) // slide2: title with own xfrm
		const xml = await partXml(await target.save(), imported.partName)

		const sp = (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?type="title"[\s\S]*?<\/p:sp>/) ?? [''])[0]
		assert(sp, 'the imported slide has its title placeholder')
		assert(/<a:off x="1115616" y="3200"\/>/.test(sp), 'the slide-owned title geometry is preserved')
		assert(!/x="1023938"/.test(sp) && !/x="990600"/.test(sp), 'no layout/master geometry was baked over it')
	})

	test('bakes placeholder-inherited run size onto runs that set none', async () => {
		// slide1's ctrTitle runs set no sz; the size comes from the master titleStyle lvl1
		// (defRPr sz="3200"). Rebinding would drop it, so preserve bakes sz onto each run.
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'preserve' }) // slide1: ctrTitle, runs carry no sz
		const xml = await partXml(await target.save(), imported.partName)

		const sp = (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?ctrTitle[\s\S]*?<\/p:sp>/) ?? [''])[0]
		const runs = [...sp.matchAll(/<a:r>[\s\S]*?<\/a:r>/g)].map((m) => m[0])
		assert(runs.length > 0, 'ctrTitle has runs')
		for (const run of runs) {
			assert(/<a:rPr[^>]*\bsz="3200"/.test(run), 'each run carries the resolved titleStyle size (3200) explicitly')
		}
	})

	test('leaves a run that sets its own sz untouched', async () => {
		// slide1's first ctrTitle run is given an explicit sz="4444"; preserve must keep it
		// while its siblings still inherit the resolved master size (3200).
		const source = await Presentation.load(await deckMixedWithExplicitTitleSize())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		const sp = (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?ctrTitle[\s\S]*?<\/p:sp>/) ?? [''])[0]
		const runs = [...sp.matchAll(/<a:r>[\s\S]*?<\/a:r>/g)].map((m) => m[0])
		assert(/\bsz="4444"/.test(runs[0]), 'the run keeps its explicit size')
		assert(!/\bsz="3200"/.test(runs[0]), 'the inherited size did not overwrite it')
		assert(
			runs.slice(1).every((r) => /\bsz="3200"/.test(r)),
			'sibling runs still inherit the resolved master size'
		)
	})

	test('materializes a bgRef idx below 1000 through the regular fillStyleLst (not bgFillStyleLst)', async () => {
		const source = await Presentation.load(await deckEmptyBgRefFillStyleLst())
		const target = await openFixture('empty')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		const bg = (xml.match(/<p:bg>[\s\S]*?<\/p:bg>/) ?? [''])[0]
		assert(bg, 'the imported slide carries an explicit background')
		assert(!/bgRef|schemeClr|phClr/.test(bg), 'the background is fully materialized to a literal')
		assert(
			/<p:bgPr><a:solidFill><a:srgbClr val="[0-9A-Fa-f]{6}"\/><\/a:solidFill><\/p:bgPr>/.test(bg),
			'idx=1 resolves the plain fillStyleLst entry 1, not a bgFillStyleLst one'
		)
	})

	test('an unresolved bgRef (idx=0) falls back to an explicit a:noFill', async () => {
		const source = await Presentation.load(await deckEmptyBgRefIdxZero())
		const target = await openFixture('empty')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		const bg = (xml.match(/<p:bg>[\s\S]*?<\/p:bg>/) ?? [''])[0]
		assertEqual(bg, '<p:bg><p:bgPr><a:noFill/></p:bgPr></p:bg>', 'idx=0 materializes to an explicit transparent fill')
	})

	test("a slide's own background is left in place, not overwritten by the inherited one", async () => {
		const source = await Presentation.load(await deckMixedSlideOwnBackground())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'preserve' }) // slide1: now carries its own p:bg
		const xml = await partXml(await target.save(), imported.partName)

		const bg = (xml.match(/<p:bg>[\s\S]*?<\/p:bg>/) ?? [''])[0]
		assert(/<a:srgbClr val="123456"/.test(bg), 'the slide keeps its own literal background')
		assertEqual(
			(xml.match(/<p:bg>/g) ?? []).length,
			1,
			'no second background was inserted alongside the slide-owned one'
		)
	})

	test('a slide with no inherited background anywhere gains no synthetic p:bg', async () => {
		const source = await Presentation.load(await deckMixedNoBackgroundAnywhere())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'preserve' }) // slide1: no own bg
		const xml = await partXml(await target.save(), imported.partName)

		assert(!/<p:bg>/.test(xml), 'no background is baked on when the source chain defines none')
	})

	test('materializes a p:style effectRef into spPr (effectLst + scene3d + sp3d lifted from the fmtScheme)', async () => {
		const source = await Presentation.load(await deckMixedEffectRefMaterialized())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		// The mutated shape's effectRef neutralizes like the rest of the style matrix…
		assert(/<a:effectRef idx="0"\/>/.test(xml), 'effectRef neutralized to idx="0" with no colour')
		// …and the theme's effectStyleLst entry 3 (effectLst + scene3d + sp3d) was lifted into spPr.
		assert(/<p:spPr[^>]*>[\s\S]*?<a:effectLst>[\s\S]*?<a:outerShdw/.test(xml), 'effectLst lifted into spPr')
		assert(/<a:scene3d>[\s\S]*?<a:camera/.test(xml), 'scene3d lifted into spPr')
		assert(/<a:sp3d>[\s\S]*?<a:bevelT/.test(xml), 'sp3d lifted into spPr')
	})

	test('leaves a run that already carries its own colour untouched (not overwritten by the inherited one)', async () => {
		const target = await openFixture('empty')
		const source = await openFixture('multi-theme')
		const imported = target.importSlide(source, 1, { theme: 'preserve' }) // slide2: body run has its own FF00FF fill
		const xml = await partXml(await target.save(), imported.partName)

		const sp = (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?type="body"[\s\S]*?<\/p:sp>/) ?? [''])[0]
		assert(sp, 'the imported slide has its body placeholder')
		assert(/<a:solidFill><a:srgbClr val="FF00FF"\/><\/a:solidFill>/.test(sp), 'the run keeps its own explicit colour')
		assertEqual((sp.match(/<a:solidFill>/g) ?? []).length, 1, 'no duplicate solidFill was baked onto the run')
	})

	test('bakes each footer-trio placeholder its OWN-TYPE inherited geometry, not another member of the trio', async () => {
		// placeholder-footer-trio's slide dt/ftr/sldNum carry no own a:xfrm and their
		// layout defines none either, so each inherits from the SAME-TYPE master
		// placeholder — three deliberately distinct boxes. A resolver that matched the
		// trio by txStyles category (all `other`) would bake the date box onto all
		// three; preserve must bake each its own-type box (values read straight out of
		// the fixture's slideMaster1.xml).
		const target = await openFixture('placeholder-footer-trio')
		const source = await openFixture('placeholder-footer-trio')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		const spOf = (type) =>
			(xml.match(new RegExp(`<p:sp>(?:(?!</p:sp>)[\\s\\S])*?type="${type}"[\\s\\S]*?</p:sp>`)) ?? [''])[0]
		const expect = {
			dt: '<a:off x="508000" y="6095999"/><a:ext cx="2540000" cy="508000"/>',
			ftr: '<a:off x="3810000" y="6349999"/><a:ext cx="4572000" cy="381000"/>',
			sldNum: '<a:off x="9906000" y="5841999"/><a:ext cx="1778000" cy="635000"/>',
		}
		for (const [type, off] of Object.entries(expect)) {
			const sp = spOf(type)
			assert(sp, `the imported slide has its ${type} placeholder`)
			assert(sp.includes(off), `${type} bakes its own-type master box (${off})`)
			// None of the three may collapse onto the date box (the first `other` in doc order).
			if (type !== 'dt') assert(!/x="508000" y="6095999"/.test(sp), `${type} did not collapse onto the date box`)
		}
	})

	test('bakes distinct placeholder-inherited sizes per explicit paragraph level', async () => {
		// The body placeholder carries 5 paragraphs at lvl 0..4, each with a bare run; the
		// master bodyStyle defines a different sz per level (2800/2400/2000/1800/1800).
		const source = await Presentation.load(await deckMultiLevelBody())
		const target = await openFixture('layout-placeholder-bodypr')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		const sp = (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?idx="1"[\s\S]*?<\/p:sp>/) ?? [''])[0]
		assert(sp, 'the imported slide still has its body placeholder')
		const runs = [...sp.matchAll(/<a:r>[\s\S]*?<\/a:r>/g)].map((m) => m[0])
		assertEqual(runs.length, 5, 'all five level runs are present')
		const sizes = runs.map((r) => (r.match(/\bsz="(\d+)"/) ?? [])[1]).join(',')
		assertEqual(sizes, '2800,2400,2000,1800,1800', 'each run bakes its own level size from bodyStyle')
	})

	test('skips a placeholder that carries no text body, and still flattens the rest of the slide', async () => {
		// `p:txBody` is minOccurs="0" on p:CT_Shape, so a placeholder holding no text
		// (an empty picture/chart/table placeholder) is ordinary PowerPoint output. The
		// four text passes must skip it without aborting — the geometry pass, which does
		// not read the text body, still bakes its inherited box.
		const source = await Presentation.load(await deckMixedPlaceholderNoTxBody())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		const sp = (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?type="subTitle"[\s\S]*?<\/p:sp>/) ?? [''])[0]
		assert(sp, 'the text-less subTitle placeholder survived the import')
		assert(!/<p:txBody/.test(sp), 'it still carries no text body')
		assert(
			/<a:off x="1371600" y="3886200"\/><a:ext cx="6400800" cy="1752600"\/>/.test(sp),
			'the geometry pass still baked its inherited layout box'
		)
		// The sibling placeholder is untouched by the skip.
		const title = (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?ctrTitle[\s\S]*?<\/p:sp>/) ?? [''])[0]
		assert(/<a:srgbClr val="333399"\/>/.test(title), 'the ctrTitle run still got its inherited colour baked')
		assert(/\bsz="3200"/.test(title), 'the ctrTitle run still got its inherited size baked')
	})

	test('falls through a source layout placeholder that carries no text body', async () => {
		// The layout tier of the style chain reads its `a:lstStyle` out of the layout
		// placeholder's `p:txBody`; with no text body that tier contributes nothing and
		// resolution must continue to the master titleStyle rather than stop.
		const source = await Presentation.load(await deckMixedLayoutPlaceholderNoTxBody())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		const sp = (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?ctrTitle[\s\S]*?<\/p:sp>/) ?? [''])[0]
		assert(
			/<a:solidFill><a:srgbClr val="333399"\/><\/a:solidFill>/.test(sp),
			'the run colour still resolves from the master titleStyle (333399)'
		)
		assert(/\bsz="3200"/.test(sp), 'the run size still resolves from the master titleStyle (3200)')
	})

	test('resolves one inherited value per paragraph level, reusing it across paragraphs at that level', async () => {
		// Two paragraphs at the same (default) level in one placeholder: the second must
		// bake the same colour/size as the first — the per-level lookup is memoized, and
		// a memo that returned a different answer on the second hit would show up here.
		const source = await Presentation.load(await deckMixedTwoParagraphsSameLevel())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		const sp = (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?ctrTitle[\s\S]*?<\/p:sp>/) ?? [''])[0]
		const paras = [...sp.matchAll(/<a:p>[\s\S]*?<\/a:p>/g)].map((m) => m[0])
		assertEqual(paras.length, 2, 'both paragraphs survived')
		assert(/Second line/.test(paras[1]), 'the spliced paragraph is the second one')
		for (const [i, para] of paras.entries()) {
			const runs = [...para.matchAll(/<a:r>[\s\S]*?<\/a:r>/g)].map((m) => m[0])
			assert(runs.length > 0, `paragraph ${i} has runs`)
			for (const run of runs) {
				assert(
					/<a:solidFill><a:srgbClr val="333399"\/><\/a:solidFill>/.test(run),
					`paragraph ${i} bakes the level-0 colour`
				)
				assert(/\bsz="3200"/.test(run), `paragraph ${i} bakes the level-0 size`)
			}
		}
	})

	test('bakes nothing onto a placeholder whose source style chain defines nothing', async () => {
		// A picture placeholder whose type/idx match no layout or master placeholder, in a
		// deck whose master defines no `p:txStyles` at all: there is no tier to resolve a
		// box, a colour, or a size from. preserve must leave the shape exactly as authored
		// rather than invent a value — the run re-binds to the destination, which is the
		// only thing left that can style it.
		const source = await Presentation.load(await deckMixedInheritsNothing())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		const sp = (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?type="pic"[\s\S]*?<\/p:sp>/) ?? [''])[0]
		assert(sp, 'the orphan picture placeholder survived the import')
		assert(!/<a:xfrm/.test(sp), 'no geometry was baked — nothing in the source chain defines one')
		assert(!/<a:solidFill/.test(sp), 'no run colour was baked')
		assert(!/\bsz="/.test(sp), 'no run size was baked')

		// The same emptied chain applies to the slide's own ctrTitle.
		const title = (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?ctrTitle[\s\S]*?<\/p:sp>/) ?? [''])[0]
		assert(!/<a:solidFill/.test(title), 'the ctrTitle run resolves no colour either')
		assert(!/\bsz="/.test(title), 'nor a size')
	})

	test('leaves runs alone when the slide fixes their colour/size at paragraph or text-body level', async () => {
		// "Already fixed" is a three-tier ladder — the run's own `a:rPr`, then the
		// paragraph's `a:pPr/a:defRPr`, then the text body's `a:lstStyle` level. Only the
		// top tier is exercised by the authored fixtures. A value set at either lower tier
		// survives a rebind just as well, so preserve must not bake over it: paragraph 1
		// fixes colour and size via `a:pPr/a:defRPr`, and paragraph 2 (which sets no
		// `a:pPr` at all) has its size fixed by the text body's `a:lstStyle`.
		const source = await Presentation.load(await deckMixedSlideFixesRunProps())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		const sp = (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?ctrTitle[\s\S]*?<\/p:sp>/) ?? [''])[0]
		const paras = [...sp.matchAll(/<a:p>[\s\S]*?<\/a:p>/g)].map((m) => m[0])
		assertEqual(paras.length, 2, 'both paragraphs survived')

		const runsOf = (para) => [...para.matchAll(/<a:r>[\s\S]*?<\/a:r>/g)].map((m) => m[0])
		const first = runsOf(paras[0])
		assert(first.length > 0, 'paragraph 1 has runs')
		for (const run of first) {
			assert(!/<a:solidFill/.test(run), 'the paragraph defRPr fixes the colour, so none was baked onto the run')
			assert(!/\bsz="/.test(run), 'the paragraph defRPr fixes the size, so none was baked onto the run')
		}
		assert(
			/<a:defRPr sz="2222"><a:solidFill><a:srgbClr val="ABCDEF"\/><\/a:solidFill><\/a:defRPr>/.test(paras[0]),
			'the paragraph defRPr itself is carried through untouched'
		)

		const second = runsOf(paras[1])
		assertEqual(second.length, 1, 'paragraph 2 has its one run')
		assert(!/\bsz="/.test(second[0]), "the text body's lstStyle fixes the size, so none was baked")
		assert(
			/<a:solidFill><a:srgbClr val="333399"\/><\/a:solidFill>/.test(second[0]),
			'its colour is not fixed anywhere on the slide, so the inherited one is still baked'
		)
	})

	test('a bgRef at exactly 1000 selects no bgFillStyleLst entry and falls back to a:noFill', async () => {
		// The `bgFillStyleLst` range starts at 1000 and the list is 1-based, so idx=1000
		// addresses entry 0 — one below the first. It resolves to nothing, the same as
		// idx=0, rather than off-by-one'ing onto the first entry.
		const source = await Presentation.load(await deckEmptyBgRefIdx1000())
		const target = await openFixture('empty')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		const bg = (xml.match(/<p:bg>[\s\S]*?<\/p:bg>/) ?? [''])[0]
		assertEqual(bg, '<p:bg><p:bgPr><a:noFill/></p:bgPr></p:bg>', 'idx=1000 resolves to nothing, not to entry 1')
	})

	test('an unresolvable lnRef/effectRef materializes nothing and still neutralizes the ref', async () => {
		// Two ways a style-matrix ref resolves to nothing on input PowerPoint accepts:
		// `idx="0"` (constantly written — "no line from the matrix") and an `idx` past the
		// end of the theme's three-entry lists (`@idx` is an unbounded unsignedInt). Both
		// must leave `spPr` alone while the ref is still neutralized, so nothing re-resolves
		// against the destination theme.
		const baseline = await openFixture('mixed')
		baseline.importSlide(await openFixture('mixed'), THEMED_SLIDE_INDEX, { theme: 'preserve' })
		const baseXml = await partXml(await baseline.save(), baseline.slides[baseline.slides.length - 1].partName)
		const countLines = (xml) => (xml.match(/<a:ln[ >]/g) ?? []).length
		assert(countLines(baseXml) > 0, 'precondition: the unmutated import materializes lines from the style matrix')

		for (const [label, build] of Object.entries({
			'idx="0"': deckMixedLnRefIdxZero,
			'idx past the end of the list': deckMixedStyleRefOutOfRange,
		})) {
			const target = await openFixture('mixed')
			const imported = target.importSlide(await Presentation.load(await build()), THEMED_SLIDE_INDEX, {
				theme: 'preserve',
			})
			const xml = await partXml(await target.save(), imported.partName)
			assert(countLines(xml) < countLines(baseXml), `${label}: no line was materialized from the style matrix`)
			assert(/<a:lnRef idx="0"\/>/.test(xml), `${label}: the lnRef is still neutralized`)
			assert(!/<a:lnRef idx="[1-9]/.test(xml), `${label}: no lnRef survives with a resolvable idx`)
		}

		// The out-of-range effectRef lifts no effect children either.
		const target = await openFixture('mixed')
		const imported = target.importSlide(
			await Presentation.load(await deckMixedStyleRefOutOfRange()),
			THEMED_SLIDE_INDEX,
			{
				theme: 'preserve',
			}
		)
		const xml = await partXml(await target.save(), imported.partName)
		assert(!/<a:scene3d>/.test(xml) && !/<a:sp3d>/.test(xml), 'no effectStyle children were lifted into spPr')
		assert(/<a:effectRef idx="0"\/>/.test(xml), 'the effectRef is still neutralized')
	})

	test('leaves a stray phClr symbolic instead of resolving it to a colour', async () => {
		// `phClr` names "whatever colour this style-matrix entry is being filled with".
		// Away from a fmtScheme entry there is no such colour, so it must survive the
		// scheme-colour sweep untouched rather than resolve to some slot.
		const source = await Presentation.load(await deckMixedStraySchemePhClr())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		const sp = (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?ctrTitle[\s\S]*?<\/p:sp>/) ?? [''])[0]
		assert(/<a:solidFill><a:schemeClr val="phClr"\/><\/a:solidFill>/.test(sp), 'the stray phClr is left symbolic')
	})

	test('substitutes only phClr in a style-matrix entry, resolving its other scheme colours normally', async () => {
		// A fmtScheme fill entry may pin part of itself to a fixed scheme slot alongside
		// the `phClr` the ref supplies. The substitution must replace only the `phClr`;
		// the fixed slot then flattens through the ordinary scheme-colour sweep.
		const source = await Presentation.load(await deckMixedFillStyleFixedSchemeStop())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		assert(
			/<a:gs pos="0"><a:srgbClr val="00E4A8"\/><\/a:gs>/.test(xml),
			'the phClr stop took the fillRef colour (accent1 = 00E4A8)'
		)
		assert(
			/<a:gs pos="100000"><a:srgbClr val="000000"\/><\/a:gs>/.test(xml),
			'the fixed tx1 stop resolved through the clrMap to dk1 (000000), not to the ref colour'
		)
		assert(!/schemeClr/.test(xml), 'neither stop was left symbolic')
	})

	test('imports a slide that resolves to no layout or master at all, baking nothing', async () => {
		// A Slide part is *permitted* to relate to a slideLayout (ECMA-376 Part 1 §13.3.8),
		// not required to, and the master is only reachable through the layout — so a slide
		// missing that one relationship resolves to neither root. Every placeholder pass
		// then has no source chain to read, and must no-op rather than throw or invent a
		// value; the slide still imports, and still flattens what it owns outright.
		const source = await Presentation.load(await deckMixedSlideNoLayoutRel())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		const sp = (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?ctrTitle[\s\S]*?<\/p:sp>/) ?? [''])[0]
		assert(sp, 'the placeholder still imported')
		assert(!/<a:xfrm/.test(sp), 'no geometry was baked — there is no layout or master to read one from')
		assert(!/<a:solidFill/.test(sp), 'no run colour was baked')
		assert(!/\bsz="/.test(sp), 'no run size was baked')
		assert(/<p:ph type="ctrTitle"\/>/.test(sp), 'it is still a placeholder, now resolving against the destination')
	})

	test('reports no inherited frame for a placeholder whose slide resolves to no layout or master', async () => {
		// The read-model sibling of the bake above: `resolvedFrame` walks the same
		// layout-then-master tiers and must skip both when the slide has neither.
		const source = await Presentation.load(await deckMixedSlideNoLayoutRel())
		const shape = source.slides[0].shapes[0]
		assert(shape.placeholder, 'the shape under test is a placeholder')
		assertEqual(shape.left, null, 'precondition: it carries no own geometry')
		assertEqual(shape.resolvedFrame, null, 'with no tier to inherit from, there is no effective frame')
	})

	test('leaves a scheme colour symbolic when the source theme slot holds no literal RGB', async () => {
		// `a:scrgbClr` is a legal `a:CT_Color` child that this reader resolves to no 6-hex
		// RGB. The clrMap still routes the token to a slot — the slot just has nothing
		// literal behind it. Flattening must degrade to leaving the token symbolic, so the
		// colour re-binds to the destination, rather than emitting a bogus `a:srgbClr`.
		// (`a:prstClr` and `a:hslClr` used to reach this branch too; both resolve now, and
		// a slot holding one is flattened to the colour it names, which is correct.)
		const source = await Presentation.load(await deckMixedUnreadableColorScheme())
		const target = await openFixture('mixed')
		const imported = target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'preserve' })
		const xml = await partXml(await target.save(), imported.partName)

		assert(/<a:schemeClr val="tx2"/.test(xml), 'the unresolvable tx2 token survives as a scheme colour')
		assert(!/<a:scrgbClr/.test(xml), 'the unresolved colour element was not copied onto the slide either')
		assert(!/<a:srgbClr val="333399"/.test(xml), 'nor was tx2 flattened to the RGB the replaced Fusion scheme held')
	})

	// Across two *different* decks, so the copy has something to copy: importing from a
	// deck the destination already holds byte-for-byte binds to what is there instead
	// (`ops/part-reuse.ts`), which would make a same-file import prove nothing here.
	test('the default (no option) still copies the source theme subgraph', async () => {
		const target = await openFixture('empty')
		const source = await openFixture('theme-colors')
		const themesBefore = countParts(target.opc, /\/theme\/theme\d+\.xml$/)
		target.importSlide(source, 0) // default: copy
		const after = countParts(target.opc, /\/theme\/theme\d+\.xml$/)
		assert(after > themesBefore, 'the default copy mode brings a source theme across')
	})

	test.skipIf(!validatorInstalled)('a preserve-imported deck stays schema-valid', async () => {
		const target = await openFixture('mixed')
		const source = await openFixture('mixed')
		target.importSlide(source, THEMED_SLIDE_INDEX, { theme: 'preserve' })
		target.importSlide(source, 5, { theme: 'preserve' }) // slide6: also themed (schemeClr + p:style)
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})

	test.skipIf(!validatorInstalled)('a carryMasterGraphics-imported deck stays schema-valid', async () => {
		const target = await openFixture('mixed')
		target.importSlide(await openFixture('mixed'), 0, { theme: 'preserve', carryMasterGraphics: true })
		target.importSlide(await Presentation.load(await deckWithMasterPicture()), 0, {
			theme: 'preserve',
			carryMasterGraphics: true,
		})
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})

	test.skipIf(!validatorInstalled)('a deck with baked placeholder geometry/size stays schema-valid', async () => {
		const target = await openFixture('mixed')
		target.importSlide(await openFixture('mixed'), 0, { theme: 'preserve' }) // slide1: geometry + run size baked
		target.importSlide(await Presentation.load(await deckMixedNoLayoutCtrTitleXfrm()), 0, { theme: 'preserve' })
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})

	// Every spliced source above claims "PowerPoint could have written this". The
	// validator is what turns that claim into a check: a variant it rejects belongs
	// in the header note as impossible input, not in a test.
	test.skipIf(!validatorInstalled)('the spliced placeholder-chain sources are themselves schema-valid', async () => {
		const cases = Object.entries({
			'no placeholder text body': deckMixedPlaceholderNoTxBody,
			'no layout placeholder text body': deckMixedLayoutPlaceholderNoTxBody,
			'two paragraphs at one level': deckMixedTwoParagraphsSameLevel,
			'empty inheritance chain': deckMixedInheritsNothing,
			'slide fixes its own run props': deckMixedSlideFixesRunProps,
			'bgRef idx 1000': deckEmptyBgRefIdx1000,
			'lnRef idx 0': deckMixedLnRefIdxZero,
			'style refs past the end of the list': deckMixedStyleRefOutOfRange,
			'stray phClr on a slide shape': deckMixedStraySchemePhClr,
			'fmtScheme entry with a fixed scheme stop': deckMixedFillStyleFixedSchemeStop,
			'slide with no slideLayout relationship': deckMixedSlideNoLayoutRel,
			'clrScheme stated in preset colours': deckMixedUnreadableColorScheme,
		})
		// Validation is batched by the package, but each case still builds and
		// re-serializes a whole deck, so a dozen at once is a memory spike inside a single
		// test. Four at a time is the same ceiling vitest.config.ts sets for the schema
		// suite.
		for (let i = 0; i < cases.length; i += 4) {
			const done = await Promise.all(
				cases.slice(i, i + 4).map(async ([name, build]) => ({
					name,
					errors: await validateBuf(Buffer.from(await build())),
				}))
			)
			for (const { name, errors } of done) {
				assertEqual(errors.length, 0, `${name}: ${JSON.stringify(errors).slice(0, 2000)}`)
			}
		}
	})

	test.skipIf(!validatorInstalled)(
		'a deck imported from a degenerate placeholder chain stays schema-valid',
		async () => {
			const target = await openFixture('mixed')
			target.importSlide(await Presentation.load(await deckMixedPlaceholderNoTxBody()), 0, { theme: 'preserve' })
			target.importSlide(await Presentation.load(await deckMixedLayoutPlaceholderNoTxBody()), 0, { theme: 'preserve' })
			target.importSlide(await Presentation.load(await deckMixedTwoParagraphsSameLevel()), 0, { theme: 'preserve' })
			target.importSlide(await Presentation.load(await deckMixedInheritsNothing()), 0, { theme: 'preserve' })
			target.importSlide(await Presentation.load(await deckMixedSlideFixesRunProps()), 0, { theme: 'preserve' })
			target.importSlide(await Presentation.load(await deckMixedStraySchemePhClr()), 0, { theme: 'preserve' })
			target.importSlide(await Presentation.load(await deckMixedSlideNoLayoutRel()), 0, { theme: 'preserve' })
			for (const build of [
				deckMixedLnRefIdxZero,
				deckMixedStyleRefOutOfRange,
				deckMixedFillStyleFixedSchemeStop,
				deckMixedUnreadableColorScheme,
			]) {
				target.importSlide(await Presentation.load(await build()), THEMED_SLIDE_INDEX, { theme: 'preserve' })
			}
			const errors = await validateBuf(Buffer.from(await target.save()))
			assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
		}
	)
})
