// Unit tests for the placeholder text-property INHERITANCE TIERS in
// src/read/api/theme-context.ts that a single real deck rarely holds all at once.
//
// The fixture suites (placeholder-inherit.pptx, multi-theme.pptx) exercise the
// bottom tier of the chain — the value a placeholder run inherits from the
// LAYOUT / MASTER placeholder style. They do not exercise the two upper tiers,
// because a PowerPoint-authored placeholder run seldom carries its own
// paragraph `a:defRPr` or a per-frame `a:lstStyle`:
//
//   1. the paragraph's own `a:pPr/a:defRPr`            (highest priority)
//   2. the text body's `a:lstStyle` level `a:defRPr`
//   3. the layout → master placeholder chain            (fixture-covered)
//
// `TextFrame` is exported, and its constructor is
// `(txBody, part, themeContext?, placeholder?)`, so a hand-authored `p:txBody`
// wrapped in a placeholder context drives tiers 1 and 2 straight through the
// read-model `Run.resolved*` getters — no fixture .pptx required. This mirrors
// the off-fixture pattern the style-accessor suite already uses. `resolveColorElement`
// is likewise exported and gets its own direct edge tests (alpha, unresolvable).

import { DOMParser } from '@xmldom/xmldom'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import PptxGenJS from '../../dist/node.js'
import { Presentation, TextFrame, AutoShape, resolveColorElement } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

/** A minimal FlattenContext: empty colour maps resolve `a:srgbClr` literally, and
 *  no layout/master roots means the bottom (placeholder-chain) tier finds nothing. */
function ctx(overrides = {}) {
	return {
		clrMap: new Map(),
		clrScheme: new Map(),
		fmtScheme: null,
		fontScheme: null,
		layoutRoot: null,
		masterRoot: null,
		...overrides,
	}
}

/** Parse hand-authored `p:txBody` inner XML into a `p:txBody` element. */
function txBodyEl(inner) {
	const xml = `<p:txBody xmlns:p="${P_NS}" xmlns:a="${A_NS}"><a:bodyPr/>${inner}</p:txBody>`
	return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

/** A placeholder `TextFrame` over `inner`, resolving against `flatten`. */
function placeholderFrame(inner, flatten = ctx()) {
	const placeholder = { ph: { type: 'body', idx: '0' }, flatten }
	// The read-side `resolved*` getters never touch `part`; a stand-in is enough.
	return new TextFrame(txBodyEl(inner), /** @type {any} */ ({}), flatten, placeholder)
}

/** First run of the first paragraph. */
function firstRun(frame) {
	const run = frame.paragraphs[0]?.runs[0]
	assert(run, 'expected a run')
	return run
}

/** Parse a single DrawingML element (`<a:srgbClr .../>`, `<a:fmtScheme>…`, etc.). */
function drawingEl(xml) {
	return new DOMParser().parseFromString(`<a:w xmlns:a="${A_NS}">${xml}</a:w>`, 'text/xml').documentElement.firstChild
}

/** An `AutoShape` over hand-authored `p:sp` XML, resolving against `flatten`. */
function autoShape(spXml, flatten = ctx()) {
	const xml = `<p:spTree xmlns:p="${P_NS}" xmlns:a="${A_NS}">${spXml}</p:spTree>`
	const spTree = new DOMParser().parseFromString(xml, 'text/xml').documentElement
	const el = spTree.getElementsByTagNameNS(P_NS, 'sp')[0]
	// Only `themeContext()` is exercised by the resolved-fill/line reads.
	return new AutoShape(el, /** @type {any} */ ({ themeContext: () => flatten }))
}

describe('resolveInheritedRunColor — the two upper tiers', () => {
	test('tier 1: a colourless run inherits its paragraph a:pPr/a:defRPr solidFill', () => {
		// The whitespace between <a:solidFill> and <a:srgbClr> is a text node the
		// colour lookup (firstChildElement) must skip.
		const run = firstRun(
			placeholderFrame(
				`<a:p><a:pPr><a:defRPr><a:solidFill>
					<a:srgbClr val="12AB34"/></a:solidFill></a:defRPr></a:pPr><a:r><a:t>x</a:t></a:r></a:p>`
			)
		)
		assertEqual(run.color, null, 'the run itself sets no colour')
		assertEqual(run.resolvedColor.hex, '12AB34', 'inherits the paragraph defRPr solidFill (tier 1)')
	})

	test('tier 2: with no paragraph defRPr, a run inherits the text body a:lstStyle level fill', () => {
		const run = firstRun(
			placeholderFrame(
				`<a:lstStyle><a:lvl1pPr><a:defRPr><a:solidFill><a:srgbClr val="0055AA"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>` +
					`<a:p><a:r><a:t>x</a:t></a:r></a:p>`
			)
		)
		assertEqual(run.resolvedColor.hex, '0055AA', 'inherits the slide lstStyle lvl1 fill (tier 2)')
	})

	test('tier 1 wins over tier 2 when both define a colour', () => {
		const run = firstRun(
			placeholderFrame(
				`<a:lstStyle><a:lvl1pPr><a:defRPr><a:solidFill><a:srgbClr val="0055AA"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>` +
					`<a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val="12AB34"/></a:solidFill></a:defRPr></a:pPr>` +
					`<a:r><a:t>x</a:t></a:r></a:p>`
			)
		)
		assertEqual(run.resolvedColor.hex, '12AB34', 'the paragraph defRPr (tier 1) governs over the lstStyle')
	})

	test('nothing in the chain defines a colour → resolvedColor is null', () => {
		// A run with no rPr, no paragraph/lstStyle fill, and a ctx with no layout/master.
		const run = firstRun(placeholderFrame(`<a:p><a:r><a:t>x</a:t></a:r></a:p>`))
		assertEqual(run.resolvedColor, null, 'no fill in any tier degrades to null, not a throw')
	})

	test('an explicit run solidFill with no colour child resolves to null', () => {
		// resolveSolidFillColor finds the a:solidFill but firstChildElement finds no
		// colour element inside it → the colour is not made literal, so null.
		const run = firstRun(placeholderFrame(`<a:p><a:r><a:rPr><a:solidFill/></a:rPr><a:t>x</a:t></a:r></a:p>`))
		assertEqual(run.resolvedColor, null, 'a childless solidFill yields no resolvable colour')
	})
})

describe('inherited run size / face / bold — the two upper tiers', () => {
	test('size: paragraph defRPr @sz, then lstStyle level @sz', () => {
		assertEqual(
			firstRun(placeholderFrame(`<a:p><a:pPr><a:defRPr sz="3600"/></a:pPr><a:r><a:t>x</a:t></a:r></a:p>`))
				.resolvedSizePt,
			36,
			'inherits the paragraph defRPr @sz (tier 1)'
		)
		assertEqual(
			firstRun(
				placeholderFrame(
					`<a:lstStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle><a:p><a:r><a:t>x</a:t></a:r></a:p>`
				)
			).resolvedSizePt,
			24,
			'inherits the lstStyle lvl1 @sz (tier 2)'
		)
	})

	test('face: a literal a:latin/@typeface in the paragraph defRPr is returned verbatim', () => {
		assertEqual(
			firstRun(
				placeholderFrame(
					`<a:p><a:pPr><a:defRPr><a:latin typeface="Rockwell"/></a:defRPr></a:pPr><a:r><a:t>x</a:t></a:r></a:p>`
				)
			).resolvedFontFace,
			'Rockwell',
			'a non-token face resolves to itself (tier 1)'
		)
	})

	test('bold: a paragraph defRPr @b="1" resolves to an inherited true', () => {
		assertEqual(
			firstRun(placeholderFrame(`<a:p><a:pPr><a:defRPr b="1"/></a:pPr><a:r><a:t>x</a:t></a:r></a:p>`)).resolvedBold,
			true,
			'inherits an explicit bold from the paragraph defRPr'
		)
	})

	test('a present-but-silent upper tier falls through to the next tier', () => {
		// The paragraph defRPr sets only @b — no @sz, no a:latin — so size and face
		// resolution must skip it and pick up the lstStyle level defRPr below.
		const run = firstRun(
			placeholderFrame(
				`<a:lstStyle><a:lvl1pPr><a:defRPr sz="1800"><a:latin typeface="Cambria"/></a:defRPr></a:lvl1pPr></a:lstStyle>` +
					`<a:p><a:pPr><a:defRPr b="1"/></a:pPr><a:r><a:t>x</a:t></a:r></a:p>`
			)
		)
		assertEqual(run.resolvedSizePt, 18, 'size skips the sz-less paragraph tier for the lstStyle tier')
		assertEqual(run.resolvedFontFace, 'Cambria', 'face skips the latin-less paragraph tier for the lstStyle tier')
		assertEqual(run.resolvedBold, true, 'bold still comes from the paragraph tier that defines it')
	})

	test('nothing in the chain defines size/face/bold → each resolves null', () => {
		const run = firstRun(placeholderFrame(`<a:p><a:r><a:t>x</a:t></a:r></a:p>`))
		assertEqual(run.resolvedSizePt, null, 'no @sz anywhere → null')
		assertEqual(run.resolvedFontFace, null, 'no a:latin anywhere → null')
		assertEqual(run.resolvedBold, null, 'no @b anywhere → null')
	})
})

describe('resolveColorElement — alpha and unresolvable edges', () => {
	test('an a:alpha transform surfaces a 0–1 alpha alongside the effective hex', () => {
		const resolved = resolveColorElement(drawingEl(`<a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr>`), ctx())
		assertEqual(resolved.hex, 'FF0000', 'base hex is the srgb value')
		assert(Math.abs(resolved.alpha - 0.5) < 1e-9, `alpha 50000 (thousandths of a %) → 0.5, got ${resolved.alpha}`)
	})

	test('an unmapped schemeClr (empty colour maps) resolves to null', () => {
		assertEqual(
			resolveColorElement(drawingEl(`<a:schemeClr val="accent1"/>`), ctx()),
			null,
			'no clrScheme entry → null'
		)
	})
})

describe('style-matrix fill/line fallback edges (resolveStyleFillColor / resolveStyleLineColor)', () => {
	test('a fillRef pointing at a gradient style entry has no single fill colour → null', () => {
		// styleRefFill returns the idx-1 fillStyleLst entry (a gradFill here), which is
		// not a solidFill, so resolveStyleFillColor has no single colour to report.
		const fmtScheme = drawingEl(
			`<a:fmtScheme><a:fillStyleLst>` +
				`<a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs>` +
				`<a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst></a:gradFill>` +
				`</a:fillStyleLst><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>`
		)
		const shape = autoShape(
			`<p:sp><p:style><a:fillRef idx="1"><a:srgbClr val="FF0000"/></a:fillRef></p:style><p:spPr/></p:sp>`,
			ctx({ fmtScheme })
		)
		assertEqual(shape.resolvedFill, null, 'a gradient style-matrix fill exposes no single resolved colour')
	})

	test('a shape with no explicit line and no p:style lnRef resolves no line colour → null', () => {
		const shape = autoShape(`<p:sp><p:spPr/></p:sp>`)
		assertEqual(shape.resolvedLine, null, 'no a:ln and no lnRef in the style matrix → null')
	})
})

describe('resolveSlideThemeParts — a broken theme chain degrades, not throws', () => {
	// Every real deck carries an intact slide → layout → master → theme chain, so
	// the "missing link" guards never fire on fixtures. Build a valid deck, then
	// delete the slide's slideLayout relationship in-memory: with no layout, the
	// whole chain collapses to null roots and empty colour maps. We assert the
	// degraded shape, not any writer colour, so this is not a write→read round-trip.
	async function slideWithNoLayoutRel() {
		const pptx = new PptxGenJS()
		pptx.addSlide().addText('hi', { x: 1, y: 1, w: 4, h: 1, color: '0000FF' })
		const zip = await JSZip.loadAsync(await pptx.stream())
		const relsName = 'ppt/slides/_rels/slide1.xml.rels'
		const rels = await zip.file(relsName).async('string')
		// Drop the single <Relationship … slideLayout … /> element.
		zip.file(relsName, rels.replace(/<Relationship\b[^>]*slideLayout[^>]*\/>/, ''))
		const broken = await zip.generateAsync({ type: 'uint8array' })
		return (await Presentation.load(broken)).slides[0]
	}

	test('themeContext resolves to empty maps and null roots when the layout link is gone', async () => {
		const slide = await slideWithNoLayoutRel()
		const ctx = slide.themeContext()
		assertEqual(ctx.clrMap.size, 0, 'no master → empty clrMap')
		assertEqual(ctx.clrScheme.size, 0, 'no theme → empty clrScheme')
		assertEqual(ctx.fmtScheme, null, 'no theme → null fmtScheme')
		assertEqual(ctx.fontScheme, null, 'no theme → null fontScheme')
		assertEqual(ctx.layoutRoot ?? null, null, 'the missing layout root is null')
		assertEqual(ctx.masterRoot ?? null, null, 'no layout → no master root')
	})

	test('a run on a chain-less slide still reads (resolution degrades to a raw value)', async () => {
		const slide = await slideWithNoLayoutRel()
		const shape = slide.shapes.find((s) => s.hasTextFrame)
		assert(shape, 'expected the text shape')
		const run = shape.textFrame.paragraphs[0].runs[0]
		// The explicit run colour still resolves; the point is nothing throws.
		assertEqual(run.resolvedColor.hex, '0000FF', 'an explicit run colour survives a broken theme chain')
	})
})
