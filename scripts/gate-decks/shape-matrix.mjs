/**
 * Gate deck: the slide-object constructs the showcase corpus never reaches.
 *
 * Corpus for `scripts/byte-identity.mjs`. The showcase decks are presentation decks, and a
 * plausible business deck has no reason to carry a native equation, a shape-bound connector,
 * an SVG picture, a Zoom tile or a chartEx chart — so every emitter behind those was
 * ungated, and a PASS said nothing about them. See `./README.md` for why a gate deck is a
 * separate thing from a showcase, and why one may only ever grow.
 *
 * One construct family per slide. Every case is reached on every run; nothing is sampled.
 *
 * What each slide is here to reach:
 *
 *   connectors    `<a:stCxn>` / `<a:endCxn>` and BOTH arities of `<p:cNvCxnSpPr>` — a
 *                 connector with no resolved binding emits it self-closing, one with a
 *                 binding emits it paired, and only a deck carrying both proves either.
 *                 All three routing presets, with and without `adj` guides.
 *   equations     the `a14` markup-compatibility envelope, which wraps the WHOLE `<p:sp>`:
 *                 a display equation and an inline one, plus a plain shape on the same
 *                 slide so the un-enveloped arm is diffed beside it.
 *   transitions   `p14:dur` and its `mc:AlternateContent` / `mc:Fallback` pair, against a
 *                 bare `<p:transition>` on the next slide. The two are different emitters.
 *   svg picture   `<asvg:svgBlip>`, the dual-rel form where a raster blip and a vector one
 *                 name the same picture.
 *   zoom          Slide, Section and Summary Zoom — three graphicData URIs through one
 *                 `mc:AlternateContent` envelope, plus the sections they navigate to.
 *   chartex       a chartEx chart, whose two mandatory style sidecars are reached through
 *                 Microsoft rel types no classic chart uses.
 *   online video  the external `video` rel paired with the MS `media` rel on one target.
 *
 * NOT here: OLE objects. An OLE payload is a real binary (a compound file or an embedded
 * package), and a gate deck that reads an asset off disk trades one blind spot for a second
 * kind of flap. `gen/slide/objects/ole.ts` therefore stays ungated — the remaining gap, and
 * the thing to fix before anyone refactors it. 3-D models need no case: `field-notes`
 * already emits one, so `model3d.ts` is covered by the showcase corpus.
 */
import TsPptx from '../../dist/node.js'

/** A 1x1 transparent PNG, inline so this deck reads no asset off disk. */
const PNG_1PX =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

/** A vector square, small enough to read inline. Emitted as the `asvg:svgBlip` half of a picture. */
const SVG_MARK = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 8 8"><rect width="8" height="8"/></svg>'

/** Raw OMML for a display equation. Hand-written: `latexToOmml` is a converter, not the thing under test. */
const OMML_DISPLAY =
	'<m:oMath><m:sSup><m:e><m:r><m:t>e</m:t></m:r></m:e><m:sup><m:r><m:t>x</m:t></m:r></m:sup></m:sSup></m:oMath>'
/** Raw OMML for an inline run, which takes the same envelope in a different position. */
const OMML_INLINE = '<m:oMath><m:r><m:t>n-1</m:t></m:r></m:oMath>'

const TITLE = { x: 0.3, y: 0.15, w: 12.7, h: 0.4, fontSize: 14, bold: true }

/** @param {import('../../dist/node.js').default} pptx */
function connectors(pptx) {
	const slide = pptx.addSlide({ sectionTitle: 'Constructs' })
	slide.addText('connectors', TITLE)
	slide.addText('A', { objectName: 'BoxA', x: 0.5, y: 1.2, w: 1.4, h: 0.8, fill: { color: 'DDEEFF' } })
	slide.addText('B', { objectName: 'BoxB', x: 5.5, y: 1.2, w: 1.4, h: 0.8, fill: { color: 'FFEEDD' } })

	// Both endpoints bound: the paired <p:cNvCxnSpPr> carrying two children.
	slide.addConnector({
		type: 'straight',
		x1: 1.9,
		y1: 1.6,
		x2: 5.5,
		y2: 1.6,
		startShape: 'BoxA',
		startShapeIdx: 3,
		endShape: 'BoxB',
		endShapeIdx: 1,
		color: '204060',
		width: 2,
	})
	// One endpoint bound: the same element with one child.
	slide.addConnector({ type: 'elbow', x1: 1.2, y1: 2.0, x2: 6.2, y2: 3.2, startShape: 'BoxA', adj: 35 })
	// Neither bound: the self-closing arm. A different code path, not a different value.
	slide.addConnector({ type: 'curved', x1: 0.5, y1: 4.0, x2: 6.5, y2: 5.0, color: 'AA3333', width: 3 })
	// Multi-bend, so the `adj` guide list is longer than one.
	slide.addConnector({ type: 'elbow', bends: 3, adj: [25, 50, 75], x1: 7.5, y1: 1.2, x2: 12.5, y2: 5.0 })
	// Reversed endpoints, which is what puts flipH/flipV on the transform.
	slide.addConnector({ type: 'straight', x1: 12.0, y1: 5.5, x2: 8.0, y2: 6.5, endArrowType: 'triangle' })
}

/** @param {import('../../dist/node.js').default} pptx */
function equations(pptx) {
	const slide = pptx.addSlide({ sectionTitle: 'Constructs' })
	slide.addText('equations', TITLE)
	// The un-enveloped arm, on the same slide so a diff shows both beside each other.
	slide.addText('plain shape', { x: 0.3, y: 0.8, w: 5, h: 0.6 })
	// Display: the whole <p:sp> goes inside <mc:AlternateContent><mc:Choice Requires="a14">.
	slide.addText([{ math: OMML_DISPLAY }], { x: 0.3, y: 1.6, w: 6, h: 1.2, fontSize: 24 })
	// Inline: the same envelope around a shape whose <a14:m> flows between ordinary runs.
	slide.addText([{ text: 'for all ' }, { math: OMML_INLINE, inline: true }, { text: ' terms' }], {
		x: 0.3,
		y: 3.0,
		w: 8,
		h: 0.8,
		fontSize: 18,
	})
}

/** @param {import('../../dist/node.js').default} pptx */
function transitions(pptx) {
	// Exact duration: the mc:AlternateContent form, a p14 Choice plus a base Fallback.
	const timed = pptx.addSlide({ sectionTitle: 'Constructs' })
	timed.addText('transition, exact duration', TITLE)
	timed.transition = { type: 'push', durationMs: 1250, variant: { dir: 'd' } }

	// No duration: the bare <p:transition>, which is the other branch of the same function.
	const coarse = pptx.addSlide({ sectionTitle: 'Constructs' })
	coarse.addText('transition, coarse speed', TITLE)
	coarse.transition = { type: 'fade', speed: 'slow', advanceOnClick: false, advanceAfterMs: 4000 }
}

/** @param {import('../../dist/node.js').default} pptx */
function pictures(pptx) {
	const slide = pptx.addSlide({ sectionTitle: 'References' })
	slide.addText('svg picture', TITLE)
	// A vector picture is a dual-rel blip: the raster fallback and the asvg:svgBlip extension.
	// `svg` is the whole source — passing `data` as well would make this an ordinary raster.
	slide.addImage({ svg: SVG_MARK, x: 0.3, y: 0.8, w: 2, h: 2 })
	// The raster-only form, so the extLst-bearing and extLst-less arms are both diffed.
	slide.addImage({ data: PNG_1PX, x: 3.0, y: 0.8, w: 2, h: 2 })
	// An online video: the external `video` rel paired with the MS `media` rel on one target.
	slide.addMedia({ type: 'online', link: 'https://www.youtube.com/embed/Dph6ynRVyUc', x: 6, y: 0.8, w: 4, h: 2.3 })
}

/** @param {import('../../dist/node.js').default} pptx */
function zooms(pptx) {
	// Cast because the published `Slide` type does not declare the three zoom methods, though
	// `SlideBuilder` implements them and the docs describe them. That is a public-surface gap,
	// not something this deck should paper over quietly — the deck still reaches the emitter.
	const slide = /** @type {any} */ (pptx.addSlide({ sectionTitle: 'References' }))
	slide.addText('zoom tiles', TITLE)
	slide.addSlideZoom({ target: 1, x: 0.3, y: 0.8, w: 3, h: 1.7 })
	slide.addSectionZoom({ sectionTitle: 'Constructs', x: 3.8, y: 0.8, w: 3, h: 1.7 })
	slide.addSummaryZoom({ x: 0.3, y: 2.9, w: 12.7, h: 3.6 })
}

/** @param {import('../../dist/node.js').default} pptx */
function chartex(pptx) {
	const slide = pptx.addSlide({ sectionTitle: 'References' })
	slide.addText('chartEx', TITLE)
	// A chartEx part reaches its two mandatory sidecars through Microsoft rel types that
	// no classic chart uses, and declares Microsoft content types for all three.
	slide.addChart([{ name: 'Stage', labels: ['Lead', 'Trial', 'Won'], values: [120, 48, 17] }], {
		type: 'funnel',
		x: 0.3,
		y: 0.8,
		w: 6,
		h: 5,
	})
}

async function compose() {
	const pptx = new TsPptx()
	pptx.layout = 'LAYOUT_WIDE'
	pptx.author = 'ts-pptx byte-identity gate'
	pptx.title = 'Slide-object construct matrix'

	// Both sections up front: `addSlide({ sectionTitle })` only files a slide into a section
	// that already exists, and Section/Summary Zoom have nothing to navigate to without one.
	pptx.addSection({ title: 'Constructs' })
	pptx.addSection({ title: 'References' })
	connectors(pptx)
	equations(pptx)
	transitions(pptx)
	pictures(pptx)
	chartex(pptx)
	zooms(pptx)
	return pptx
}

/** @param {string} outFile @returns {Promise<string>} */
export async function build(outFile) {
	const pptx = await compose()
	return await pptx.writeFile({ fileName: outFile })
}

export const gateDeck = {
	slug: 'shape-matrix',
	title: 'Slide-object construct matrix',
	description:
		'Byte-identity corpus for the slide-object emitters no showcase reaches — connectors, equations, timed transitions, SVG pictures, Zoom tiles, chartEx.',
	fileName: 'gate_shape_matrix.pptx',
	build,
}
