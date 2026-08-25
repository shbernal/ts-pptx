/**
 * Authoring speaker notes onto a slide of a *loaded* deck ({@link Slide.addNotes}).
 *
 * The third way a notes slide comes into being, after the two in `notes-master.ts`
 * (carried from another `Presentation` by `importSlide({ importNotes: true })`, or
 * emitted by the generator and spliced on by `appendSlides`). Those two both *move*
 * an existing part; this one has no source part to move, so it builds one.
 *
 * The part it builds is the generator's, not a second design: the fixed three-
 * placeholder frame comes from {@link makeXmlNotesSlideSkeleton}, and the body
 * paragraphs carry the same elements and attributes the write path emits for a plain
 * `pptx.addSlide().addNotes(text)`. `test/read/add-notes.test.js` asserts that
 * equivalence, so the two cannot drift apart unnoticed — the same guarantee
 * `src/ooxml/` gives the constants both halves share.
 *
 * The one difference is empty-element *spelling*: the write path emits `a:rPr` as an
 * open/close pair and this emits it self-closed, because a part the read model
 * authors is reserialized through `read/oxml/dom.ts`, which self-closes every empty
 * element. Writing the open/close form here would not survive the first save, so the
 * form that does is written in the first place and the test normalizes the spelling
 * rather than pretending the bytes match.
 *
 * Notes text is split on `\n` into paragraphs, matching the write-side `addNotes`;
 * runs carry no formatting of their own. Per-run formatting, hyperlinks, and
 * multi-run paragraphs are reached afterwards through `slide.notesTextFrame`, which
 * is the read model's normal text-editing surface.
 */

import { el, raw, voidEl } from '../../../gen/oxml/el.js'
import { makeXmlNotesSlideSkeleton } from '../../../gen/slide/notes.js'
import { attr, firstChild, getElements, parseXml, removeChildrenByQName, type Element } from '../../oxml/dom.js'
import { relativePartName } from '../../opc/partnames.js'
import { NOTES_MASTER_REL, NOTES_SLIDE_REL, SLIDE_REL } from '../../../ooxml/rel-types.js'
import { InternalError, PackageReadError } from '../../../errors.js'
import { ensureNotesMasterForAuthoring } from './notes-master.js'
import type { Slide } from '../slide.js'
import { OOXML_NS } from '../../../ooxml/namespaces.js'

const NOTES_SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'
const textEncoder = new TextEncoder()

/**
 * The `a:p` children of a notes body, one per `\n`-separated line: the write path's
 * output for an unformatted note, modulo the empty-element spelling described above.
 * An empty line yields a paragraph with no run, exactly as `genXmlNotesParagraphs`
 * does.
 * @param {string} text - the note, `\n` separating paragraphs
 * @return {string} concatenated `a:p` XML
 */
function notesParagraphsXml(text: string): string {
	return text
		.split('\n')
		.map((line) => line.replace(/\r/g, ''))
		.map((line) =>
			el('a:p', null, [
				raw(
					line === ''
						? ''
						: el('a:r', null, [raw(voidEl('a:rPr', { lang: 'en-US', dirty: 0 })), raw(el('a:t', null, line))])
				),
				raw(voidEl('a:endParaRPr', { lang: 'en-US', dirty: 0 })),
			])
		)
		.join('')
}

/**
 * Replace a notes body's paragraphs with `text`, in the body's own document.
 *
 * The path taken when the slide *already* has a notes part — an imported one, or one
 * an earlier `addNotes` wrote. Only the `a:p` children are replaced, so the part's
 * geometry, its `a:bodyPr`/`a:lstStyle`, and its other two placeholders survive. The
 * body's hyperlink relationships are deliberately *not* pruned: a caller replacing
 * text has not asked for its rels to be rewritten, and a now-unreferenced notes
 * hyperlink rel is inert (nothing resolves it on save).
 * @param {Element} txBody - the body placeholder's `p:txBody`
 * @param {string} text - the note, `\n` separating paragraphs
 */
function replaceNotesParagraphs(txBody: Element, text: string): void {
	const doc = txBody.ownerDocument
	if (!doc) throw new InternalError('oxml/node-has-no-document', 'Notes body has no owner document')
	removeChildrenByQName(txBody, ['a:p'])
	// Parsed inside a namespace-declaring wrapper, then imported: the same move
	// `animation.ts` makes to bring authored XML into a loaded part's document.
	const wrapper = parseXml(`<w xmlns:a="${OOXML_NS.a}">${notesParagraphsXml(text)}</w>`).documentElement
	if (!wrapper)
		throw new InternalError('oxml/node-has-no-document', 'Authored notes paragraphs did not parse to a root')
	for (const paragraph of Array.from(wrapper.childNodes)) txBody.appendChild(doc.importNode(paragraph, true))
}

/**
 * The `p:txBody` of a notes slide's body placeholder (`p:ph type="body"`), or `null`
 * when the part carries none. A raw walk rather than a `NotesSlide`/`TextFrame` hop,
 * because the caller needs the element itself to rewrite and the modeled surface
 * would resolve a notes theme context it has no use for.
 * @param {Element | null} root - the notes slide's `p:notes` root
 * @return {Element | null} the body placeholder's text body
 */
function notesBodyTxBody(root: Element | null): Element | null {
	const cSld = root && firstChild(root, 'p:cSld')
	const spTree = cSld && firstChild(cSld, 'p:spTree')
	if (!spTree) return null
	for (const sp of getElements(spTree, 'p:sp')) {
		const nvSpPr = firstChild(sp, 'p:nvSpPr')
		const nvPr = nvSpPr && firstChild(nvSpPr, 'p:nvPr')
		const ph = nvPr && firstChild(nvPr, 'p:ph')
		if (ph && attr(ph, 'type') === 'body') return firstChild(sp, 'p:txBody')
	}
	return null
}

/**
 * Give `slide` the speaker notes `text`, creating its notes slide part when it has
 * none, and return the notes part's name.
 *
 * When the slide already has a notes part, only the body paragraphs are rewritten
 * (see {@link replaceNotesParagraphs}). When it has none, the part is built and
 * wired: `slide → notesSlide`, and on the new part the two relationships a notes
 * slide always carries — `notesMaster` as rId1 and a `slide` back-reference as rId2,
 * the ids and order the write path reserves.
 *
 * The notes master is resolved through {@link ensureNotesMasterForAuthoring}: this
 * deck's own when it has one, otherwise a fresh one bound to a clone of the deck's
 * theme. That is the single-notesMaster-per-presentation rule the import and append
 * paths already follow, so the three cannot produce a second master between them.
 * @param {Slide} slide - the slide to annotate
 * @param {string} text - the note, `\n` separating paragraphs
 * @return {string} the notes slide's partname
 */
export function authorNotes(slide: Slide, text: string): string {
	const opc = slide.presentation.opc

	const existingRel = slide.relationships.byType(NOTES_SLIDE_REL)[0]
	if (existingRel) {
		const partName = slide.relationships.resolveTarget(existingRel.id)
		const part = opc.part(partName)
		if (!part)
			throw new PackageReadError(
				'package/part-missing',
				`addNotes: slide ${slide.partName} has a notesSlide rel to a missing part ${partName}`
			)
		const txBody = notesBodyTxBody(part.dom.documentElement)
		if (!txBody)
			throw new PackageReadError(
				'package/part-has-no-root',
				`addNotes: notes slide ${partName} has no body placeholder to write into`
			)
		replaceNotesParagraphs(txBody, text)
		part.markDirty()
		return partName
	}

	// No notes part yet: build one. The slide-number field caches this slide's
	// 1-based position, the value the write path stamps; PowerPoint recomputes it.
	const slideNum = slide.index + 1
	const partName = opc.reservePartNameLike('/ppt/notesSlides/notesSlide1.xml')
	opc.addPart(
		partName,
		NOTES_SLIDE_CONTENT_TYPE,
		textEncoder.encode(makeXmlNotesSlideSkeleton(notesParagraphsXml(text), slideNum))
	)

	slide.relationships.add(NOTES_SLIDE_REL, relativePartName(slide.partName, partName))

	const notesRels = opc.relationshipsFor(partName)
	const masterPartName = ensureNotesMasterForAuthoring(slide.presentation, slide.partName)
	notesRels.addWithId('rId1', NOTES_MASTER_REL, relativePartName(partName, masterPartName))
	notesRels.addWithId('rId2', SLIDE_REL, relativePartName(partName, slide.partName))

	return partName
}
