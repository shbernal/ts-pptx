/**
 * Carrying speaker notes and the deck's single notes master across an import or an append.
 *
 * A `notesSlide` is not self-contained: it back-references the slide it annotates and forward-
 * references the deck's notes master. `CT_NotesMasterIdList` allows at most one master per
 * presentation, so a deck that already has one keeps it — the destination's notes styling wins —
 * and a source master is copied only into a deck that has none. That single rule is why both the
 * `Presentation`-to-`Presentation` path and the generator-XML path funnel through
 * `registerNotesMaster`: they differ only in how the *part* comes into being.
 */

import { createElement, getOrAddChild, removeChildrenByQName, setAttr } from '../../oxml/dom.js'
import { relativePartName } from '../../opc/partnames.js'
import { copyPart, newOwnedScope, type ImportContext } from './part-copy.js'
import { isSharedByPageCopies } from './page-owned.js'
import type { Presentation } from '../presentation.js'
import {
	NOTES_MASTER_CONTENT_TYPE,
	NOTES_MASTER_REL,
	NOTES_SLIDE_REL,
	SLIDE_REL,
	THEME_CONTENT_TYPE,
	THEME_REL,
} from '../../../ooxml/rel-types.js'
import { PackageReadError } from '../../../errors.js'
import { makeXmlNotesMaster } from '../../../gen/slide/notes.js'
import { resolveSlideThemeParts } from '../theme-context.js'
import { presentationRels } from './deck-target.js'

const textEncoder = new TextEncoder()

/**
 * Carry the source slide's speaker notes onto the just-imported slide (the
 * `importNotes` option). The slide copy itself dropped the `notesSlide` rel, so
 * this copies the source `notesSlide` part into a fresh partname, wires a
 * `slide → notesSlide` rel on the new slide, and rebuilds the copied notesSlide's
 * own relationships:
 *
 * - its `slide` back-rel is repointed at the new slide (`newSlidePartName`) — the
 *   source slide is *not* copied (that would be circular);
 * - its `notesMaster` rel is resolved through {@link ensureNotesMaster}, which
 *   reuses this deck's notesMaster when it has one and copies the source's only
 *   when it has none (a deck may have at most one notesMaster);
 * - any other internal target (media, etc.) is copied via {@link copyPart}, under an
 *   ownership scope so a second copy of the page gets notes parts of its own.
 *
 * No-op when the source slide has no notes. Content-type registration for the
 * copied parts is handled by `addPart`/{@link copyPart}.
 */
export function carryNotes(
	dest: Presentation,
	source: Presentation,
	ctx: ImportContext,
	sourceSlidePartName: string,
	newSlidePartName: string
): void {
	const sourceSlideRels = source.opc.relationshipsFor(sourceSlidePartName)
	const notesRel = sourceSlideRels.byType(NOTES_SLIDE_REL)[0]
	if (!notesRel) return // slide has no speaker notes
	const sourceNotesPartName = sourceSlideRels.resolveTarget(notesRel.id)
	const sourceNotesPart = source.opc.part(sourceNotesPartName)
	if (!sourceNotesPart) return

	// Copy the notesSlide bytes into a fresh partname, then wire slide → notesSlide.
	const newNotesPartName = dest.opc.reservePartNameLike(sourceNotesPartName)
	dest.opc.addPart(newNotesPartName, sourceNotesPart.contentType, sourceNotesPart.bytes)
	dest.opc.relationshipsFor(newSlidePartName).add(NOTES_SLIDE_REL, relativePartName(newSlidePartName, newNotesPartName))

	// A notes slide is a part its page *owns* (see `page-owned.ts`), and so is
	// whatever it owns in turn. Two copies of one source page — `importSlide`
	// called twice, or one `importSlides` batch naming the page twice — each get
	// notes of their own here; without a scope the second copy would share the
	// first's OLE embedding or chart through the registry, which is a package
	// PowerPoint refuses to open. Media stays shared, as everywhere else.
	const owned = newOwnedScope()
	owned.set(sourceNotesPartName, newNotesPartName)

	// Rebuild the copied notesSlide's relationships. Preserve each source rel id so
	// the notesSlide body's r:id references stay valid; only the targets are rewritten.
	const notesSourceRels = source.opc.relationshipsFor(sourceNotesPartName)
	const notesTargetRels = dest.opc.relationshipsFor(newNotesPartName)
	for (const rel of notesSourceRels) {
		if (rel.type === SLIDE_REL) {
			// Back-reference to the annotated slide → repoint at the new slide (don't copy it).
			notesTargetRels.addWithId(rel.id, SLIDE_REL, relativePartName(newNotesPartName, newSlidePartName))
			continue
		}
		if (rel.type === NOTES_MASTER_REL) {
			const notesMaster = ensureNotesMaster(dest, ctx, notesSourceRels.resolveTarget(rel.id))
			notesTargetRels.addWithId(rel.id, NOTES_MASTER_REL, relativePartName(newNotesPartName, notesMaster))
			continue
		}
		if (rel.targetMode === 'External') {
			notesTargetRels.addWithId(rel.id, rel.type, rel.target, 'External')
			continue
		}
		const newTarget = copyPart(
			ctx,
			notesSourceRels.resolveTarget(rel.id),
			isSharedByPageCopies(rel.type) ? undefined : owned
		)
		notesTargetRels.addWithId(rel.id, rel.type, relativePartName(newNotesPartName, newTarget))
	}
}

/**
 * Resolve the notesMaster an imported `notesSlide` should bind to, honouring the
 * single-notesMaster-per-presentation rule (`p:notesMasterIdLst` holds 0..1
 * `p:notesMasterId`). If this deck already has a notesMaster it is reused and the
 * source's is *not* copied (the destination's notes styling wins); otherwise the
 * source notesMaster (and, via {@link copyPart}, its theme) is copied and
 * registered in `presentation.xml`. Returns the destination notesMaster partname.
 */
function ensureNotesMaster(dest: Presentation, ctx: ImportContext, sourceNotesMasterPartName: string): string {
	const presRels = presentationRels(dest)
	const existing = presRels.byType(NOTES_MASTER_REL)[0]
	if (existing) return presRels.resolveTarget(existing.id)

	// No notesMaster yet: copy the source's (pulls its theme) and register it.
	return registerNotesMaster(dest, copyPart(ctx, sourceNotesMasterPartName))
}

/**
 * Wire an already-added notesMaster part into `presentation.xml`: a `notesMaster`
 * relationship plus the single `p:notesMasterId` entry that `CT_NotesMasterIdList`
 * allows. Returns the partname, so callers can use it as a rel target.
 *
 * Split out of {@link ensureNotesMaster} because the two ways a notesMaster arrives
 * — copied from another `Presentation`, or authored by a generator and injected by
 * {@link Presentation.appendSlides} — differ only in how the *part* is created, not in how it is
 * registered.
 */
function registerNotesMaster(dest: Presentation, notesMasterPartName: string): string {
	const presPart = dest.presentationPart
	const presRels = presentationRels(dest)
	const relId = presRels.add(NOTES_MASTER_REL, relativePartName(presPart.partName, notesMasterPartName)).id

	const root = presPart.dom.documentElement
	if (!root)
		throw new PackageReadError(
			'package/part-has-no-root',
			'presentation.xml has no document element to register a notes master in'
		)
	// `p:notesMasterIdLst` follows `p:sldMasterIdLst` in CT_Presentation order.
	const lst = getOrAddChild(root, 'p:notesMasterIdLst', [
		'p:handoutMasterIdLst',
		'p:sldIdLst',
		'p:sldSz',
		'p:notesSz',
		'p:embeddedFontLst',
		'p:custShowLst',
		'p:photoAlbum',
		'p:custDataLst',
		'p:kinsoku',
		'p:defaultTextStyle',
		'p:modifyVerifier',
		'p:extLst',
	])
	// CT_NotesMasterIdList holds a single p:notesMasterId; replace any stray entry.
	removeChildrenByQName(lst, ['p:notesMasterId'])
	const entry = createElement(presPart.dom, 'p:notesMasterId')
	setAttr(entry, 'r:id', relId)
	lst.appendChild(entry)
	presPart.markDirty()
	return notesMasterPartName
}

/**
 * Resolve the notesMaster an *appended* slide's notes should bind to. Same
 * single-notesMaster rule as {@link ensureNotesMaster}: this deck's own wins when it
 * has one, so the destination's notes styling is preserved and `master.xml` is
 * discarded. Otherwise the generator's notes master is installed, together with the
 * theme its `.rels` requires (the normal write path emits that as `theme2.xml`).
 */
export function ensureNotesMasterFromXml(dest: Presentation, master: { xml: string; themeXml: string }): string {
	const presRels = presentationRels(dest)
	const existing = presRels.byType(NOTES_MASTER_REL)[0]
	if (existing) return presRels.resolveTarget(existing.id)

	const masterPartName = dest.opc.reservePartNameLike('/ppt/notesMasters/notesMaster1.xml')
	dest.opc.addPart(masterPartName, NOTES_MASTER_CONTENT_TYPE, textEncoder.encode(master.xml))

	// A notesMaster's .rels must resolve a theme; reserve alongside any theme the
	// destination already owns rather than assuming theme2.xml is free.
	const themePartName = dest.opc.reservePartNameLike('/ppt/theme/theme1.xml')
	dest.opc.addPart(themePartName, THEME_CONTENT_TYPE, textEncoder.encode(master.themeXml))
	dest.opc.relationshipsFor(masterPartName).add(THEME_REL, relativePartName(masterPartName, themePartName))

	return registerNotesMaster(dest, masterPartName)
}

/**
 * Resolve the notesMaster a *newly authored* notes slide should bind to
 * ({@link Slide.addNotes} on a loaded deck). Same single-notesMaster rule as
 * {@link ensureNotesMaster} and {@link ensureNotesMasterFromXml}: this deck's own
 * wins when it has one. The third way a notesMaster arrives, and the only one with
 * no source deck and no generator instance behind it, so it has to build both parts
 * itself.
 *
 * The *theme* is where it differs from the append path. `ensureNotesMasterFromXml`
 * installs the generator's `theme2.xml` because a generator deck's notes were
 * authored against it; there is no such theme here, so the deck's **own** theme
 * (reached through the annotated slide's layout → master → theme chain) is cloned
 * into a fresh part instead. That is both the closer match to "the destination's
 * notes styling wins" and the safer package: the notesMaster gets a theme part of
 * its own rather than a second relationship onto the slide master's, which is the
 * arrangement PowerPoint writes (`theme1.xml` for the master, `theme2.xml` for the
 * notes master).
 *
 * Throws when the deck has no theme to clone — a package that malformed cannot be
 * given a valid notesMaster, and silently emitting one without a resolvable theme
 * is what PowerPoint reports as a repair prompt rather than a bad edit.
 * @param {Presentation} dest - the deck being authored onto
 * @param {string} slidePartName - the slide whose notes are being authored, used to find the theme
 * @return {string} the destination notesMaster partname
 */
export function ensureNotesMasterForAuthoring(dest: Presentation, slidePartName: string): string {
	const presRels = presentationRels(dest)
	const existing = presRels.byType(NOTES_MASTER_REL)[0]
	if (existing) return presRels.resolveTarget(existing.id)

	const themePartName = resolveSlideThemeParts(dest.opc, slidePartName).themePartName
	const themePart = themePartName ? dest.opc.part(themePartName) : undefined
	if (!themePart)
		throw new PackageReadError(
			'package/part-missing',
			`addNotes: no theme reachable from ${slidePartName} to bind a new notes master to`
		)

	const masterPartName = dest.opc.reservePartNameLike('/ppt/notesMasters/notesMaster1.xml')
	dest.opc.addPart(masterPartName, NOTES_MASTER_CONTENT_TYPE, textEncoder.encode(makeXmlNotesMaster()))

	// A notesMaster's .rels must resolve a theme; clone the deck's rather than share
	// the slide master's part. Reserved alongside any theme this deck already owns.
	const notesThemePartName = dest.opc.reservePartNameLike('/ppt/theme/theme1.xml')
	dest.opc.addPart(notesThemePartName, themePart.contentType, themePart.bytes)
	dest.opc.relationshipsFor(masterPartName).add(THEME_REL, relativePartName(masterPartName, notesThemePartName))

	return registerNotesMaster(dest, masterPartName)
}
