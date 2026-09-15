/**
 * ts-pptx: `[Content_Types].xml`
 *
 * Emit the package content-types part: Default entries for the media extensions
 * actually used by the deck (plus the font default when faces are embedded) and
 * Override entries for every part the package skeleton writes.
 *
 * Entries that exist only because a construct family put a part in the package arrive as
 * {@link ContentTypeContributions} — already resolved to data and grouped by where they land.
 * The order below is a property of the format, not of whoever assembled that list, so a deck
 * written by a program that links three families and one that links ten emit the same bytes.
 */

import { CRLF, XML_DECL } from '../../constants-internal.js'
import type {
	PresSlideInternal,
	SlideLayoutInternal,
	SlideMasterInternal,
	SlideRelMedia,
} from '../../types/internal.js'
import { avContentType } from '../../media/content-type.js'
import { type EmbeddedFont, FONT_DATA_CONTENT_TYPE, FONT_DATA_EXTENSION } from '../../embedded-fonts.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { overrideName, PRESENTATION_PATH, slideLayoutPath, SLIDE_MASTER_PATH, slidePath } from './part-paths.js'
import { OOXML_NS } from '../../ooxml/namespaces.js'
import {
	CORE_PROPS_CONTENT_TYPE,
	CUSTOM_PROPS_CONTENT_TYPE,
	EXTENDED_PROPS_CONTENT_TYPE,
	PRES_PROPS_CONTENT_TYPE,
	PRESENTATION_MAIN_CONTENT_TYPE,
	RELATIONSHIPS_CONTENT_TYPE,
	SLIDE_CONTENT_TYPE,
	SLIDE_LAYOUT_CONTENT_TYPE,
	SLIDE_MASTER_CONTENT_TYPE,
	TABLE_STYLES_CONTENT_TYPE,
	THEME_CONTENT_TYPE,
	VIEW_PROPS_CONTENT_TYPE,
} from '../../ooxml/rel-types.js'

/**
 * Some Override entries have always been emitted with a leading space. It is insignificant
 * whitespace, but it is in the bytes this library has shipped for years, so it is reproduced
 * verbatim rather than normalized (see AGENTS.md "Byte identity": whitespace diffs are a STOP).
 */
const LEADING_SPACE = { openPrefix: ' ' }

/** One `Default` entry: a file extension, and the content type every part named with it carries. */
export interface ContentTypeDefault {
	readonly extension: string
	readonly contentType: string
}

/**
 * One `Override` entry: an absolute part name and its content type. `leadingSpace` asks for the
 * historic insignificant space before the tag — see {@link LEADING_SPACE}; which entries carry it
 * is a fact about the bytes already shipped, not a style choice.
 */
export interface ContentTypeOverride {
	readonly partName: string
	readonly contentType: string
	readonly leadingSpace?: boolean
}

/**
 * The entries a deck needs beyond the skeleton every package has, grouped by the point in the
 * emission order they belong to. `perSlide` and `perLayout` are index-aligned with the deck's
 * slides and layouts; an entry is emitted immediately after that slide's or layout's own Override.
 */
export interface ContentTypeContributions {
	/** Default entries, merged after the deck's own media extensions. */
	readonly defaults: readonly ContentTypeDefault[]
	/** Overrides sitting between the presentation part and the slide master. */
	readonly presentation: readonly ContentTypeOverride[]
	readonly perSlide: ReadonlyArray<readonly ContentTypeOverride[]>
	readonly perLayout: ReadonlyArray<readonly ContentTypeOverride[]>
	/** Overrides after the layout block, before docProps closes the part. */
	readonly trailing: readonly ContentTypeOverride[]
}

/** Nothing beyond the skeleton — the shape a deck written with no construct families contributes. */
export const NO_CONTENT_TYPE_CONTRIBUTIONS: ContentTypeContributions = Object.freeze({
	defaults: [],
	presentation: [],
	perSlide: [],
	perLayout: [],
	trailing: [],
})

function contentDefault(extension: string, contentType: string): string {
	return voidEl('Default', { Extension: extension, ContentType: contentType })
}

function override(partName: string, contentType: string, fmt?: { openPrefix: string }): string {
	return voidEl('Override', { PartName: partName, ContentType: contentType }, fmt)
}

/** Render a contributed Override, honouring its leading-space flag. */
function contributedOverride(entry: ContentTypeOverride): string {
	return override(entry.partName, entry.contentType, entry.leadingSpace ? LEADING_SPACE : undefined)
}

/**
 * Generate XML ContentType
 * @param opts - the deck's parts, plus whatever the construct families it was written with add
 * @returns XML
 */
export function makeXmlContTypes(opts: {
	slides: PresSlideInternal[]
	slideLayouts: SlideLayoutInternal[]
	masterSlide?: SlideMasterInternal
	hasCustomProps?: boolean
	embeddedFonts?: EmbeddedFont[]
	contributions?: ContentTypeContributions
}): string {
	const { slides, slideLayouts, masterSlide, hasCustomProps, embeddedFonts } = opts
	const contributions = opts.contributions ?? NO_CONTENT_TYPE_CONTRIBUTIONS
	const parts: string[] = [contentDefault('xml', 'application/xml'), contentDefault('rels', RELATIONSHIPS_CONTENT_TYPE)]

	// STEP 1 - Emit Default Extension entries only for media types actually used by the deck.
	// Walk slides + slideLayouts + masterSlide _relsMedia[] and dedupe by extension.
	// Skip 'online' rels (no part written) and rels missing extn/type.
	const extnTypeMap = new Map<string, string>()
	const ctTargets: Array<{ _relsMedia?: SlideRelMedia[] }> = []
	;(slides || []).forEach((s) => ctTargets.push(s))
	;(slideLayouts || []).forEach((l) => ctTargets.push(l))
	if (masterSlide) ctTargets.push(masterSlide)
	ctTargets.forEach((target) => {
		;(target._relsMedia || []).forEach((rel) => {
			if (rel.type === 'online' || !rel.extn || !rel.type) return
			// A/V rel `type` is `${mtype}/${extn}` (e.g. `audio/mp3`); resolve the part's
			// Default content type to what PowerPoint authors (`audio/mpeg`). Image rels
			// already carry their final content type (imageContentType).
			const contentType = rel.type.startsWith('audio/')
				? avContentType(rel.extn, 'audio')
				: rel.type.startsWith('video/')
					? avContentType(rel.extn, 'video')
					: rel.type
			if (!extnTypeMap.has(rel.extn)) extnTypeMap.set(rel.extn, contentType)
		})
	})
	extnTypeMap.forEach((type, extn) => {
		parts.push(contentDefault(extn, type))
	})
	// Contributed Defaults land after the deck's own media extensions, and only if nothing has
	// claimed the extension: one Extension may appear once, and an OLE payload can beat a chart's
	// embedded workbook to `xlsx`.
	contributions.defaults.forEach((entry) => {
		if (extnTypeMap.has(entry.extension)) return
		extnTypeMap.set(entry.extension, entry.contentType)
		parts.push(contentDefault(entry.extension, entry.contentType))
	})
	// Embedded fonts: one Default covers every `.fntdata` part (emitted only when fonts are embedded).
	if ((embeddedFonts || []).some((font) => font.faces.some((face) => face.bytes))) {
		parts.push(contentDefault(FONT_DATA_EXTENSION, FONT_DATA_CONTENT_TYPE))
	}

	// STEP 2: Add presentation and slide master(s)/slide(s)
	parts.push(override(overrideName(PRESENTATION_PATH), PRESENTATION_MAIN_CONTENT_TYPE))
	contributions.presentation.forEach((entry) => parts.push(contributedOverride(entry)))
	// Only one slideMaster part (`slideMaster1.xml`) is written; emit a single matching Override
	// rather than one per slide (which would dangle, since `slideMaster2..N.xml` do not exist).
	parts.push(override(overrideName(SLIDE_MASTER_PATH), SLIDE_MASTER_CONTENT_TYPE))
	slides.forEach((_slide, idx) => {
		parts.push(override(overrideName(slidePath(idx + 1)), SLIDE_CONTENT_TYPE))
		;(contributions.perSlide[idx] || []).forEach((entry) => parts.push(contributedOverride(entry)))
	})

	// STEP 3: Core PPT
	parts.push(override('/ppt/presProps.xml', PRES_PROPS_CONTENT_TYPE))
	parts.push(override('/ppt/viewProps.xml', VIEW_PROPS_CONTENT_TYPE))
	parts.push(override('/ppt/theme/theme1.xml', THEME_CONTENT_TYPE))
	// notesMaster1.xml.rels references ../theme/theme2.xml; emit a matching Override so the part resolves
	parts.push(override('/ppt/theme/theme2.xml', THEME_CONTENT_TYPE))
	parts.push(override('/ppt/tableStyles.xml', TABLE_STYLES_CONTENT_TYPE))

	// STEP 4: Add Slide Layouts
	slideLayouts.forEach((_layout, idx) => {
		parts.push(override(overrideName(slideLayoutPath(idx + 1)), SLIDE_LAYOUT_CONTENT_TYPE))
		;(contributions.perLayout[idx] || []).forEach((entry) => parts.push(contributedOverride(entry)))
	})

	// STEP 5: Everything the deck's construct families put in the package past this point.
	contributions.trailing.forEach((entry) => parts.push(contributedOverride(entry)))

	// LAST: Finish XML (Resume core)
	parts.push(override('/docProps/core.xml', CORE_PROPS_CONTENT_TYPE, LEADING_SPACE))
	parts.push(override('/docProps/app.xml', EXTENDED_PROPS_CONTENT_TYPE, LEADING_SPACE))
	if (hasCustomProps) parts.push(override('/docProps/custom.xml', CUSTOM_PROPS_CONTENT_TYPE, LEADING_SPACE))

	return XML_DECL + CRLF + el('Types', { xmlns: OOXML_NS.ct }, parts.map(raw))
}
