/**
 * Where the parts of a written package live.
 *
 * Every part is named at least three times: the assembler writes it, `[Content_Types].xml`
 * declares an `Override` for it, and one or more `.rels` parts target it. Each of those was
 * its own template literal in its own module, and the three spellings differ — a zip entry
 * name, a leading-slash part name, a relationship target resolved against the part that
 * states it — so nothing tied them together. A part written but not declared is a
 * PowerPoint repair, and so is a relationship pointing at a part nobody wrote.
 *
 * The canonical spelling here is the zip entry name: `ppt/slides/slide1.xml`, no leading
 * slash, which is what {@link ZipWriter.add} takes. The other two are derivations —
 * {@link overrideName} for the content-types declaration, {@link targetFromPresentation} and
 * {@link targetFromPptSubpart} for a relationship target — so a path can only be spelled once
 * and then transformed, never re-typed in a second dialect.
 *
 * Indices are 1-based throughout, matching the OPC naming rather than the arrays that drive it.
 */

import { FONT_DATA_EXTENSION } from '../../embedded-fonts.js'

/** `ppt/presentation.xml`, the package's root presentation part. */
export const PRESENTATION_PATH = 'ppt/presentation.xml'

/** `ppt/slideMasters/slideMaster1.xml`. Exactly one master part is written. */
export const SLIDE_MASTER_PATH = 'ppt/slideMasters/slideMaster1.xml'

/** `ppt/notesMasters/notesMaster1.xml`. Exactly one notes-master part is written. */
export const NOTES_MASTER_PATH = 'ppt/notesMasters/notesMaster1.xml'

/** The `n`-th slide part. */
export function slidePath(n: number): string {
	return `ppt/slides/slide${n}.xml`
}

/** The `n`-th slide-layout part. */
export function slideLayoutPath(n: number): string {
	return `ppt/slideLayouts/slideLayout${n}.xml`
}

/**
 * The `n`-th notes-slide part. One is written for every slide, empty notes included, so the
 * index is the slide's own — see the assembler, which keeps the two enumerations in step.
 */
export function notesSlidePath(n: number): string {
	return `ppt/notesSlides/notesSlide${n}.xml`
}

/**
 * The comment part for the `n`-th slide. Written only for slides that carry comments, so the
 * index is a slide number with gaps in it rather than a running count of comment parts.
 */
export function commentPath(n: number): string {
	return `ppt/comments/comment${n}.xml`
}

/** The `n`-th embedded-font part, holding one whole face's raw bytes. */
export function fontPath(n: number): string {
	return `ppt/fonts/font${n}.${FONT_DATA_EXTENSION}`
}

/**
 * The `.rels` part that carries `path`'s relationships: `ppt/slides/_rels/slide1.xml.rels`.
 * A part with no relationships has no such part; this only says where one would go.
 */
export function relsPath(path: string): string {
	const cut = path.lastIndexOf('/')
	return `${path.slice(0, cut)}/_rels/${path.slice(cut + 1)}.rels`
}

/**
 * `path` as `[Content_Types].xml` spells it in an `Override/@PartName`: the same path with the
 * leading slash ECMA-376 Part 2 requires of a part name.
 */
export function overrideName(path: string): string {
	return `/${path}`
}

/**
 * `path` as a relationship target stated by a part directly under `ppt/` — presentation.xml,
 * whose rels reach `slides/slide1.xml` with no traversal at all.
 */
export function targetFromPresentation(path: string): string {
	return path.slice('ppt/'.length)
}

/**
 * `path` as a relationship target stated by a part one level below `ppt/` — a slide, a layout,
 * a notes slide, the master: `../slides/slide1.xml`.
 */
export function targetFromPptSubpart(path: string): string {
	return `../${targetFromPresentation(path)}`
}
