/**
 * The deck's shared chrome → {@link ChromeIr}: theme colours and fonts, plus one
 * `defineSlideMaster` descriptor per source layout.
 *
 * **Only a standalone output needs any of this,** and it is the part of a deck that survives
 * a conversion worst. A template-anchored output reuses the source package, so its masters,
 * layouts and theme are the originals byte for byte; everything below is the approximation
 * that replaces them when there is no template to inherit from. The notes it records are
 * therefore tier-scoped: real for a standalone script, and suppressed by the template-anchored
 * printer because they do not describe its output.
 *
 * **Three ceilings meet here, and none of them moves by working harder on this file:**
 *
 * 1. `a:fmtScheme` is unreachable from *both* directions — no read accessor, and the write
 *    path emits a hardcoded Office one. A shape whose look comes from `p:style` is repainted.
 * 2. `p:txStyles` has a write counterpart (`SlideMasterProps.textStyles`) and no reader, so
 *    the per-level placeholder defaults are invisible.
 * 3. Master and layout *decoration* — anything that is not a placeholder — is documented as
 *    out of the read model's scope (`src/read/api/chrome.ts` header); the import paths carry
 *    it byte-for-byte instead. A logo on a layout cannot be seen, let alone re-authored.
 *
 * **Layout placeholders are deliberately not emitted, and that is a write-path finding rather
 * than a read-path one.** `addPlaceholdersToSlideLayouts` seeds *every* slide with each layout
 * placeholder the slide did not populate, as an empty text shape
 * (`src/gen/define/placeholder.ts`). Since this converter transcribes every source shape as
 * concrete absolute-positioned content and never binds one to a placeholder, reproducing the
 * layout's placeholders would add a ghost shape to every slide for each of them — shapes the
 * source deck does not have. Dropping them costs an editing affordance; emitting them would
 * cost the output's shape tree.
 */
import type { Presentation } from '../../read/api/presentation.js'
import type { SlideMaster, Theme } from '../../read/api/chrome.js'
import type { SlideBackground } from '../../read/api/slide-background.js'
import type { ChromeIr, IrValue, MasterIr, ThemeIr } from '../ir.js'
import type { NoteScope } from '../fidelity.js'
import type { AssetResolver } from './shape.js'
import { hasDecorativeShapes, hasFormatScheme, hasTextStyles } from './detect.js'
import { compact, literalColor, orUndefined } from './values.js'

/**
 * The `p:clrMap` the write path always emits (`src/gen/slide/master.ts`). A source master
 * that maps a token elsewhere — a deck whose light and dark slots are swapped, the usual
 * case — resolves every `schemeClr` in the deck differently, and there is no setter for it.
 */
const DEFAULT_COLOR_MAP: Record<string, string> = {
	bg1: 'lt1',
	tx1: 'dk1',
	bg2: 'lt2',
	tx2: 'dk2',
	accent1: 'accent1',
	accent2: 'accent2',
	accent3: 'accent3',
	accent4: 'accent4',
	accent5: 'accent5',
	accent6: 'accent6',
	hlink: 'hlink',
	folHlink: 'folHlink',
}

/** Build the chrome IR for a deck. */
export function chromeToIr(pres: Presentation, notes: NoteScope, assets: AssetResolver): ChromeIr {
	const masters = pres.masters()

	if (masters.length > 1) {
		notes.note(
			'master.multiple',
			'flattened',
			'unsupported',
			`this deck has ${masters.length} slide masters; the write API models a deck as one shared master with many layouts, so every layout is rebuilt under a single master and only the first master's theme and colour map survive`
		)
	}

	const first = masters[0]
	const theme = first?.theme ?? null
	recordMasterLosses(masters, theme, notes)

	return { theme: themeToIr(theme), masters: layoutsToIr(masters, notes, assets) }
}

/** The theme's colour scheme and font faces, in `ThemeProps` spelling. */
function themeToIr(theme: Theme | null): ThemeIr {
	if (!theme) return {}

	const scheme = theme.colorScheme
	const colors: Record<string, string> = {}
	for (const [slot, hex] of Object.entries(scheme)) {
		if (hex !== null) colors[slot] = literalColor(hex)
	}

	const fonts = theme.fontScheme
	return compact({
		headFontFace: orUndefined(fonts?.major.latin ?? null),
		bodyFontFace: orUndefined(fonts?.minor.latin ?? null),
		headFontFaceEA: orUndefined(fonts?.major.ea ?? null),
		bodyFontFaceEA: orUndefined(fonts?.minor.ea ?? null),
		headFontFaceCS: orUndefined(fonts?.major.cs ?? null),
		bodyFontFaceCS: orUndefined(fonts?.minor.cs ?? null),
		colorScheme: Object.keys(colors).length > 0 ? colors : undefined,
	}) as ThemeIr
}

/** The losses that belong to the master/theme tier as a whole rather than to one layout. */
function recordMasterLosses(masters: SlideMaster[], theme: Theme | null, notes: NoteScope): void {
	if (theme && hasFormatScheme(theme.element_)) {
		notes.note(
			'theme.fmtScheme',
			'approximated',
			'unread',
			"the theme's format scheme (a:fmtScheme — the three fill, line and effect style lists a shape's p:style indexes into) has no read accessor and no write-API option, so the output carries Office's default; a shape whose fill or outline came from the theme rather than from its own a:ln keeps its colour but not its width, dash or effect"
		)
	}

	for (const master of masters) {
		const root = master.part.dom.documentElement
		if (hasTextStyles(root)) {
			notes.note(
				'master.txStyles',
				'dropped',
				'unread',
				"the master's per-level text styles (p:txStyles — the default size, face, colour, indent and bullet of each of the nine list levels) have no read accessor, so placeholder text falls back to PowerPoint's built-in defaults; SlideMasterProps.textStyles could author them if they could be seen"
			)
		}
		if (hasDecorativeShapes(root)) {
			notes.note(
				'master.decoration',
				'dropped',
				'unwritable',
				'this slide master carries shapes that are not placeholders (a logo, a rule, a background graphic); the read model now decodes them (SlideMaster.shapes), but defineSlideMaster authors a layout rather than a master, and its objects union expresses none of what real master decoration is made of — groups, custom geometry, effects — so they cannot be re-authored'
			)
		}

		const map = master.colorMap
		const remapped = Object.entries(DEFAULT_COLOR_MAP).filter(
			([token, slot]) => map[token as keyof typeof map] !== slot
		)
		if (remapped.length > 0) {
			notes.note(
				'master.colorMap',
				'dropped',
				'unwritable',
				`this master remaps ${remapped.map(([token]) => token).join(', ')} to a different theme slot and the write path always emits the identity map, so every scheme colour that passes through as a token resolves to a different hex than it did in the source`
			)
		}
	}
}

/**
 * One `defineSlideMaster` descriptor per source layout, in the same gallery order
 * `Presentation.layouts()` reports — which is what lets a slide bind by index.
 *
 * Titles are made unique because they are the lookup key `addSlide({ masterTitle })` matches
 * on, and `p:cSld@name` is not unique: a multi-master deck routinely carries several layouts
 * called "Title and Content", and two `defineSlideMaster` calls sharing a title would make
 * every slide bind to whichever won.
 */
function layoutsToIr(masters: SlideMaster[], notes: NoteScope, assets: AssetResolver): MasterIr[] {
	const out: MasterIr[] = []
	const used = new Set<string>()
	// Rolled up rather than noted per layout: a 12-layout deck loses decoration and placeholder
	// definitions on most of them, and twelve near-identical paragraphs at the top of the
	// emitted script bury the per-shape notes underneath that a reader can actually act on.
	const decorated: string[] = []
	let placeholderCount = 0
	let placeholderLayouts = 0

	const renamed: string[] = []
	const collided: string[] = []

	for (const master of masters) {
		for (const layout of master.layouts) {
			const index = out.length
			const flattened = flattenAttributeValue(layout.name)
			if (flattened !== layout.name) renamed.push(`${JSON.stringify(layout.name)} → ${JSON.stringify(flattened)}`)
			const base = flattened || `Layout ${index + 1}`
			let title = base
			for (let suffix = 2; used.has(title); suffix++) title = `${base} (${suffix})`
			if (title !== base) collided.push(`${JSON.stringify(base)} → ${JSON.stringify(title)}`)
			used.add(title)

			if (hasDecorativeShapes(layout.part.dom.documentElement)) decorated.push(title)
			const placeholders = layout.placeholders.length
			if (placeholders > 0) {
				placeholderCount += placeholders
				placeholderLayouts++
			}

			// The layout's own background if it has one, else the master's — which is what a slide
			// bound to that layout actually shows, and the only tier the write API can set it at.
			const background = backgroundProps(layout.background ?? master.background, notes, assets)
			out.push({ layoutIndex: index, props: compact({ title, background }) as Record<string, IrValue> })
		}
	}

	if (collided.length > 0) {
		notes.note(
			'master.nameCollision',
			'approximated',
			'unsupported',
			`${collided.length} layout name(s) are not unique in the source — normally because the deck has several masters, each with its own "Title and Content" — and a layout title here is also the key addSlide({ masterTitle }) binds on, so duplicates cannot both exist. The later ones are suffixed and the slides bound to them follow: ${collided.join('; ')}`
		)
	}
	if (renamed.length > 0) {
		notes.note(
			'master.name',
			'flattened',
			'unwritable',
			`${renamed.length} layout name(s) contain a line break or tab, which cannot survive as a layout title: the write path emits the name as a literal XML attribute value rather than a character reference, and XML normalises whitespace in an attribute to spaces. Collapsed here so the emitted script says what the output will contain — ${renamed.join('; ')}`
		)
	}
	if (decorated.length > 0) {
		notes.note(
			'master.decoration',
			'dropped',
			'unwritable',
			`${decorated.length} slide layout(s) carry shapes that are not placeholders — ${decorated.join(', ')} — and while the read model now decodes them (SlideLayout.shapes), the converter does not map a layout's shape tree back into defineSlideMaster({ objects }): that union covers a plain rect, line, image, chart or text box, not the groups, custom geometry and effects real layout decoration is made of`
		)
	}
	if (placeholderCount > 0) {
		notes.note(
			'master.placeholders',
			'dropped',
			'unsupported',
			`the ${placeholderCount} placeholder definition(s) across ${placeholderLayouts} slide layout(s) are not reproduced: the write path seeds every slide with each layout placeholder the slide did not populate, and since every source shape is authored here as concrete content rather than into a placeholder, re-declaring them would add an empty shape to every slide rather than a prompt`
		)
	}

	return out
}

/**
 * Collapse the whitespace an XML attribute value cannot carry.
 *
 * Not a style choice: XML normalises a literal tab, carriage return or line feed inside an
 * attribute value to a space (XML 1.0 §3.3.3), and the write path emits a layout title as a
 * literal rather than as a `&#10;` character reference. So a PowerPoint layout named across
 * two lines — "Abschnitts-⏎überschrift" is a real one, in the German built-in set — comes back
 * with a space either way. Doing it here rather than leaving it to the serializer means the
 * emitted script states the title the output will actually have, and the round trip compares
 * two decks that agree instead of reporting a rename it cannot act on.
 */
function flattenAttributeValue(value: string): string {
	return value.replace(/[\t\r\n]+/g, ' ')
}

/** A master/layout background reduced to what `BackgroundProps` accepts. */
function backgroundProps(
	background: SlideBackground | null,
	notes: NoteScope,
	assets: AssetResolver
): IrValue | undefined {
	if (!background || background.type === 'none') return undefined

	switch (background.type) {
		case 'solid':
			return background.color ? { color: literalColor(background.color.effectiveHex) } : undefined
		case 'image': {
			const asset = background.partName === null ? null : assets.assetFor(background.partName)
			return asset ? { data: asset } : undefined
		}
		case 'themeRef': {
			// `p:bgRef` indexes the theme's background fill list, which the write path cannot
			// author. The read model resolves it, so the flat colour survives even though the
			// reference — and therefore its response to a theme change — does not.
			const fill = background.resolvedFill
			if (fill?.type === 'solid' && fill.color) {
				notes.note(
					'master.background',
					'flattened',
					'unwritable',
					"this layout's background is a theme reference (p:bgRef into the theme's background fill list), which has no write-API option; the colour it currently resolves to is baked in and stops following the theme"
				)
				return { color: literalColor(fill.color.effectiveHex) }
			}
			notes.note(
				'master.background',
				'dropped',
				'unwritable',
				"this layout's background is a theme reference to a non-solid fill, which has no write-API option; the layout is emitted with no background"
			)
			return undefined
		}
		default:
			notes.note(
				'master.background',
				'dropped',
				'unsupported',
				`a ${background.type} layout background is not expressible through the write API's background option, so the layout is emitted with no background`
			)
			return undefined
	}
}
