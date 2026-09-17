/**
 * `TextFrame` → `addText(TextProps[], TextPropsOptions)`.
 *
 * Two structural mismatches shape this mapper, and both are why the conversion is not a
 * field-by-field copy:
 *
 * 1. **The read model is a tree; `addText` takes a flat list.** OOXML nests runs inside
 *    paragraphs; `TextProps[]` is one flat array where a paragraph boundary is a
 *    `breakLine` flag on the item that *precedes* it. So the last run of each paragraph
 *    carries the break, and a paragraph's own properties are replicated onto each of its
 *    runs — the write path reads alignment and spacing off whichever run starts a line.
 *
 * 2. **Vocabularies differ on both sides of the same concept.** OOXML spells alignment
 *    `l`/`ctr`/`r`/`just`, the write API spells it `left`/`center`/`right`/`justify`; two
 *    of the six `@algn` values have no write spelling at all. Every such translation is a
 *    table here rather than a pass-through, so an unmappable token produces a note instead
 *    of an invalid option.
 *
 * Run-level reads prefer the *explicit* value over the resolved one, but only where leaving
 * the option out really does leave the value inherited. That is the part worth checking
 * rather than assuming: the round-trip harness found that an omitted run colour and an
 * omitted vertical anchor are both filled in by the write path — black, and centred — so
 * "stay quiet and let it inherit" was silently repainting text and re-anchoring bodies. Both
 * now emit the *resolved* value with a note, which freezes it against a later theme edit and
 * renders correctly, in preference to staying faithful in the IR and wrong on the slide.
 */
import type { BodyProperties, BulletDetail, BulletStyle, Paragraph, Run, TextFrame } from '../../read/api/text.js'
import type { Hyperlink } from '../../read/api/hyperlink.js'
import type { AnyShape } from '../../read/api/shapes.js'
import { BODY_INSET_DEFAULTS_PT } from '../../ooxml/body-insets.js'
import { TEXT_AUTONUM_SCHEMES, TEXT_VERTICAL } from '../../ooxml/st-enums.js'
import { HALIGN_BY_TEXT_ALIGN } from '../../ooxml/text-align.js'
import type { NoteScope } from '../fidelity.js'
import type { IrValue } from '../ir.js'
import type { MapContext } from './context.js'
import { hasEquation } from './detect.js'
import {
	ANCHOR_TO_VALIGN,
	colorOption,
	compact,
	literalColor,
	orUndefined,
	pointsToInches,
	schemeColorOption,
} from './values.js'

/**
 * `a:pPr/@algn` → `HAlign`. `dist` and `thaiDist` (distributed justification) are
 * deliberately absent: the write API has no spelling for them, and mapping them onto
 * `justify` would be a silent visual change rather than a declared loss.
 */
const ALIGN = HALIGN_BY_TEXT_ALIGN

/** `AutofitMode` → the write API's `fit`. */
const FIT: Record<string, string> = { none: 'none', normAutofit: 'shrink', spAutoFit: 'resize' }

/**
 * `a:bodyPr/@vert` values the write API's `vert` accepts. `TextVertType` is the whole
 * `ST_TextVerticalType`, so this is `TEXT_VERTICAL` itself, and `text.vert` fires only for a token
 * outside the schema.
 */
const WRITABLE_VERT: ReadonlySet<string> = new Set<string>(TEXT_VERTICAL)

/**
 * The `a:buAutoNum/@type` values the write API's `numberType` names.
 *
 * This used to be the *discriminator* between a numbered and a character bullet, back when
 * the read model reported both as one tagged string. `BulletDetail.kind` does that now, so
 * the lookup has one job left: telling a scheme the write path can spell from one it cannot.
 *
 * Built from the tuple the write API's `numberType` is derived from and its emitter checks, so
 * the two cannot disagree about a scheme.
 */
const AUTO_NUMBER_TYPES: ReadonlySet<string> = new Set(TEXT_AUTONUM_SCHEMES)

/**
 * `Paragraph.bulletDetail` → the write API's `bullet` option.
 *
 * This function used to take a **tagged** string — `'none'`, `'char:<glyph>'` or
 * `'autoNum:<type>'` — and its first version read it as a bare glyph. The result was silent
 * and universal: a paragraph that explicitly suppressed its bullet (`a:buNone`, which is most
 * of them) came back with a literal `n` bullet, because `'none'.codePointAt(0)` is `n`; a real
 * character bullet came back as `c`, from the `char:` tag; and a numbered list came back as
 * `a`. Every converted deck was affected and nothing failed, which is why the round-trip check
 * exists — and why the accessor is now a discriminated union with no parsing left to get wrong.
 */
function bulletOption(bullet: BulletDetail, notes: NoteScope): IrValue {
	// Explicit suppression, and it must stay explicit: an omitted `bullet` lets the
	// destination list style put one back.
	if (bullet.kind === 'none') return false

	// A picture bullet's bytes would have to be re-embedded through the asset resolver, which
	// the paragraph mapper does not carry — `addText`'s `bullet.image` could author it.
	if (bullet.kind === 'picture') {
		notes.note(
			'text.bullet.picture',
			'approximated',
			'unsupported',
			'this paragraph uses an image as its bullet glyph (a:buBlip); the text mapper has no asset resolver to re-embed it with, so the bullet falls back to the default character'
		)
		return true
	}

	const style = bulletStyle(bullet, notes)

	if (bullet.kind === 'autoNum') {
		if (!AUTO_NUMBER_TYPES.has(bullet.scheme)) {
			notes.note(
				'text.bullet.numberType',
				'approximated',
				'unwritable',
				`numbering scheme "${bullet.scheme}" is outside the set the write API names, so the list falls back to the default scheme`
			)
			return compact({ type: 'number', ...startAtOption(bullet.startAt), ...style }) ?? { type: 'number' }
		}
		return (
			compact({ type: 'number', numberType: bullet.scheme, ...startAtOption(bullet.startAt), ...style }) ?? {
				type: 'number',
			}
		)
	}

	// Zero-padded to four digits, and that is load-bearing rather than cosmetic: the write
	// path tests `characterCode` against /^[0-9A-Fa-f]{4}$/ and, on a miss, warns to the
	// console and substitutes its own default glyph. So "6E" — a perfectly good code point —
	// silently became a different bullet character, visible only in the rendered slide.
	const code = bullet.char.codePointAt(0) ?? 0x2022
	if (code > 0xffff) {
		notes.note(
			'text.bullet.glyph',
			'approximated',
			'unwritable',
			`bullet glyph U+${code.toString(16).toUpperCase()} is outside the Basic Multilingual Plane and characterCode takes a four-digit code, so the bullet falls back to the write path's default glyph`
		)
	}
	return compact({ characterCode: code.toString(16).toUpperCase().padStart(4, '0'), ...style }) ?? {}
}

/**
 * `a:pPr/@marL` / `@indent` → `paraMarginLeft` / `paraIndent`.
 *
 * A stated margin carries as its number. An *absent* one is the interesting case, because
 * leaving the key out does not reproduce it: every `bullet` state except `'inherit'` makes the
 * write path put a margin on the element — its hanging default for a drawn bullet, zero for
 * `bullet: false` — so a paragraph that inherits its margin has to ask for silence explicitly.
 * Under `bullet: 'inherit'` nothing is written either way and the key would be noise, which is
 * why the bullet option decides rather than the paragraph being asked twice.
 * @param {number|null} valuePt - the paragraph's own margin in points, `null` when unset
 * @param {IrValue} bulletValue - the `bullet` option this paragraph maps onto
 * @return {IrValue|undefined} the option value, or `undefined` to leave the key out
 */
function paraMarginOption(valuePt: number | null, bulletValue: IrValue): IrValue | undefined {
	if (valuePt !== null) return valuePt
	return bulletValue === 'inherit' ? undefined : 'inherit'
}

/**
 * `a:buAutoNum/@startAt` → `numberStartAt`.
 *
 * Numbering is content rather than styling: a list continuing "5. Deploy" that comes back as
 * "1. Deploy" is a different slide. `1` is the schema default, so emitting it would only add
 * noise to the printed script.
 */
function startAtOption(startAt: number | null): Record<string, IrValue> {
	return startAt === null || startAt === 1 ? {} : { numberStartAt: startAt }
}

/**
 * The bullet's own font, size and colour (`a:buFont` / `a:buSzPct` / `a:buClr`) as the write
 * API's `fontFace` / `size` / `color`.
 *
 * `a:buSzPts` has no write-API counterpart — `bullet.size` is a percentage of the run size —
 * so an absolute bullet size is the one part of this that still cannot carry.
 */
function bulletStyle(bullet: BulletStyle, notes: NoteScope): Record<string, IrValue> {
	if (bullet.sizePt !== null) {
		notes.note(
			'text.bullet.sizePt',
			'dropped',
			'unwritable',
			'this bullet sets an absolute glyph size (a:buSzPts); bullet.size is a percentage of the run size, so the absolute value has no write-API expression and the glyph follows the text size'
		)
	}

	// 25–400% is ST_TextBulletSizePercent's range, and the write path clamps an out-of-range
	// value to the nearest bound rather than accepting it. Print the clamped number so the
	// script says what the deck will get, and declare the difference — printing the source
	// value instead would make the script warn on every run.
	const sizePct = bullet.sizePct
	const clampedPct = sizePct === null ? null : Math.min(400, Math.max(25, sizePct))
	if (sizePct !== null && clampedPct !== sizePct) {
		notes.note(
			'text.bullet.sizePct',
			'approximated',
			'unwritable',
			`bullet glyph size ${sizePct}% is outside the 25–400% range the write API accepts, so it is printed as ${clampedPct}% — the bound the write path clamps it to`
		)
	}

	return (
		compact({
			fontFace: orUndefined(bullet.font),
			size: clampedPct ?? undefined,
			color: bulletColor(bullet, notes),
		}) ?? {}
	)
}

/**
 * The bullet's colour, preferring the raw `schemeClr` token over the resolved literal for the
 * same reason {@link runColor} does: a token keeps tracking the destination theme.
 */
function bulletColor(bullet: BulletStyle, notes: NoteScope): string | undefined {
	return colorOption(
		{
			scheme: bullet.colorRef.scheme,
			ownHex: bullet.colorRef.srgb,
			resolvedHex: bullet.colorRef.resolved?.effectiveHex ?? null,
		},
		notes,
		'text.bullet.schemeToken',
		'bullet'
	)
}

/**
 * Per-run character formatting, shared by shape text and table-cell text.
 *
 * The three decoration tokens carry their **explicit off** — `u="none"`, `strike="noStrike"`,
 * `cap="none"` — rather than collapsing it into the unstated case. Stating one is a different
 * fact from stating nothing: run properties resolve down the `a:lstStyle` → placeholder →
 * layout → master chain, so a run that would take `u="sng"` from its list style and states
 * `u="none"` is *not* underlined, while the same run with the attribute dropped is. Only an
 * absent attribute (`null` here) means "state nothing"; each off token is a member of its
 * enumeration in its own right (ECMA-376 §20.1.10.81 / §20.1.10.78) and would be redundant
 * with omission otherwise.
 */
export function runOptions(run: Run, ctx: MapContext): Record<string, IrValue> | undefined {
	const { notes } = ctx
	const underline = run.underline
	const hyperlink = hyperlinkOption(run, ctx)
	// PowerPoint paints a link underlined when the run states no `u`
	// (`test/read/fixtures/slide-jump-link.pptx` renders both its linked runs that way), and the
	// write path states `u="sng"` on every link that sets no underline. The two look the same, but
	// the rebuilt run no longer leaves its underline to the link, and there is no option to say so.
	if (hyperlink !== undefined && underline === null) {
		notes.note(
			'text.hyperlink.underline',
			'flattened',
			'unwritable',
			'this run links somewhere and states no underline, which PowerPoint paints underlined all the same; the write path writes u="sng" on a link that sets none, so the text looks the same but states an underline it used to leave to the link'
		)
	}

	return compact({
		bold: orUndefined(run.bold),
		italic: orUndefined(run.italic),
		underline: underline === null ? undefined : { style: underline },
		strike: orUndefined(run.strike),
		fontSize: orUndefined(run.fontSizePt),
		fontFace: orUndefined(run.fontName),
		color: runColor(run, notes),
		highlight: run.highlight ? literalColor(run.highlight.effectiveHex) : undefined,
		caps: orUndefined(run.caps),
		// The write API spells baseline shift as a percentage, the same unit the read
		// model reports, so superscript/subscript survive without a preset round-trip.
		baseline: orUndefined(run.baselinePct),
		hyperlink,
	})
}

/**
 * A run's colour, preferring the raw `schemeClr` token over the resolved literal so the
 * text keeps tracking the destination theme. Only the ten tokens the write path's `clrMap`
 * covers survive as tokens; the other seven degrade to a literal hex there anyway, so
 * passing one through unchanged would produce a silently different colour.
 */
function runColor(run: Run, notes: NoteScope): string | undefined {
	const scheme = run.schemeColor
	if (scheme !== null)
		return schemeColorOption(scheme, run.resolvedColor?.effectiveHex ?? null, notes, 'text.color.schemeToken', 'run')
	if (run.color !== null) return literalColor(run.color)

	// A run with no colour of its own inherits one, and leaving the option out does *not* pass
	// that inheritance along: `addText` fills an uncoloured non-placeholder run with
	// DEF_FONT_COLOR, so omitting it repaints the text black. That makes this a loss either
	// way, and which loss depends on whether the inherited colour can be resolved — so the two
	// outcomes are declared separately rather than under one note that would overstate one and
	// understate the other.
	const inherited = run.resolvedColor
	if (!inherited) {
		notes.note(
			'text.color.default',
			'approximated',
			'unread',
			'this run inherits its colour and nothing resolves what it inherits (a list style the read model does not walk), so the write path paints it black — the one case where the output colour is not merely frozen but possibly wrong'
		)
		return undefined
	}
	notes.note(
		'text.color.inherited',
		'flattened',
		'unsupported',
		'this run inherits its colour from the placeholder, master or theme; the write path paints an uncoloured run black instead of leaving it to inherit, so the inherited colour is resolved and baked in and no longer tracks a theme change'
	)
	return literalColor(inherited.effectiveHex)
}

/** The `@action` of a jump to another slide of the deck. */
const SLIDE_JUMP = 'ppaction://hlinksldjump'

/**
 * A run's hyperlink: a URL, or a jump to another slide of this deck by its number.
 *
 * The read model reports a slide jump as the target's part name, and the write API takes a slide
 * number, which {@link MapContext.slideNumberOf} resolves. Anything else a run can link to has no
 * run-level spelling (a show jump such as "next slide", which the write side emits at shape level
 * only; a custom show; another file; a slide jump from a layout's shape) and is noted. A slide
 * jump used to be dropped with a comment saying the deck-level walk handled it, which it did not,
 * so the text kept its link formatting and went nowhere.
 */
function hyperlinkOption(run: Run, ctx: MapContext): IrValue | undefined {
	return linkOption(run.hyperlink, ctx, RUN_LINK)
}

/** The show-jump actions `HyperlinkProps.action` spells, by the `@action` each one is written as. */
const SHOW_JUMP_ACTIONS: ReadonlyMap<string, string> = new Map(
	['firstslide', 'previousslide', 'nextslide', 'lastslide', 'lastslideviewed', 'endshow'].map((jump) => [
		`ppaction://hlinkshowjump?jump=${jump}`,
		jump,
	])
)

/** What {@link linkOption} can spell where the link hangs, and how to describe a loss there. */
interface LinkSite {
	/** The fidelity note's construct key. */
	construct: 'text.hyperlink' | 'shape.hyperlink'
	/** Whether this site's option takes `action`; only a shape's does. */
	takesAction: boolean
	/** Names the thing that carries the link, opening the note. */
	subject: string
	/** What the option can spell, for the note. */
	spells: string
	/** What is lost, closing the note. */
	loss: string
}

const RUN_LINK: LinkSite = {
	construct: 'text.hyperlink',
	takesAction: false,
	subject: "this run's hyperlink",
	spells:
		"all a run's hyperlink option spells (a show jump, a custom show, another file, or a slide jump from a layout's shape)",
	loss: 'the text keeps its formatting and links nowhere',
}

const SHAPE_LINK: LinkSite = {
	construct: 'shape.hyperlink',
	takesAction: false,
	subject: "this shape's hyperlink",
	spells: "all a shape's hyperlink option spells (a custom show, another file, or a slide jump from a layout's shape)",
	loss: 'the shape is emitted without it and clicking it does nothing',
}

/**
 * The `@action` PowerPoint puts on an audio or video picture's `p:cNvPr`. It is structural rather
 * than a link anyone set -- the media emitter writes it itself (`gen/slide/objects/media.ts`), and
 * `addMedia` writes it again -- so it is not a hyperlink to carry and not a loss to note.
 */
const MEDIA_ACTION = 'ppaction://media'

/** A shape's own click hyperlink, as `addShape`/`addText`/`addImage` take it. */
export function shapeHyperlinkOption(shape: AnyShape, ctx: MapContext): IrValue | undefined {
	const link = shape.hyperlink
	if (link?.action === MEDIA_ACTION) return undefined
	return linkOption(link, ctx, { ...SHAPE_LINK, takesAction: true })
}

/**
 * One reading of a click hyperlink for both sites that carry one. A run and a shape take the same
 * `HyperlinkProps` apart from `action`, which only a shape's is written as, so the difference is a
 * flag rather than a second copy of the mapping.
 */
function linkOption(link: Hyperlink | null, ctx: MapContext, site: LinkSite): IrValue | undefined {
	if (!link) return undefined
	const tooltip = link.tooltip ?? undefined
	if (link.url) return compact({ url: link.url, tooltip })
	const slide = link.action === SLIDE_JUMP && link.targetPartName ? ctx.slideNumberOf(link.targetPartName) : null
	if (slide !== null) return compact({ slide, tooltip })
	const jump = link.action === null ? undefined : SHOW_JUMP_ACTIONS.get(link.action)
	if (site.takesAction && jump !== undefined) return compact({ action: jump, tooltip })
	ctx.notes.note(
		site.construct,
		'dropped',
		'unwritable',
		`${site.subject} (${link.action ?? 'an internal target'}) is neither a URL nor a jump to another slide of the deck, which is ${site.spells}, so ${site.loss}`
	)
	return undefined
}

/** The constructs and wording {@link noteUnreadText} records a shape's or a table cell's losses in. */
const UNREAD_TEXT = {
	shape: {
		field: 'text.field',
		equation: 'text.equation',
		subject: 'this shape',
		call: 'addText',
		equationLoss:
			'the shape is emitted without it - even though TextProps.math (and the ts-pptx/math subpath) could author one',
	},
	cell: {
		field: 'table.cell.field',
		equation: 'table.cell.equation',
		subject: 'this table cell',
		call: 'addTable',
		equationLoss: 'the cell is emitted without it',
	},
} as const

/**
 * Note the text a frame holds that no run carries, and return the text its runs do.
 *
 * Authorable text is *runs*. `TextFrame.text` also counts `a:fld` field text (a slide number, date
 * or footer placeholder), which has no run behind it, and an OMML equation contributes to neither.
 * Both were checked for an auto shape only, so the same field in a table cell was lost with no note.
 * @param frame - the text frame
 * @param element - the element holding it, searched for an equation
 * @param notes - the scope the losses are recorded on
 * @param holder - whether the frame belongs to a shape or a table cell
 * @returns the frame's run text
 */
export function noteUnreadText(
	frame: TextFrame,
	element: unknown,
	notes: NoteScope,
	holder: keyof typeof UNREAD_TEXT
): string {
	const wording = UNREAD_TEXT[holder]
	const runText = frame.paragraphs.flatMap((p) => p.runs.map((run) => run.text)).join('')

	// Run text is a subsequence of the frame's text, so anything longer came from a field.
	if (frame.text.replaceAll('\n', '').length > runText.length) {
		notes.note(
			wording.field,
			'dropped',
			'unread',
			`${wording.subject} holds an automatic field (a:fld - a slide number, date or footer), which has no accessor and no ${wording.call} expression, so its text is not reproduced`
		)
	}

	// An OMML equation contributes nothing to `TextFrame.text`, so an equation-only shape looks like
	// an empty box and would otherwise emit as bare geometry.
	if (hasEquation(element)) {
		notes.note(
			wording.equation,
			'dropped',
			'unread',
			`${wording.subject} holds an OMML equation, which no accessor exposes, so ${wording.equationLoss}`
		)
	}
	return runText
}

/** Paragraph-level properties, replicated onto each of the paragraph's runs. */
function paragraphOptions(paragraph: Paragraph, notes: NoteScope): Record<string, IrValue> {
	const spacing = paragraph.lineSpacing
	const bullet = paragraph.bulletDetail
	const align = paragraph.align

	if (align !== null && !(align in ALIGN)) {
		notes.note(
			'text.align',
			'dropped',
			'unwritable',
			`distributed alignment "${align}" has no write-API spelling, so the paragraph falls back to its inherited alignment`
		)
	}
	if (paragraph.spaceBeforePt === 0 || paragraph.spaceAfterPt === 0) {
		notes.note(
			'text.paraSpaceZero',
			'dropped',
			'unwritable',
			'this paragraph explicitly sets zero space before or after (a:spcBef / a:spcAft of 0), which suppresses the spacing its list style would otherwise apply; the write path treats 0 as "unset" and emits nothing, so the inherited spacing comes back'
		)
	}
	// `text.indent` used to be filed here — a paragraph's own `a:pPr/@marL` and `@indent` had no
	// write option at all, so every margin it stated was replaced by whatever the `bullet` state
	// wrote. `paraMarginLeft` / `paraIndent` state them now, so the mapper carries the values
	// instead of excusing their loss and the note is gone rather than unmapped.
	const bulletValue = bullet === null ? 'inherit' : bulletOption(bullet, notes)

	return (
		compact({
			align: align === null ? undefined : ALIGN[align],
			indentLevel: paragraph.level === 0 ? undefined : paragraph.level,
			paraMarginLeft: paraMarginOption(paragraph.marginLeftPt, bulletValue),
			paraIndent: paraMarginOption(paragraph.indentPt, bulletValue),
			paraSpaceBefore: orUndefined(paragraph.spaceBeforePt),
			paraSpaceAfter: orUndefined(paragraph.spaceAfterPt),
			lineSpacing: spacing?.type === 'points' ? spacing.valuePt : undefined,
			lineSpacingMultiple: spacing?.type === 'percent' ? spacing.percent / 100 : undefined,
			// A paragraph with no bullet child of its own states nothing, and `'inherit'` is
			// how the write path says that. `undefined` would NOT do — an omitted `bullet`
			// emits an explicit `a:buNone`, which suppresses whatever the destination list
			// style has. That was this converter's most frequent loss (`text.bullet.inherited`,
			// 34 of 44 corpus fixtures) until the option gained a third state. Resolved above
			// rather than here: the margin keys need to know which bullet state was chosen, and
			// calling `bulletOption` twice would file its notes twice.
			bullet: bulletValue,
		}) ?? {}
	)
}

/**
 * Flatten a text frame into the `TextProps[]` first argument of `addText`.
 *
 * An empty paragraph becomes an item with empty text and a `breakLine`, which is how a
 * blank line survives — dropping it would silently close up vertical space the author put
 * there deliberately.
 */
export function textRuns(frame: TextFrame, ctx: MapContext): IrValue[] {
	const paragraphs = frame.paragraphs
	const items: IrValue[] = []

	paragraphs.forEach((paragraph, paragraphIndex) => {
		const paraOpts = paragraphOptions(paragraph, ctx.notes)
		const runs = paragraph.runs
		// The paragraph break rides on the last run of every paragraph but the final one;
		// a trailing break would add an empty line the frame never had.
		const breaks = paragraphIndex < paragraphs.length - 1

		if (runs.length === 0) {
			items.push({ text: '', options: compact({ ...paraOpts, breakLine: breaks }) ?? {} })
			return
		}

		runs.forEach((run, runIndex) => {
			const isLastRun = runIndex === runs.length - 1
			// `bullet` goes on the first run only, unlike the other paragraph properties. The
			// write path treats a bullet on a run that is *not* starting a line as a request for
			// a new paragraph — and clears that run's `breakLine` while it is at it — so
			// replicating the bullet split every bulleted paragraph that had more than one run.
			// It reads paragraph properties off whichever run opens the line, so once is enough.
			const continuation = { ...paraOpts }
			delete continuation['bullet']
			const options = compact({
				...(runIndex === 0 ? paraOpts : continuation),
				...runOptions(run, ctx),
				...(isLastRun && breaks ? { breakLine: true } : {}),
			})
			items.push(compact({ text: run.text, options }) ?? { text: run.text })
		})
	})

	return items
}

/**
 * Frame-level layout for the `TextPropsOptions` second argument — everything that belongs
 * to the body as a whole rather than to any paragraph.
 *
 * `resolvedAnchor` is the one place a *resolved* read is the right one: vertical anchoring
 * has no per-run fallback, so an unset `a:bodyPr/@anchor` genuinely means "whatever the
 * placeholder says", and an appended shape does not inherit that.
 */
export function textFrameOptions(frame: TextFrame, notes: NoteScope): Record<string, IrValue> {
	const body = frame.bodyProperties
	const anchor = frame.resolvedAnchor
	const vert = body?.vert ?? null

	if (vert !== null && !WRITABLE_VERT.has(vert)) {
		notes.note(
			'text.vert',
			'dropped',
			'unwritable',
			`text direction "${vert}" is outside TextVertType, so the body falls back to horizontal`
		)
	}

	return (
		compact({
			// Spelled out even when nothing in the source set it. `resolvedAnchor` returning
			// null means PowerPoint would default the body to top — but `addText` defaults a
			// non-placeholder body to *centre*, so leaving the option out re-anchors the text.
			valign: anchor === null ? 'top' : ANCHOR_TO_VALIGN[anchor],
			// `@wrap` is `square`/`none` in OOXML and a boolean in the write API.
			wrap: body?.wrap === null || body?.wrap === undefined ? undefined : body.wrap !== 'none',
			vert: vert !== null && WRITABLE_VERT.has(vert) ? vert : undefined,
			fit: fitOption(frame, notes),
			margin: marginOption(body),
		}) ?? {}
	)
}

/**
 * `a:bodyPr`'s autofit as the write API's `fit`.
 *
 * A `normAutofit` that bakes `fontScale` or `lnSpcReduction` is a *different state* from a
 * bare `<a:normAutofit/>`, so the two get different spellings. ECMA-376 §21.1.2.1.3 defaults
 * each attribute to 100%/0% only when it is **omitted**; PowerPoint recomputes an unbaked
 * scale on edit and draws the baked one exactly as written until then. Flattening both to
 * `'shrink'` therefore re-emits a frame that paints its text at full size — a deck baked at
 * `fontScale="40000"` comes back two and a half times too large — and the object form is
 * already what the write API accepts, so nothing has to be lost here.
 *
 * Out-of-range percentages are the one arm that still loses something. The write path rejects
 * anything outside 0–100 and drops the attribute with a warning, so passing one through would
 * turn a declared loss into a silent one; the frame falls back to bare `'shrink'` with a note.
 */
function fitOption(frame: TextFrame, notes: NoteScope): IrValue | undefined {
	const autofit = frame.autofit
	if (autofit === null) return undefined
	if (autofit !== 'normAutofit') return FIT[autofit]

	const baked = compact({
		fontScale: bakedPct(frame.autofitFontScale, 'fontScale', notes),
		lnSpcReduction: bakedPct(frame.autofitLineSpaceReduction, 'lnSpcReduction', notes),
	})
	return baked ? { type: 'shrink', ...baked } : FIT[autofit]
}

/** A baked `a:normAutofit` percentage, or `undefined` (with a note) when the write API would reject it. */
function bakedPct(value: number | null, name: 'fontScale' | 'lnSpcReduction', notes: NoteScope): number | undefined {
	if (value === null) return undefined
	if (value >= 0 && value <= 100) return value
	notes.note(
		`text.autofit.${name}`,
		'dropped',
		'unwritable',
		`baked a:normAutofit/@${name} of ${value}% is outside the 0-100 the write API accepts, so the frame re-emits a bare <a:normAutofit/>`
	)
	return undefined
}

/**
 * `a:bodyPr` insets as the write API's `margin`.
 *
 * A unit change, not just a copy: the read model reports insets in points, `margin` takes
 * **inches**, and it warns on any value `>= 1` on the assumption it is a stray points
 * value. Passing points straight through would therefore both inset the text by roughly
 * 72× too much and trip that warning.
 *
 * Emitted as the four-tuple whenever any inset is set, since a partial tuple is not
 * expressible; an unset side takes PowerPoint's own default rather than zero, so those are
 * spelled out — defaulting them to 0 would visibly reflow every body that sets one side.
 */

function marginOption(body: BodyProperties | null): IrValue | undefined {
	const insets = body?.insetsPt
	if (!insets) return undefined
	const { left, right, top, bottom } = insets
	if (left === undefined && right === undefined && top === undefined && bottom === undefined) return undefined
	return [
		pointsToInches(top ?? BODY_INSET_DEFAULTS_PT.top),
		pointsToInches(right ?? BODY_INSET_DEFAULTS_PT.right),
		pointsToInches(bottom ?? BODY_INSET_DEFAULTS_PT.bottom),
		pointsToInches(left ?? BODY_INSET_DEFAULTS_PT.left),
	]
}
