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
import type { NoteScope } from '../fidelity.js'
import type { IrValue } from '../ir.js'
import { compact, isWritableSchemeToken, literalColor, orUndefined, pointsToInches } from './values.js'

/**
 * `a:pPr/@algn` → `HAlign`. `dist` and `thaiDist` (distributed justification) are
 * deliberately absent: the write API has no spelling for them, and mapping them onto
 * `justify` would be a silent visual change rather than a declared loss.
 */
const ALIGN: Record<string, string> = { l: 'left', ctr: 'center', r: 'right', just: 'justify' }

/** `a:bodyPr/@anchor` → `VAlign`. */
const ANCHOR: Record<string, string> = { t: 'top', ctr: 'middle', b: 'bottom' }

/** `AutofitMode` → the write API's `fit`. */
const FIT: Record<string, string> = { none: 'none', normAutofit: 'shrink', spAutoFit: 'resize' }

/** `a:bodyPr/@vert` values the write API's `vert` accepts (`TextVertType`). */
const WRITABLE_VERT = new Set(['eaVert', 'horz', 'mongolianVert', 'vert', 'vert270', 'wordArtVert', 'wordArtVertRtl'])

/**
 * The `a:buAutoNum/@type` values the write API's `numberType` names.
 *
 * This used to be the *discriminator* between a numbered and a character bullet, back when
 * the read model reported both as one tagged string. `BulletDetail.kind` does that now, so
 * the set has one job left: telling a scheme the write path can spell from one it cannot.
 */
const AUTO_NUMBER_TYPES = new Set([
	'alphaLcParenBoth',
	'alphaLcParenR',
	'alphaLcPeriod',
	'alphaUcParenBoth',
	'alphaUcParenR',
	'alphaUcPeriod',
	'arabicParenBoth',
	'arabicParenR',
	'arabicPeriod',
	'arabicPlain',
	'romanLcParenBoth',
	'romanLcParenR',
	'romanLcPeriod',
	'romanUcParenBoth',
	'romanUcParenR',
	'romanUcPeriod',
])

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

	// 25–400% is the range the write path accepts; outside it the option is rejected with a
	// console warning and the glyph silently falls back to full size, so it is declared here.
	const sizePct = bullet.sizePct
	if (sizePct !== null && (sizePct < 25 || sizePct > 400)) {
		notes.note(
			'text.bullet.sizePct',
			'approximated',
			'unwritable',
			`bullet glyph size ${sizePct}% is outside the 25–400% range the write API accepts, so no a:buSzPct is emitted and the glyph size is left to be inherited`
		)
	}

	return (
		compact({
			fontFace: orUndefined(bullet.font),
			size: sizePct === null || sizePct < 25 || sizePct > 400 ? undefined : sizePct,
			color: bulletColor(bullet, notes),
		}) ?? {}
	)
}

/**
 * The bullet's colour, preferring the raw `schemeClr` token over the resolved literal for the
 * same reason {@link runColor} does: a token keeps tracking the destination theme.
 */
function bulletColor(bullet: BulletStyle, notes: NoteScope): string | undefined {
	const scheme = bullet.schemeColor
	if (isWritableSchemeToken(scheme)) return scheme as string
	if (bullet.color !== null) return literalColor(bullet.color)
	const resolved = bullet.resolvedColor
	if (!resolved) return undefined
	if (scheme !== null) {
		notes.note(
			'text.bullet.schemeToken',
			'approximated',
			'unwritable',
			`bullet scheme colour "${scheme}" is outside the ten tokens the write path maps, so it is baked to a literal hex and stops tracking the theme`
		)
	}
	return literalColor(resolved.effectiveHex)
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
export function runOptions(run: Run, notes: NoteScope): Record<string, IrValue> | undefined {
	const underline = run.underline

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
		hyperlink: hyperlinkOption(run),
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
	if (isWritableSchemeToken(scheme)) return scheme as string
	if (scheme !== null) {
		notes.note(
			'text.color.schemeToken',
			'approximated',
			'unwritable',
			`scheme colour "${scheme}" is outside the ten tokens the write path maps, so it is baked to a literal hex and stops tracking the theme`
		)
		const resolved = run.resolvedColor
		return resolved ? literalColor(resolved.effectiveHex) : undefined
	}
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
			'this run inherits its colour and nothing resolves what it inherits (a table style tier, or a list style the read model does not walk), so the write path paints it black — the one case where the output colour is not merely frozen but possibly wrong'
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

/**
 * A run's hyperlink. Only external links map: the read model reports an internal jump as a
 * target part name, while the write API takes a slide *number*, which the deck-level walk
 * resolves — so a slide link is handled there, not here.
 */
function hyperlinkOption(run: Run): IrValue | undefined {
	const link = run.hyperlink
	if (!link?.url) return undefined
	return compact({ url: link.url, tooltip: link.tooltip ?? undefined })
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
	if (paragraph.marginLeftPt !== null || paragraph.indentPt !== null) {
		notes.note(
			'text.indent',
			'dropped',
			'unwritable',
			"a paragraph's own indent (a:pPr/@marL, @indent) is not expressible; only the discrete indentLevel is, so hanging indents flatten to the level default"
		)
	}

	return (
		compact({
			align: align === null ? undefined : ALIGN[align],
			indentLevel: paragraph.level === 0 ? undefined : paragraph.level,
			paraSpaceBefore: orUndefined(paragraph.spaceBeforePt),
			paraSpaceAfter: orUndefined(paragraph.spaceAfterPt),
			lineSpacing: spacing?.type === 'points' ? spacing.valuePt : undefined,
			lineSpacingMultiple: spacing?.type === 'percent' ? spacing.percent / 100 : undefined,
			bullet: bullet === null ? undefined : bulletOption(bullet, notes),
		}) ?? {}
	)
}

/**
 * Record the one bullet loss that has no expression at all, once per shape.
 *
 * A paragraph with no bullet child of its own inherits whatever the layout's or master's
 * list style says. There is no way to say "inherit" through the write API — an omitted
 * `bullet` makes it emit an explicit `a:buNone` — so an inherited bullet is suppressed. The
 * read side cannot see the inherited value either (`a:lvl1pPr` list styles are unread), so
 * the converter can neither reproduce it nor tell whether there was one, which is why this
 * is stated as a possibility rather than a fact.
 */
function noteInheritedBullets(paragraphs: readonly Paragraph[], notes: NoteScope): void {
	if (!paragraphs.some((paragraph) => paragraph.bulletDetail === null)) return
	notes.note(
		'text.bullet.inherited',
		'dropped',
		'unread',
		'at least one paragraph sets no bullet of its own and inherits one from the layout or master list style; the write path cannot express "inherit" and emits an explicit a:buNone, so any inherited bullet is suppressed — a no-op where the inherited style had none, and a visible change where it did not'
	)
}

/**
 * Flatten a text frame into the `TextProps[]` first argument of `addText`.
 *
 * An empty paragraph becomes an item with empty text and a `breakLine`, which is how a
 * blank line survives — dropping it would silently close up vertical space the author put
 * there deliberately.
 */
export function textRuns(frame: TextFrame, notes: NoteScope): IrValue[] {
	const paragraphs = frame.paragraphs
	const items: IrValue[] = []
	noteInheritedBullets(paragraphs, notes)

	paragraphs.forEach((paragraph, paragraphIndex) => {
		const paraOpts = paragraphOptions(paragraph, notes)
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
				...runOptions(run, notes),
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
			valign: anchor === null ? 'top' : ANCHOR[anchor],
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
function bakedPct(value: number | null, name: string, notes: NoteScope): number | undefined {
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
const DEFAULT_INSETS_PT = { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 }

function marginOption(body: BodyProperties | null): IrValue | undefined {
	const insets = body?.insetsPt
	if (!insets) return undefined
	const { left, right, top, bottom } = insets
	if (left === undefined && right === undefined && top === undefined && bottom === undefined) return undefined
	return [
		pointsToInches(top ?? DEFAULT_INSETS_PT.top),
		pointsToInches(right ?? DEFAULT_INSETS_PT.right),
		pointsToInches(bottom ?? DEFAULT_INSETS_PT.bottom),
		pointsToInches(left ?? DEFAULT_INSETS_PT.left),
	]
}
