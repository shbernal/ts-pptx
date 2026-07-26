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
 * Run-level reads prefer the *explicit* value over the resolved one throughout. A resolved
 * value folds in the placeholder/master/theme inheritance chain, so writing it back would
 * bake inherited styling into the shape as if it had been authored there — the deck would
 * look right today and stop tracking its own theme.
 */
import type { BodyProperties, Paragraph, Run, TextFrame } from '../../read/api/text.js'
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
 * `a:buAutoNum/@type` values, used to tell a numbered bullet from a character bullet.
 * The read model reports both as a single string — the character for `a:buChar`, the
 * `@type` token for `a:buAutoNum` — so membership here is the only discriminator.
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

function bulletOption(bullet: string, notes: NoteScope): IrValue {
	notes.note(
		'text.bullet.style',
		'flattened',
		'unread',
		"a bullet's own font, size and colour have no accessor (a:buFont / a:buSzPct / a:buClr), so it renders in the body font and colour"
	)
	if (AUTO_NUMBER_TYPES.has(bullet)) {
		notes.note(
			'text.bullet.numberStartAt',
			'dropped',
			'unread',
			'a numbered bullet restarts at 1: nothing reads a:buAutoNum/@startAt, though addText accepts numberStartAt'
		)
		return { type: 'number', style: bullet }
	}
	return { characterCode: (bullet.codePointAt(0) ?? 0x2022).toString(16).toUpperCase() }
}

/** Per-run character formatting, shared by shape text and table-cell text. */
export function runOptions(run: Run, notes: NoteScope): Record<string, IrValue> | undefined {
	const underline = run.underline
	const strike = run.strike
	const caps = run.caps

	return compact({
		bold: orUndefined(run.bold),
		italic: orUndefined(run.italic),
		underline: underline === null || underline === 'none' ? undefined : { style: underline },
		strike: strike === null || strike === 'noStrike' ? undefined : strike,
		fontSize: orUndefined(run.fontSizePt),
		fontFace: orUndefined(run.fontName),
		color: runColor(run, notes),
		highlight: run.highlight ? literalColor(run.highlight.effectiveHex) : undefined,
		caps: caps === null || caps === 'none' ? undefined : caps,
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
	return run.color === null ? undefined : literalColor(run.color)
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
	const bullet = paragraph.bullet
	const align = paragraph.align

	if (align !== null && !(align in ALIGN)) {
		notes.note(
			'text.align',
			'dropped',
			'unwritable',
			`distributed alignment "${align}" has no write-API spelling, so the paragraph falls back to its inherited alignment`
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
 * Flatten a text frame into the `TextProps[]` first argument of `addText`.
 *
 * An empty paragraph becomes an item with empty text and a `breakLine`, which is how a
 * blank line survives — dropping it would silently close up vertical space the author put
 * there deliberately.
 */
export function textRuns(frame: TextFrame, notes: NoteScope): IrValue[] {
	const paragraphs = frame.paragraphs
	const items: IrValue[] = []

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
			const options = compact({
				...paraOpts,
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
	const autofit = frame.autofit
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
			valign: anchor === null ? undefined : ANCHOR[anchor],
			// `@wrap` is `square`/`none` in OOXML and a boolean in the write API.
			wrap: body?.wrap === null || body?.wrap === undefined ? undefined : body.wrap !== 'none',
			vert: vert !== null && WRITABLE_VERT.has(vert) ? vert : undefined,
			fit: autofit === null ? undefined : FIT[autofit],
			margin: marginOption(body),
		}) ?? {}
	)
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
