/**
 * ts-pptx: DrawingML text runs & paragraphs
 *
 * The run/paragraph half of text-body generation: paragraph properties
 * (`<a:pPr>`, bullets, spacing), run properties (`<a:rPr>`), the `<a:r>` run
 * builder, the shrink-autofit helper, and the grouping/rendering of a flat run
 * list into `<a:p>` paragraphs. The `text-body.ts` container layer builds on top.
 */

import { BulletType, SlideObjectType } from '../../enums.js'
import { DEF_BULLET_MARGIN, DEF_TEXT_GLOW, DEF_TEXT_SHADOW } from '../../constants-internal.js'
import type { ObjectOptions, TableCell, TextFitShrinkProps, TextProps, TextPropsOptions } from '../../types/index.js'
import type { SlideObject } from '../../types/internal.js'
import { createColorElement } from './color.js'
import { createGlowElement, createShadowElement } from './effect.js'
import { genXmlColorSelection, solidPaint } from './fill.js'
import { setOrClear } from '../utils.js'
import { inch2Emu, lineWidthToEmu, percentToFixedPercent, ptsToEmuLenient } from '../../units-internal.js'
import { ptToHundredths } from '../../units.js'
import { warn } from '../../diagnostics.js'
import { el, raw, voidEl, type XmlAttrs } from '../oxml/el.js'
import {
	clampCharSpacingSpc,
	clampFontSizeSz,
	clampLineSpacingMultiplePct,
	clampLineSpacingPts,
	clampParaIndentEmu,
	clampParaMarginEmu,
} from './clamp.js'
import { genXmlInlineMath, genXmlMathParagraph } from './math.js'
import { InvalidOptionError } from '../../errors.js'

/** The 2018 hyperlink-color extension namespace, written on the `<ahyp:hlinkClr>` element itself. */
const AHYP_NS = 'http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor'

/**
 * `<a:buChar>` cannot go through the element builder. Its `char` attribute carries a *pre-escaped*
 * numeric character reference (`&#x2022;` — see `BulletType`), and the builder escapes every
 * attribute value, which would emit `&amp;#x2022;` and render that text literally as the bullet.
 * There is no raw-attribute escape hatch, so this one stays a template.
 */
function buChar(char: string): string {
	return `<a:buChar char="${char}"/>`
}

/**
 * `<a:extLst>` marking a hyperlink whose color should follow the text color rather than the
 * theme's hyperlink color. The leading spaces are byte-significant: this block was authored as
 * indented string concatenation and that indentation reaches the file, so it is described here
 * with `openPrefix`/`closePrefix` rather than being tidied away.
 */
const HLINK_TEXT_COLOR_EXT = el(
	'a:extLst',
	null,
	raw(
		el(
			'a:ext',
			{ uri: '{A12FA001-AC4F-418D-AE19-62706E023703}' },
			raw(voidEl('ahyp:hlinkClr', { 'xmlns:ahyp': AHYP_NS, val: 'tx' }, { openPrefix: '   ' })),
			{ openPrefix: '  ', closePrefix: '  ' }
		)
	),
	{ openPrefix: ' ', closePrefix: ' ' }
)

/**
 * Whether a `bullet` value makes the paragraph DRAW a bullet (a glyph, a number or a picture).
 *
 * Two call sites downstream of the emitter act on "this paragraph has a bullet" rather than on
 * "this paragraph says something about its bullet", and both were plain truthiness tests until
 * `'inherit'` — a truthy string that draws nothing — joined the union. Left as truthiness they
 * would have broken the line grouping (every run starting its own paragraph) and eaten a real
 * leading `•` from text that emits no `a:buChar` to duplicate it.
 * @param {TextPropsOptions['bullet']} bullet - the option value
 * @return {boolean} true when bullet markup is emitted
 */
function emitsBulletMarkup(bullet: TextPropsOptions['bullet']): boolean {
	return !!bullet && bullet !== 'inherit'
}

/**
 * One paragraph margin attribute's value in EMU, or `null` when the attribute is not written.
 *
 * Three inputs, because the option has three states and so does the attribute: an explicit
 * number is clamped and written, `'inherit'` writes nothing, and an absent option falls back to
 * whatever the `bullet` arm decided — which is also `null` for `bullet: 'inherit'`, the one arm
 * that states nothing on its own.
 * @param {number|'inherit'|undefined} option - `paraMarginLeft` / `paraIndent` as given
 * @param {number|null} fallbackEmu - the bullet arm's default, `null` for no attribute
 * @param {Function} clamp - the attribute's ST_* range clamp (points in, EMU out)
 * @return {number|null} the EMU value to write, or null to omit the attribute
 */
function resolveParagraphMargin(
	option: number | 'inherit' | undefined,
	fallbackEmu: number | null,
	clamp: (points: number) => number
): number | null {
	if (option === undefined) return fallbackEmu
	if (option === 'inherit') return null
	return clamp(option)
}

/**
 * Generate XML Paragraph Properties
 * @param {SlideObject|TextProps} textObj - text object
 * @param {boolean} isDefault - array of default relations
 * @return {string} XML
 */
export function genXmlParagraphProperties(textObj: SlideObject | TextProps, isDefault: boolean): string {
	// `options` is always present on text objects reaching here; narrow it once (both union
	// members have all-optional props, so an empty object is a valid fallback).
	const opts: NonNullable<typeof textObj.options> = textObj.options ?? {}
	let strXmlBullet = ''
	let strXmlBulletColor = ''
	let strXmlLnSpc = ''
	let strXmlParaSpc = ''
	let strXmlTabStops = ''
	const tag = isDefault ? 'a:lvl1pPr' : 'a:pPr'
	let bulletMarL = ptsToEmuLenient(DEF_BULLET_MARGIN)
	// The paragraph's own margins (`a:pPr/@marL` and `@indent`, in EMU). Each `bullet` arm below
	// records the default it has always written and `paraMarginLeft`/`paraIndent` override it;
	// `null` means the attribute is not written at all, which is what leaves the margin to the
	// `a:lstStyle` -> placeholder -> layout -> master chain. They are resolved and appended once,
	// after the chain, rather than by each arm — the arms decide a default, not the output.
	let defaultMarL: number | null = null
	let defaultIndent: number | null = null
	// The no-bullet arm writes `indent` before `marL`; every other arm writes `marL` first.
	// Attribute order carries no meaning to a reader, but it is byte-significant and the demo
	// decks are pinned byte-for-byte, so each arm keeps the order it has always emitted.
	let indentBeforeMarL = false

	// NOTE: this open tag is deliberately NOT built with `openTag`/`el`. When `rtlMode` is set the
	// historical template emits `rtl="1" ` with a TRAILING space while every attribute appended
	// below contributes a LEADING one, so `rtl` + `algn` produces a DOUBLE space between them.
	// `openTag` joins attributes with exactly one space, so that layout is not expressible — and
	// the demo deck contains no RTL text (zero `rtl="1"` parts), so the byte gate could not catch
	// the change. Left as-is; the children below are built with the element builder.
	let paragraphPropXml = `<${tag}${opts.rtlMode ? ' rtl="1" ' : ''}`

	// A: Build paragraphProperties
	{
		// OPTION: align
		if (opts.align) {
			switch (opts.align) {
				case 'left':
					paragraphPropXml += ' algn="l"'
					break
				case 'right':
					paragraphPropXml += ' algn="r"'
					break
				case 'center':
					paragraphPropXml += ' algn="ctr"'
					break
				case 'justify':
					paragraphPropXml += ' algn="just"'
					break
				default:
					paragraphPropXml += ''
					break
			}
		}

		if (opts.lineSpacing) {
			strXmlLnSpc = el('a:lnSpc', null, raw(voidEl('a:spcPts', { val: clampLineSpacingPts(opts.lineSpacing) })))
		} else if (opts.lineSpacingMultiple) {
			const val = clampLineSpacingMultiplePct(opts.lineSpacingMultiple)
			strXmlLnSpc = el('a:lnSpc', null, raw(voidEl('a:spcPct', { val })))
		}

		// OPTION: indent
		// `a:p/@lvl` is ST_TextIndentLevelType (0-8) and the value is written straight into the
		// attribute, so an unusable one is a repair prompt rather than a wrong-looking slide.
		// The guard used to be truthiness plus `> 0`, which `Infinity` passes: `lvl="Infinity"`.
		if (opts.indentLevel) {
			if (Number.isInteger(opts.indentLevel) && opts.indentLevel > 0 && opts.indentLevel <= 8)
				paragraphPropXml += ` lvl="${opts.indentLevel}"`
			else
				warn(
					'text/invalid-indent-level',
					`indentLevel ${String(opts.indentLevel)} must be a whole number from 0 to 8; ignoring it.`
				)
		}

		// OPTION: Paragraph Spacing: Before/After
		// `NaN` is falsy, so truthiness plus `> 0` is the whole guard. A non-finite value is
		// deliberately left to reach `ptToHundredths`, which refuses it — dropping the attribute
		// here instead would silently lose a spacing the caller asked for.
		if (opts.paraSpaceBefore && opts.paraSpaceBefore > 0) {
			strXmlParaSpc += el('a:spcBef', null, raw(voidEl('a:spcPts', { val: ptToHundredths(opts.paraSpaceBefore) })))
		}
		if (opts.paraSpaceAfter && opts.paraSpaceAfter > 0) {
			strXmlParaSpc += el('a:spcAft', null, raw(voidEl('a:spcPts', { val: ptToHundredths(opts.paraSpaceAfter) })))
		}

		// OPTION: bullet
		// NOTE: OOXML uses the unicode character set for Bullets
		// EX: Unicode Character 'BULLET' (U+2022) ==> '<a:buChar char="&#x2022;"/>'
		if (typeof opts.bullet === 'object') {
			const bulletImage = opts.bullet.image
			const isPictureBullet = !!(bulletImage && (bulletImage.path || bulletImage.data))
			if (opts.bullet?.indent) bulletMarL = ptsToEmuLenient(opts.bullet.indent)
			// Every bullet form hangs the first line by the same margin, whichever glyph it draws.
			defaultMarL = opts.indentLevel && opts.indentLevel > 0 ? bulletMarL + bulletMarL * opts.indentLevel : bulletMarL
			defaultIndent = -bulletMarL
			// `buClr` colors a glyph/number; it has no effect on a picture bullet, so skip it for `buBlip`.
			if (opts.bullet.color && !isPictureBullet)
				strXmlBulletColor = el('a:buClr', null, raw(createColorElement(opts.bullet.color)))

			// `<a:buSzPct/>` val is thousandths of a percent; ST_TextBulletSizePercent allows 25%-400%.
			//
			// Emitted ONLY when the caller asked for a size. It used to be unconditional, pinned
			// to 100% when unset, and that is not the same thing as leaving it out: an explicit
			// `<a:buSzPct val="100000"/>` *overrides* whatever bullet size the layout's or
			// master's list style sets, so every bullet this path wrote silently forced its glyph
			// back to full size. The same class of bug as the explicit `a:buNone` an omitted
			// `bullet` emits — and invisible until `Paragraph.bulletDetail` gave the round-trip
			// check something to see it with, which reported the added 100% on every bulleted
			// fixture.
			// 25-400% is ST_TextBulletSizePercent's range (and what PowerPoint's bullet-size
			// dialog accepts), so an out-of-range value goes through the same clamp every other
			// percentage option uses: a finite one moves to the nearest bound and warns, and a
			// value that is not a number throws. Dropping the attribute instead — what this did
			// until the policies were unified — resized the glyph to whatever the list style
			// inherits, which is a discarded request reported as a warning.
			const bulletSizePct =
				opts.bullet.size === undefined
					? undefined
					: percentToFixedPercent(opts.bullet.size, 'bullet/size-out-of-range', 'bullet.size', 25, 400)
			const strXmlBulletSize = bulletSizePct === undefined ? '' : voidEl('a:buSzPct', { val: bulletSizePct })
			// NOTE: the builder escapes `typeface`, so the manual `encodeXmlEntities` that used to
			// wrap it here is gone — keeping both would double-escape (`&` -> `&amp;amp;`).
			const strXmlBulletFont = opts.bullet.fontFace ? voidEl('a:buFont', { typeface: opts.bullet.fontFace }) : ''

			if (isPictureBullet) {
				// Picture bullet: <a:buBlip> references a slide media rel registered in addText() (`_rId`).
				// No `buFont` (there is no glyph typeface), but `buSzPct` still scales the image height.
				if (opts.bullet._rId) {
					// SVG bullet: the blip embeds the PNG preview (`_rId`) and references the SVG via the
					// `asvg:svgBlip` extension (`_rIdSvg`), the same dual-rel form addImage() emits for SVG.
					const svgExt = opts.bullet._rIdSvg
						? el(
								'a:extLst',
								null,
								raw(
									el(
										'a:ext',
										{ uri: '{96DAC541-7B7A-43D3-8B79-37D633B846F1}' },
										raw(
											voidEl('asvg:svgBlip', {
												'xmlns:asvg': 'http://schemas.microsoft.com/office/drawing/2016/SVG/main',
												'r:embed': `rId${opts.bullet._rIdSvg}`,
											})
										)
									)
								)
							)
						: ''
					const blip = svgExt
						? el('a:blip', { 'r:embed': `rId${opts.bullet._rId}` }, raw(svgExt))
						: voidEl('a:blip', { 'r:embed': `rId${opts.bullet._rId}` })
					strXmlBullet = strXmlBulletSize + el('a:buBlip', null, raw(blip))
				} else {
					// rel was not registered (eg: bullet on a context without a slide target) - fall back to a glyph
					warn(
						'bullet/image-embed-failed',
						'picture `bullet.image` could not be embedded; using a default bullet glyph'
					)
					strXmlBullet = strXmlBulletSize + strXmlBulletFont + buChar(BulletType.DEFAULT)
				}
			} else if (opts.bullet.type && opts.bullet.type.toString().toLowerCase() === 'number') {
				strXmlBullet =
					strXmlBulletSize +
					(strXmlBulletFont || voidEl('a:buFont', { typeface: '+mj-lt' })) +
					voidEl('a:buAutoNum', {
						type: opts.bullet.numberType || 'arabicPeriod',
						startAt: opts.bullet.numberStartAt || '1',
					})
			} else if (opts.bullet.characterCode) {
				let bulletCode = `&#x${opts.bullet.characterCode};`

				// Check value for hex-ness (s/b 4 char hex)
				if (!/^[0-9A-Fa-f]{4}$/.test(opts.bullet.characterCode)) {
					warn(
						'bullet/invalid-character-code',
						'`bullet.characterCode` should be a 4-digit unicode character (ex: 22AB)!'
					)
					bulletCode = BulletType.DEFAULT
				}

				strXmlBullet = strXmlBulletSize + strXmlBulletFont + buChar(bulletCode)
			} else {
				strXmlBullet = strXmlBulletSize + strXmlBulletFont + buChar(BulletType.DEFAULT)
			}
		} else if (opts.bullet === 'inherit') {
			// The paragraph states NOTHING about its bullet: no `a:buChar`/`a:buAutoNum`, no
			// `a:buNone`, and none of the margin attributes the other arms write. Bullet
			// properties resolve down the `a:lstStyle` -> placeholder -> layout -> master chain,
			// so silence is what lets the list style keep reaching this paragraph — `a:buNone`
			// *overrides* it, which is a different fact even where the inherited style has no
			// bullet, because a later edit to the master then stops arriving.
			//
			// Omission cannot mean this: it has meant `a:buNone` since the writer existed and
			// every deck authored against it depends on that. Same resolution as `fill: { type:
			// 'inherit' }` (#10) — name the state that had no name, leave the one that has.
		} else if (opts.bullet) {
			defaultMarL = opts.indentLevel && opts.indentLevel > 0 ? bulletMarL + bulletMarL * opts.indentLevel : bulletMarL
			defaultIndent = -bulletMarL
			// No `a:buSzPct` here either — `bullet: true` asks for a bullet, not for one pinned
			// to 100% of the body size in defiance of the master's list style. Same reasoning as
			// the `a:buNone` note just below, and it keeps `bullet: true` byte-identical to
			// `bullet: { type: 'bullet' }`, which the object branch above now also leaves out.
			strXmlBullet = buChar(BulletType.DEFAULT)
		} else if (!opts.bullet) {
			// We only add this when the user explicitely asks for no bullet, otherwise, it can override the master defaults!
			// FIX: specify zero indent and marL or default will be hanging paragraph
			defaultMarL = 0
			defaultIndent = 0
			indentBeforeMarL = true
			strXmlBullet = voidEl('a:buNone')
		}

		// OPTION: paraMarginLeft / paraIndent
		// The paragraph's own `@marL`/`@indent`, which until now only the `bullet` arms decided.
		// That conflated two facts on one element: `bullet: false` could not suppress a bullet
		// without ALSO flattening an inherited hanging indent to zero, and no state could set a
		// margin without drawing a bullet. An explicit value wins over the arm's default in every
		// state, and `'inherit'` takes the attribute out entirely — the third state omission
		// cannot spell here either, since omission is what writes the default.
		const marLEmu = resolveParagraphMargin(opts.paraMarginLeft, defaultMarL, clampParaMarginEmu)
		const indentEmu = resolveParagraphMargin(opts.paraIndent, defaultIndent, clampParaIndentEmu)
		const marLAttr = marLEmu === null ? '' : ` marL="${marLEmu}"`
		const indentAttr = indentEmu === null ? '' : ` indent="${indentEmu}"`
		paragraphPropXml += indentBeforeMarL ? indentAttr + marLAttr : marLAttr + indentAttr

		// OPTION: tabStops
		if (opts.tabStops && Array.isArray(opts.tabStops)) {
			const tabStopsXml = opts.tabStops
				.map((stop) => voidEl('a:tab', { pos: inch2Emu(stop.position || 1), algn: stop.alignment || 'l' }))
				.join('')
			strXmlTabStops = el('a:tabLst', null, raw(tabStopsXml))
		}

		// B: Close Paragraph-Properties
		// IMPORTANT: strXmlLnSpc, strXmlParaSpc, and strXmlBullet require strict ordering - anything out of order is ignored. (PPT-Online, PPT for Mac)
		paragraphPropXml += '>' + strXmlLnSpc + strXmlParaSpc + strXmlBulletColor + strXmlBullet + strXmlTabStops
		if (isDefault) paragraphPropXml += genXmlTextRunProperties(opts, true)
		paragraphPropXml += '</' + tag + '>'
	}

	return paragraphPropXml
}

/**
 * Generate XML Text Run Properties (`a:rPr`)
 * @param {ObjectOptions|TextPropsOptions} opts - text options
 * @param {boolean} isDefault - whether these are the default text run properties
 * @return {string} XML
 */
export function genXmlTextRunProperties(opts: ObjectOptions | TextPropsOptions, isDefault: boolean): string {
	let runProps = ''
	const runPropsTag = isDefault ? 'a:defRPr' : 'a:rPr'

	// BEGIN runProperties (ex: `<a:rPr lang="en-US" sz="1600" b="1" dirty="0">`)
	const underline =
		typeof opts.underline === 'object' && opts.underline?.style ? opts.underline.style : opts.hyperlink ? 'sng' : null
	const baseline = opts.baseline
		? Math.round(opts.baseline * 50)
		: opts.subscript
			? -40000
			: opts.superscript
				? 30000
				: null
	// NOTE: attribute ORDER is byte-significant. Listing every attribute in emission order (null =
	// omitted) makes that order reviewable, where the old `+=` chain buried it in control flow.
	const attrs: XmlAttrs = {
		lang: opts.lang ? opts.lang : 'en-US',
		altLang: opts.lang ? 'en-US' : null,
		// NOTE: clamp+round so sizes like '7.5' or out-of-range values wont cause corrupt presentations
		sz: opts.fontSize ? clampFontSizeSz(opts.fontSize) : null,
		// NOTE: `b`/`i` were written as `opts.bold ? '1' : '0'` inside a truthiness guard, so the
		// "0" arm was unreachable — the emitted value is always "1".
		b: opts?.bold ? '1' : null,
		i: opts?.italic ? '1' : null,
		strike: opts?.strike ? (typeof opts.strike === 'string' ? opts.strike : 'sngStrike') : null,
		cap: opts?.caps ? opts.caps : null,
		u: underline,
		baseline,
		spc: opts.charSpacing ? clampCharSpacingSpc(opts.charSpacing) : null,
		kern: opts.charSpacing ? 0 : null, // IMPORTANT: Also disable kerning; otherwise text won't actually expand
		dirty: '0',
	}

	// Color / Font / Highlight / Outline / Effects are children of <a:rPr>, so add them now before closing the runProperties tag
	const hasShadow = !!opts.shadow && opts.shadow.type !== 'none'
	if (
		opts.color ||
		opts.fontFace ||
		opts.outline ||
		opts.glow ||
		hasShadow ||
		(typeof opts.underline === 'object' && opts.underline.color)
	) {
		// NOTE: children must follow CT_TextCharacterProperties order: ln, fill, effectLst, highlight, uFill, latin/ea/cs
		if (opts.outline && typeof opts.outline === 'object') {
			runProps += el(
				'a:ln',
				{ w: lineWidthToEmu(opts.outline.size || 0.75) },
				raw(genXmlColorSelection(opts.outline.color || 'FFFFFF'))
			)
		}
		if (opts.color) runProps += genXmlColorSelection(solidPaint(opts.color, opts.transparency))
		// EFFECTS: glow and shadow share a single <a:effectLst> (only one is allowed per CT_TextCharacterProperties; glow precedes shadow per CT_EffectList)
		if (opts.glow || hasShadow) {
			runProps += el('a:effectLst', null, [
				opts.glow ? raw(createGlowElement(opts.glow, DEF_TEXT_GLOW)) : null,
				hasShadow ? raw(createShadowElement(opts.shadow, DEF_TEXT_SHADOW)) : null,
			])
		}
		if (opts.highlight) runProps += el('a:highlight', null, raw(createColorElement(opts.highlight)))
		if (typeof opts.underline === 'object' && opts.underline.color)
			runProps += el('a:uFill', null, raw(genXmlColorSelection(opts.underline.color)))
		if (opts.fontFace) {
			// Match how PowerPoint writes a font picked from the UI: the chosen typeface goes in the
			// Latin (`<a:latin>`) and complex-script (`<a:cs>`) slots. The East Asian slot (`<a:ea>`) is
			// only written when an EA face is explicitly chosen (`fontFaceEA`); otherwise it inherits the
			// theme. Forcing a Latin-only font into `<a:ea>` — especially with the bogus charset values
			// PowerPoint never emits on ea/cs — duplicates/ghosts text in Office 365.
			// NOTE: order must be latin, ea, cs per CT_TextCharacterProperties.
			// `fontFace`/`fontFaceEA` are caller-supplied and unsanitized upstream, so they are escaped:
			// an unescaped `"` or `&` in a font name closes the attribute early and emits a
			// non-parseable slide part, which PowerPoint reports as a file needing repair. The builder
			// does that escaping now, so the manual `encodeXmlEntities` calls are gone (keeping both
			// would double-escape).
			runProps += voidEl('a:latin', { typeface: opts.fontFace, pitchFamily: '34', charset: '0' })
			if (opts.fontFaceEA) runProps += voidEl('a:ea', { typeface: opts.fontFaceEA })
			runProps += voidEl('a:cs', { typeface: opts.fontFace })
		}
	}

	// Hyperlink support
	if (opts.hyperlink) {
		if (typeof opts.hyperlink !== 'object')
			throw new InvalidOptionError(
				'hyperlink/not-an-object',
				"text `hyperlink` option should be an object. Ex: `hyperlink:{url:'https://github.com'}` "
			)
		else if (!opts.hyperlink.url && !opts.hyperlink.slide && !opts.hyperlink.action)
			throw new InvalidOptionError(
				'hyperlink/missing-target',
				'text `hyperlink` requires either `url`, `slide`, or `action`'
			)
		// An action-only hyperlink (an action-button navigation) lives on the shape's `<p:cNvPr>`
		// (see `cNvPrHyperlink`), NOT on the text run — a labeled action button emits no run-level
		// `<a:hlinkClick>`.
		else if (opts.hyperlink.url || opts.hyperlink.slide) {
			// runProps += '<a:uFill>'+ genXmlColorSelection('0000FF') +'</a:uFill>'; // Breaks PPT2010!
			// NOTE: `tooltip` is escaped by the builder now (the manual `encodeXmlEntities` is gone), and
			// it is written even when absent — an empty `tooltip=""` is part of today's bytes.
			const linkAttrs: XmlAttrs = opts.hyperlink.url
				? {
						'r:id': `rId${opts.hyperlink._rId}`,
						invalidUrl: '',
						action: '',
						tgtFrame: '',
						tooltip: opts.hyperlink.tooltip ?? '',
						history: '1',
						highlightClick: '0',
						endSnd: '0',
					}
				: {
						'r:id': `rId${opts.hyperlink._rId}`,
						action: 'ppaction://hlinksldjump',
						tooltip: opts.hyperlink.tooltip ?? '',
					}
			// An explicit text color means the link must carry the "follow text color" extension, so
			// the element becomes paired; otherwise it self-closes.
			runProps += opts.color
				? el('a:hlinkClick', linkAttrs, raw(HLINK_TEXT_COLOR_EXT))
				: voidEl('a:hlinkClick', linkAttrs)
		}
	}

	// END runProperties
	return el(runPropsTag, attrs, raw(runProps))
}

/**
 * Build textBody text runs [`<a:r></a:r>`] for paragraphs [`<a:p>`]
 * @param {TextProps} textObj - Text object
 * @return {string} XML string
 */
export function genXmlTextRun(textObj: TextProps): string {
	// NOTE: Dont create full rPr runProps for empty [lineBreak] runs
	// Why? The size of the lineBreak wont match (eg: below it will be 18px instead of the correct 36px)
	// Do this:
	/*
		<a:p>
			<a:pPr algn="r"/>
			<a:endParaRPr lang="en-US" sz="3600" dirty="0"/>
		</a:p>
	*/
	// NOT this:
	/*
		<a:p>
			<a:pPr algn="r"/>
			<a:r>
				<a:rPr lang="en-US" sz="3600" dirty="0">
					<a:solidFill>
						<a:schemeClr val="accent5"/>
					</a:solidFill>
					<a:latin typeface="Times" pitchFamily="34" charset="0"/>
					<a:ea typeface="Times" pitchFamily="34" charset="-122"/>
					<a:cs typeface="Times" pitchFamily="34" charset="-120"/>
				</a:rPr>
				<a:t></a:t>
			</a:r>
			<a:endParaRPr lang="en-US" dirty="0"/>
		</a:p>
	*/

	// Return paragraph with text run
	if (textObj.text === undefined || textObj.text === null) return ''
	// `<a:t>` takes a TEXT child, so the builder escapes it — same `encodeXmlEntities` this used to call.
	return el('a:r', null, [
		raw(genXmlTextRunProperties(textObj.options ?? {}, false)),
		raw(el('a:t', null, String(textObj.text))),
	])
}

/**
 * Builds `<a:normAutofit>` with explicit fontScale/lnSpcReduction for "shrink text on overflow"
 * @param {TextFitShrinkProps} fit - shrink fit options
 * @return {string} XML string (`<a:normAutofit .../>`)
 * @see ECMA-376 CT_TextNormAutofit (attributes in 1000ths of a percent)
 */
export function genXmlNormAutofit(fit: TextFitShrinkProps): string {
	// fontScale/lnSpcReduction are authored as a percent (0-100); OOXML stores them in 1000ths
	// of a percent. Both are ST_TextFontScalePercentOrPercentString / ST_TextFontScalePercent
	// ranges the shared percentage clamp already speaks: an out-of-range but finite scale moves
	// to the nearest bound and warns, and a non-number throws. They used to be dropped, which
	// left the shrink autofit silently un-parameterised.
	const pct = (val: number | undefined, name: string): number | null =>
		val === undefined || val === null ? null : percentToFixedPercent(val, 'text/invalid-fit-percentage', `fit.${name}`)

	// `pct` returns null for an absent value, and the builder omits null attributes.
	return voidEl('a:normAutofit', {
		fontScale: pct(fit.fontScale, 'fontScale'),
		lnSpcReduction: pct(fit.lnSpcReduction, 'lnSpcReduction'),
	})
}

// A run of formatted text within a paragraph. Every run reaching STEP 5/6 of genXmlTextBody
// carries an `options` bag (assigned in STEP 4), so model it as required.
export type RunProps = TextProps & { options: TextPropsOptions }

/**
 * Group the flat run list into paragraphs (lines): a new line starts on a display-math run,
 * an alignment change, or a bullet, and closes after a breakLine. Returns the per-line run arrays.
 */
export function groupRunsIntoLines(arrTextObjects: RunProps[], opts: ObjectOptions): RunProps[][] {
	const arrLines: RunProps[][] = []
	let arrTexts: RunProps[] = []
	arrTextObjects.forEach((textObj, idx) => {
		// A0: A DISPLAY math equation is its own paragraph — flush any pending runs and
		// give it its own line so STEP 6 emits the <a14:m><m:oMathPara> wrapper instead of text runs.
		// Inline math (dn-inline-math) flows mid-paragraph, so it falls through to the run-buffering
		// path below and STEP 6 emits a bare <a14:m><m:oMath> run alongside the plain text runs.
		if (textObj.math && !textObj.inline) {
			if (arrTexts.length > 0) {
				arrLines.push(arrTexts)
				arrTexts = []
			}
			arrLines.push([textObj])
			return
		}

		// A: Align or Bullet trigger new line
		if (arrTexts.length > 0 && (textObj.options.align || opts.align)) {
			// Only start a new paragraph when align *changes*
			if (textObj.options.align !== arrTextObjects[idx - 1]?.options.align) {
				arrLines.push(arrTexts)
				arrTexts = []
			}
		} else if (arrTexts.length > 0 && emitsBulletMarkup(textObj.options.bullet)) {
			arrLines.push(arrTexts)
			arrTexts = []
			textObj.options.breakLine = false // For cases with both `bullet` and `brekaLine` - prevent double lineBreak
		}

		// B: Add this text to current line
		arrTexts.push(textObj)

		// C: BreakLine begins new line **after** adding current text
		if (arrTexts.length > 0 && textObj.options.breakLine) {
			// Avoid starting a para right as loop is exhausted
			if (idx + 1 < arrTextObjects.length) {
				arrLines.push(arrTexts)
				arrTexts = []
			}
		}

		// D: Flush buffer
		if (idx + 1 === arrTextObjects.length) arrLines.push(arrTexts)
	})
	return arrLines
}

/**
 * Render each grouped line to an `<a:p>` paragraph: paragraph props, inherited run options,
 * text runs (and inline/display math), and the closing endParaRPr.
 */
export function renderTextParagraphsXml(
	arrLines: RunProps[][],
	slideObj: SlideObject | TableCell,
	opts: ObjectOptions
): string {
	let strSlideXml = ''
	arrLines.forEach((line) => {
		// A DISPLAY equation owns its whole paragraph: emit the oMathPara wrapper and skip runs.
		// An inline equation (even when it is the line's only run) flows as a run and is emitted below.
		const firstRun = line[0]
		if (line.length === 1 && firstRun?.math && !firstRun.inline) {
			strSlideXml += genXmlMathParagraph(firstRun.math)
			return
		}

		let reqsClosingFontSize = false

		// A: Accumulate the paragraph's children; the `<a:p>` wrapper closes over them at the end.
		let paraXml = ''
		// NOTE: the `<a:pPr ${rtlMode ? ' rtl="1" ' : ''}` seed that used to sit here was dead —
		// `genXmlParagraphProperties` builds its own open tag (with the same `rtlMode` check) and
		// overwrote this on the first run, and nothing read it in between.
		let paragraphPropEmitted = false

		// B: Start paragraph, loop over lines and add text runs
		line.forEach((textObj, idx) => {
			// A: Set line index
			textObj.options._lineIdx = idx

			// A.1: Add soft break if not the first run of the line.
			if (idx > 0 && textObj.options.softBreakBefore) {
				paraXml += voidEl('a:br')
			}

			// B: Inherit pPr-type options from parent shape's `options`.
			// `setOrClear`, not plain assignment: where neither the run nor the shape states one,
			// the key stays off the run's bag rather than being written as an `undefined`. Run
			// option bags are spread — `{ ...itext.options }` in `text-body.ts`, and a placeholder's
			// options onto a slide's in `gen/define/text.ts` — and there the two are not the same.
			setOrClear(textObj.options, 'align', textObj.options.align || opts.align)
			setOrClear(textObj.options, 'lineSpacing', textObj.options.lineSpacing || opts.lineSpacing)
			setOrClear(
				textObj.options,
				'lineSpacingMultiple',
				textObj.options.lineSpacingMultiple || opts.lineSpacingMultiple
			)
			setOrClear(textObj.options, 'indentLevel', textObj.options.indentLevel || opts.indentLevel)
			setOrClear(textObj.options, 'paraSpaceBefore', textObj.options.paraSpaceBefore || opts.paraSpaceBefore)
			setOrClear(textObj.options, 'paraSpaceAfter', textObj.options.paraSpaceAfter || opts.paraSpaceAfter)
			// `??`, not `||`, on these two: `0` is a meaningful margin (flush with the frame, and
			// the override that suppresses a bullet's hanging indent), where the options above have
			// no zero worth stating. A falsy test would silently swap a run's explicit `0` for the
			// shape's value.
			setOrClear(textObj.options, 'paraMarginLeft', textObj.options.paraMarginLeft ?? opts.paraMarginLeft)
			setOrClear(textObj.options, 'paraIndent', textObj.options.paraIndent ?? opts.paraIndent)

			// OOXML allows only one `<a:pPr>` per `<a:p>`, and it must precede any `<a:r>` runs.
			// The paragraph's properties are the FIRST run's, decided once: this used to retry on
			// each subsequent run until one produced non-empty XML, which was unreachable while
			// every `bullet` state wrote something (an omitted one still emitted `marL="0"
			// indent="0"` + `a:buNone`) — and became wrong the moment `bullet: 'inherit'` made an
			// empty pPr possible. A retry would take its properties from a *continuation* run,
			// which by convention states no bullet, so the paragraph got back the very `a:buNone`
			// the first run asked to leave out — and got it appended AFTER that run's `<a:r>`,
			// where a `pPr` is not allowed.
			if (!paragraphPropEmitted) {
				paragraphPropEmitted = true
				// IMPORTANT: Empty "pPr" blocks will generate needs-repair/corrupt msg
				paraXml += genXmlParagraphProperties(textObj, false).replace('<a:pPr></a:pPr>', '')
			}
			// C: Inherit any main options (color, fontSize, etc.)
			// NOTE: We only pass the text.options to genXmlTextRun (not the Slide.options),
			// so the run building function cant just fallback to Slide.color, therefore, we need to do that here before passing options below.
			// FILTER RULE: Hyperlinks should not inherit `color` from main options (let PPT default to local color, eg: blue on MacOS)
			const textOptions = textObj.options as TextPropsOptions & Record<string, unknown>
			Object.entries(opts)
				.filter(([key]) => !(textObj.options.hyperlink && key === 'color'))
				.forEach(([key, val]) => {
					// NOTE: This loop will pick up unecessary keys (`x`, etc.), but it doesnt hurt anything
					if (key !== 'bullet' && !textOptions[key]) textOptions[key] = val
				})

			// D: Add formatted textrun
			// When this paragraph emits bullet markup (`bullet:true` or any object
			// form), strip a single leading bullet glyph (+ optional whitespace) from
			// the first run's text. Otherwise PowerPoint renders two bullets — one
			// from the paragraph-level `<a:buChar/>` and one from the literal glyph
			// in `<a:t>`. Mid-text glyphs are unaffected, and so are `bullet:false`,
			// `bullet:'inherit'` and no-bullet — none of them emit a glyph to double.
			let _textRunObj = textObj
			if (idx === 0 && emitsBulletMarkup(line[0]?.options.bullet) && typeof textObj.text === 'string') {
				const _stripped = textObj.text.replace(/^[\u2022\u25E6\u25AA\u25AB\u25CF\u25CB\u2023\u2043\u2219]\s*/, '')
				if (_stripped !== textObj.text) {
					_textRunObj = { text: _stripped, options: textObj.options }
				}
			}
			// Drop empty-string runs that are pure paragraph-break artifacts. A segment that begins
			// with "\n" (common when mixing RTL/LTR runs) splits into a leading empty piece; once the
			// paragraph break has been applied in STEP 5 that empty piece would otherwise emit a junk
			// `<a:r>...<a:t></a:t></a:r>` trailing the previous line. A *lone* empty run is the line
			// itself (an intentional blank paragraph, eg: "line1\n\nline3"), so keep it when it is the
			// only run on the line.
			const isEmptyBreakArtifact = _textRunObj.text === '' && line.length > 1
			if (_textRunObj.math) {
				// Inline native equation (dn-inline-math): a bare <a14:m><m:oMath> run between plain runs.
				paraXml += genXmlInlineMath(_textRunObj.math)
			} else if (!isEmptyBreakArtifact) {
				paraXml += genXmlTextRun(_textRunObj)
			}

			// E: Flag close fontSize for empty [lineBreak] elements
			if ((!textObj.text && opts.fontSize) || textObj.options.fontSize) {
				reqsClosingFontSize = true
				setOrClear(opts, 'fontSize', opts.fontSize || textObj.options.fontSize)
			}
		})

		/* C: Append 'endParaRPr' (when needed) and close current open paragraph
		 * NOTE: Add 'endParaRPr' with font/size props or PPT default (Arial/18pt en-us) is used making row "too tall"/not honoring options
		 */
		// `sz` rides along in the first two branches only: the last one omits it even when
		// `opts.fontSize` is set, which is why the attribute set is built per-branch.
		const sizedAttrs = (): XmlAttrs => ({
			lang: opts.lang || 'en-US',
			sz: opts.fontSize ? clampFontSizeSz(opts.fontSize) : null,
			dirty: '0',
		})
		if (slideObj._type === SlideObjectType.tablecell && (opts.fontSize || opts.fontFace)) {
			if (opts.fontFace) {
				// Mirror genXmlTextRunProperties: Latin + complex-script slots carry the face; East Asian slot
				// inherits the theme unless `fontFaceEA` is set. Escaping is the builder's job now.
				paraXml += el('a:endParaRPr', sizedAttrs(), [
					raw(voidEl('a:latin', { typeface: opts.fontFace, charset: '0' })),
					opts.fontFaceEA ? raw(voidEl('a:ea', { typeface: opts.fontFaceEA })) : null,
					raw(voidEl('a:cs', { typeface: opts.fontFace })),
				])
			} else {
				paraXml += voidEl('a:endParaRPr', sizedAttrs())
			}
		} else if (reqsClosingFontSize) {
			// Empty [lineBreak] lines should not contain runProp, however, they need to specify fontSize in `endParaRPr`
			paraXml += voidEl('a:endParaRPr', sizedAttrs())
		} else {
			// Added 20180101 to address PPT-2007 issues
			paraXml += voidEl('a:endParaRPr', { lang: opts.lang || 'en-US', dirty: '0' })
		}

		// D: End paragraph
		strSlideXml += el('a:p', null, raw(paraXml))
	})
	return strSlideXml
}
