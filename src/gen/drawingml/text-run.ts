/**
 * ts-pptx: DrawingML text runs & paragraphs
 *
 * The run/paragraph half of text-body generation: paragraph properties
 * (`<a:pPr>`, bullets, spacing), run properties (`<a:rPr>`), the `<a:r>` run
 * builder, the shrink-autofit helper, and the grouping/rendering of a flat run
 * list into `<a:p>` paragraphs. The `text-body.ts` container layer builds on top.
 */

import { BulletType, SlideObjectType } from '../../core-enums.js'
import { DEF_BULLET_MARGIN, DEF_TEXT_GLOW, DEF_TEXT_SHADOW } from '../../core-enums-internal.js'
import type {
	ObjectOptions,
	TableCell,
	TextFitShrinkProps,
	TextProps,
	TextPropsOptions,
} from '../../core-interfaces.js'
import type { SlideObject } from '../../types/internal.js'
import { createColorElement } from './color.js'
import { createGlowElement, createShadowElement } from './effect.js'
import { genXmlColorSelection } from './fill.js'
import { inch2Emu, lineWidthToEmu, valToPts } from '../../units-internal.js'
import { FIXED_PCT_PER_PERCENT, PERCENT_SCALE, ptToHundredths } from '../../units.js'
import { warn } from '../../log.js'
import { el, raw, voidEl, type XmlAttrs } from '../oxml/el.js'
import { clampCharSpacingSpc, clampFontSizeSz, clampLineSpacingPts } from './clamp.js'
import { genXmlInlineMath, genXmlMathParagraph } from './math.js'

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
	let bulletMarL = valToPts(DEF_BULLET_MARGIN)

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
			const val = Math.round(opts.lineSpacingMultiple * PERCENT_SCALE)
			strXmlLnSpc = el('a:lnSpc', null, raw(voidEl('a:spcPct', { val })))
		}

		// OPTION: indent
		if (opts.indentLevel && !isNaN(Number(opts.indentLevel)) && opts.indentLevel > 0) {
			paragraphPropXml += ` lvl="${opts.indentLevel}"`
		}

		// OPTION: Paragraph Spacing: Before/After
		if (opts.paraSpaceBefore && !isNaN(Number(opts.paraSpaceBefore)) && opts.paraSpaceBefore > 0) {
			strXmlParaSpc += el('a:spcBef', null, raw(voidEl('a:spcPts', { val: ptToHundredths(opts.paraSpaceBefore) })))
		}
		if (opts.paraSpaceAfter && !isNaN(Number(opts.paraSpaceAfter)) && opts.paraSpaceAfter > 0) {
			strXmlParaSpc += el('a:spcAft', null, raw(voidEl('a:spcPts', { val: ptToHundredths(opts.paraSpaceAfter) })))
		}

		// OPTION: bullet
		// NOTE: OOXML uses the unicode character set for Bullets
		// EX: Unicode Character 'BULLET' (U+2022) ==> '<a:buChar char="&#x2022;"/>'
		if (typeof opts.bullet === 'object') {
			const bulletImage = opts.bullet.image
			const isPictureBullet = !!(bulletImage && (bulletImage.path || bulletImage.data))
			if (opts.bullet?.indent) bulletMarL = valToPts(opts.bullet.indent)
			// `buClr` colors a glyph/number; it has no effect on a picture bullet, so skip it for `buBlip`.
			if (opts.bullet.color && !isPictureBullet)
				strXmlBulletColor = el('a:buClr', null, raw(createColorElement(opts.bullet.color)))

			// `<a:buSzPct/>` val is thousandths of a percent; ST_TextBulletSizePercent allows 25%-400%
			let bulletSizePct = PERCENT_SCALE
			if (opts.bullet.size !== undefined) {
				const bulletSize = Number(opts.bullet.size)
				// 25–400% is the range PowerPoint's bullet-size dialog accepts (and the
				// range `<a:buSzPct>` renders sensibly); values outside it are rejected
				// rather than clamped so the caller notices the bad input.
				if (isNaN(bulletSize) || bulletSize < 25 || bulletSize > 400) {
					warn('`bullet.size` must be a percentage between 25 and 400!')
				} else {
					bulletSizePct = Math.round(bulletSize * FIXED_PCT_PER_PERCENT)
				}
			}
			const strXmlBulletSize = voidEl('a:buSzPct', { val: bulletSizePct })
			// NOTE: the builder escapes `typeface`, so the manual `encodeXmlEntities` that used to
			// wrap it here is gone — keeping both would double-escape (`&` -> `&amp;amp;`).
			const strXmlBulletFont = opts.bullet.fontFace ? voidEl('a:buFont', { typeface: opts.bullet.fontFace }) : ''

			// Every bullet form below hangs the first line by the same margin; the attributes belong to
			// the hand-built open tag above, so they stay a fragment rather than becoming builder attrs.
			const bulletIndentAttrs = (): string =>
				` marL="${
					opts.indentLevel && opts.indentLevel > 0 ? bulletMarL + bulletMarL * opts.indentLevel : bulletMarL
				}" indent="-${bulletMarL}"`

			if (isPictureBullet) {
				// Picture bullet: <a:buBlip> references a slide media rel registered in addText() (`_rId`).
				// No `buFont` (there is no glyph typeface), but `buSzPct` still scales the image height.
				paragraphPropXml += bulletIndentAttrs()
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
					warn('picture `bullet.image` could not be embedded; using a default bullet glyph')
					strXmlBullet = strXmlBulletSize + strXmlBulletFont + buChar(BulletType.DEFAULT)
				}
			} else if (opts.bullet.type && opts.bullet.type.toString().toLowerCase() === 'number') {
				paragraphPropXml += bulletIndentAttrs()
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
					warn('`bullet.characterCode` should be a 4-digit unicode character (ex: 22AB)!')
					bulletCode = BulletType.DEFAULT
				}

				paragraphPropXml += bulletIndentAttrs()
				strXmlBullet = strXmlBulletSize + strXmlBulletFont + buChar(bulletCode)
			} else {
				paragraphPropXml += bulletIndentAttrs()
				strXmlBullet = strXmlBulletSize + strXmlBulletFont + buChar(BulletType.DEFAULT)
			}
		} else if (opts.bullet) {
			paragraphPropXml += ` marL="${
				opts.indentLevel && opts.indentLevel > 0 ? bulletMarL + bulletMarL * opts.indentLevel : bulletMarL
			}" indent="-${bulletMarL}"`
			strXmlBullet = voidEl('a:buSzPct', { val: '100000' }) + buChar(BulletType.DEFAULT)
		} else if (!opts.bullet) {
			// We only add this when the user explicitely asks for no bullet, otherwise, it can override the master defaults!
			paragraphPropXml += ' indent="0" marL="0"' // FIX: specify zero indent and marL or default will be hanging paragraph
			strXmlBullet = voidEl('a:buNone')
		}

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
		if (opts.color) runProps += genXmlColorSelection({ color: opts.color, transparency: opts.transparency })
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
			throw new Error("ERROR: text `hyperlink` option should be an object. Ex: `hyperlink:{url:'https://github.com'}` ")
		else if (!opts.hyperlink.url && !opts.hyperlink.slide && !opts.hyperlink.action)
			throw new Error("ERROR: 'hyperlink requires either `url`, `slide`, or `action`'")
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
	// NOTE: fontScale/lnSpcReduction are authored as a percent (0-100); OOXML stores them in 1000ths of a percent.
	const pct = (val: number | undefined, name: string): number | null => {
		if (val === undefined || val === null) return null
		if (typeof val !== 'number' || isNaN(val) || val < 0 || val > 100) {
			warn(`fit.${name} must be a number between 0 and 100 (percent); received ${String(val)} - attribute ignored.`)
			return null
		}
		return Math.round(val * FIXED_PCT_PER_PERCENT)
	}

	// `pct` returns null for an absent/rejected value, and the builder omits null attributes.
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
		} else if (arrTexts.length > 0 && textObj.options.bullet && arrTexts.length > 0) {
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
		// NOTE: `rtlMode` is like other opts, its propagated up to each text:options, so just check the 1st one
		let paragraphPropXml = `<a:pPr ${line[0]?.options?.rtlMode ? ' rtl="1" ' : ''}`
		let paragraphPropEmitted = false

		// B: Start paragraph, loop over lines and add text runs
		line.forEach((textObj, idx) => {
			// A: Set line index
			textObj.options._lineIdx = idx

			// A.1: Add soft break if not the first run of the line.
			if (idx > 0 && textObj.options.softBreakBefore) {
				paraXml += voidEl('a:br')
			}

			// B: Inherit pPr-type options from parent shape's `options`
			textObj.options.align = textObj.options.align || opts.align
			textObj.options.lineSpacing = textObj.options.lineSpacing || opts.lineSpacing
			textObj.options.lineSpacingMultiple = textObj.options.lineSpacingMultiple || opts.lineSpacingMultiple
			textObj.options.indentLevel = textObj.options.indentLevel || opts.indentLevel
			textObj.options.paraSpaceBefore = textObj.options.paraSpaceBefore || opts.paraSpaceBefore
			textObj.options.paraSpaceAfter = textObj.options.paraSpaceAfter || opts.paraSpaceAfter

			// OOXML allows only one `<a:pPr>` per `<a:p>`, and it must precede any `<a:r>` runs.
			// Emit paragraph properties exactly once, derived from the first run that yields non-empty pPr XML.
			if (!paragraphPropEmitted) {
				paragraphPropXml = genXmlParagraphProperties(textObj, false)
				const cleaned = paragraphPropXml.replace('<a:pPr></a:pPr>', '') // IMPORTANT: Empty "pPr" blocks will generate needs-repair/corrupt msg
				if (cleaned) {
					paraXml += cleaned
					paragraphPropEmitted = true
				}
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
			// in `<a:t>`. Mid-text glyphs and `bullet:false`/no-bullet are unaffected.
			let _textRunObj = textObj
			if (idx === 0 && line[0]?.options.bullet && typeof textObj.text === 'string') {
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
				opts.fontSize = opts.fontSize || textObj.options.fontSize
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
