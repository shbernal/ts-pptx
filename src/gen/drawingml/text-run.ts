/**
 * PptxGenJS: DrawingML text runs & paragraphs
 *
 * The run/paragraph half of text-body generation: paragraph properties
 * (`<a:pPr>`, bullets, spacing), run properties (`<a:rPr>`), the `<a:r>` run
 * builder, the shrink-autofit helper, and the grouping/rendering of a flat run
 * list into `<a:p>` paragraphs. The `text-body.ts` container layer builds on top.
 */

import { BulletType, DEF_BULLET_MARGIN, DEF_TEXT_GLOW, DEF_TEXT_SHADOW, SlideObjectType } from '../../core-enums.js'
import type {
	ObjectOptions,
	SlideObject,
	TableCell,
	TextFitShrinkProps,
	TextProps,
	TextPropsOptions,
} from '../../core-interfaces.js'
import {
	createColorElement,
	createGlowElement,
	createShadowElement,
	encodeXmlEntities,
	genXmlColorSelection,
	inch2Emu,
	lineWidthToEmu,
	valToPts,
} from '../../gen-utils.js'
import { FIXED_PCT_PER_PERCENT, PERCENT_SCALE, ptToHundredths } from '../../units.js'
import { warn } from '../../log.js'
import { clampCharSpacingSpc, clampFontSizeSz, clampLineSpacingPts } from './clamp.js'
import { genXmlInlineMath, genXmlMathParagraph } from './math.js'

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
			strXmlLnSpc = `<a:lnSpc><a:spcPts val="${clampLineSpacingPts(opts.lineSpacing)}"/></a:lnSpc>`
		} else if (opts.lineSpacingMultiple) {
			strXmlLnSpc = `<a:lnSpc><a:spcPct val="${Math.round(opts.lineSpacingMultiple * PERCENT_SCALE)}"/></a:lnSpc>`
		}

		// OPTION: indent
		if (opts.indentLevel && !isNaN(Number(opts.indentLevel)) && opts.indentLevel > 0) {
			paragraphPropXml += ` lvl="${opts.indentLevel}"`
		}

		// OPTION: Paragraph Spacing: Before/After
		if (opts.paraSpaceBefore && !isNaN(Number(opts.paraSpaceBefore)) && opts.paraSpaceBefore > 0) {
			strXmlParaSpc += `<a:spcBef><a:spcPts val="${ptToHundredths(opts.paraSpaceBefore)}"/></a:spcBef>`
		}
		if (opts.paraSpaceAfter && !isNaN(Number(opts.paraSpaceAfter)) && opts.paraSpaceAfter > 0) {
			strXmlParaSpc += `<a:spcAft><a:spcPts val="${ptToHundredths(opts.paraSpaceAfter)}"/></a:spcAft>`
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
				strXmlBulletColor = `<a:buClr>${createColorElement(opts.bullet.color)}</a:buClr>`

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
			const strXmlBulletSize = `<a:buSzPct val="${bulletSizePct}"/>`
			const strXmlBulletFont = opts.bullet.fontFace
				? `<a:buFont typeface="${encodeXmlEntities(opts.bullet.fontFace)}"/>`
				: ''

			if (isPictureBullet) {
				// Picture bullet: <a:buBlip> references a slide media rel registered in addText() (`_rId`).
				// No `buFont` (there is no glyph typeface), but `buSzPct` still scales the image height.
				paragraphPropXml += ` marL="${
					opts.indentLevel && opts.indentLevel > 0 ? bulletMarL + bulletMarL * opts.indentLevel : bulletMarL
				}" indent="-${bulletMarL}"`
				if (opts.bullet._rId) {
					if (opts.bullet._rIdSvg) {
						// SVG bullet: the blip embeds the PNG preview (`_rId`) and references the SVG via the
						// `asvg:svgBlip` extension (`_rIdSvg`), the same dual-rel form addImage() emits for SVG.
						strXmlBullet =
							`${strXmlBulletSize}<a:buBlip><a:blip r:embed="rId${opts.bullet._rId}">` +
							'<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
							`<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId${opts.bullet._rIdSvg}"/>` +
							'</a:ext></a:extLst></a:blip></a:buBlip>'
					} else {
						strXmlBullet = `${strXmlBulletSize}<a:buBlip><a:blip r:embed="rId${opts.bullet._rId}"/></a:buBlip>`
					}
				} else {
					// rel was not registered (eg: bullet on a context without a slide target) - fall back to a glyph
					warn('picture `bullet.image` could not be embedded; using a default bullet glyph')
					strXmlBullet = `${strXmlBulletSize}${strXmlBulletFont}<a:buChar char="${BulletType.DEFAULT}"/>`
				}
			} else if (opts.bullet.type && opts.bullet.type.toString().toLowerCase() === 'number') {
				paragraphPropXml += ` marL="${
					opts.indentLevel && opts.indentLevel > 0 ? bulletMarL + bulletMarL * opts.indentLevel : bulletMarL
				}" indent="-${bulletMarL}"`
				strXmlBullet = `${strXmlBulletSize}${strXmlBulletFont || '<a:buFont typeface="+mj-lt"/>'}<a:buAutoNum type="${opts.bullet.numberType || 'arabicPeriod'}" startAt="${
					opts.bullet.numberStartAt || '1'
				}"/>`
			} else if (opts.bullet.characterCode) {
				let bulletCode = `&#x${opts.bullet.characterCode};`

				// Check value for hex-ness (s/b 4 char hex)
				if (!/^[0-9A-Fa-f]{4}$/.test(opts.bullet.characterCode)) {
					warn('`bullet.characterCode` should be a 4-digit unicode character (ex: 22AB)!')
					bulletCode = BulletType.DEFAULT
				}

				paragraphPropXml += ` marL="${
					opts.indentLevel && opts.indentLevel > 0 ? bulletMarL + bulletMarL * opts.indentLevel : bulletMarL
				}" indent="-${bulletMarL}"`
				strXmlBullet = strXmlBulletSize + strXmlBulletFont + '<a:buChar char="' + bulletCode + '"/>'
			} else {
				paragraphPropXml += ` marL="${
					opts.indentLevel && opts.indentLevel > 0 ? bulletMarL + bulletMarL * opts.indentLevel : bulletMarL
				}" indent="-${bulletMarL}"`
				strXmlBullet = `${strXmlBulletSize}${strXmlBulletFont}<a:buChar char="${BulletType.DEFAULT}"/>`
			}
		} else if (opts.bullet) {
			paragraphPropXml += ` marL="${
				opts.indentLevel && opts.indentLevel > 0 ? bulletMarL + bulletMarL * opts.indentLevel : bulletMarL
			}" indent="-${bulletMarL}"`
			strXmlBullet = `<a:buSzPct val="100000"/><a:buChar char="${BulletType.DEFAULT}"/>`
		} else if (!opts.bullet) {
			// We only add this when the user explicitely asks for no bullet, otherwise, it can override the master defaults!
			paragraphPropXml += ' indent="0" marL="0"' // FIX: specify zero indent and marL or default will be hanging paragraph
			strXmlBullet = '<a:buNone/>'
		}

		// OPTION: tabStops
		if (opts.tabStops && Array.isArray(opts.tabStops)) {
			const tabStopsXml = opts.tabStops
				.map((stop) => `<a:tab pos="${inch2Emu(stop.position || 1)}" algn="${stop.alignment || 'l'}"/>`)
				.join('')
			strXmlTabStops = `<a:tabLst>${tabStopsXml}</a:tabLst>`
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
	runProps +=
		'<' + runPropsTag + ' lang="' + (opts.lang ? opts.lang : 'en-US') + '"' + (opts.lang ? ' altLang="en-US"' : '')
	runProps += opts.fontSize ? ` sz="${clampFontSizeSz(opts.fontSize)}"` : '' // NOTE: clamp+round so sizes like '7.5' or out-of-range values wont cause corrupt presentations
	runProps += opts?.bold ? ` b="${opts.bold ? '1' : '0'}"` : ''
	runProps += opts?.italic ? ` i="${opts.italic ? '1' : '0'}"` : ''

	runProps += opts?.strike ? ` strike="${typeof opts.strike === 'string' ? opts.strike : 'sngStrike'}"` : ''
	runProps += opts?.caps ? ` cap="${opts.caps}"` : ''
	if (typeof opts.underline === 'object' && opts.underline?.style) {
		runProps += ` u="${opts.underline.style}"`
	} else if (opts.hyperlink) {
		runProps += ' u="sng"'
	}
	if (opts.baseline) {
		runProps += ` baseline="${Math.round(opts.baseline * 50)}"`
	} else if (opts.subscript) {
		runProps += ' baseline="-40000"'
	} else if (opts.superscript) {
		runProps += ' baseline="30000"'
	}
	runProps += opts.charSpacing ? ` spc="${clampCharSpacingSpc(opts.charSpacing)}" kern="0"` : '' // IMPORTANT: Also disable kerning; otherwise text won't actually expand
	runProps += ' dirty="0">'
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
			runProps += `<a:ln w="${lineWidthToEmu(opts.outline.size || 0.75)}">${genXmlColorSelection(opts.outline.color || 'FFFFFF')}</a:ln>`
		}
		if (opts.color) runProps += genXmlColorSelection({ color: opts.color, transparency: opts.transparency })
		// EFFECTS: glow and shadow share a single <a:effectLst> (only one is allowed per CT_TextCharacterProperties; glow precedes shadow per CT_EffectList)
		if (opts.glow || hasShadow) {
			runProps += '<a:effectLst>'
			if (opts.glow) runProps += createGlowElement(opts.glow, DEF_TEXT_GLOW)
			if (hasShadow) runProps += createShadowElement(opts.shadow, DEF_TEXT_SHADOW)
			runProps += '</a:effectLst>'
		}
		if (opts.highlight) runProps += `<a:highlight>${createColorElement(opts.highlight)}</a:highlight>`
		if (typeof opts.underline === 'object' && opts.underline.color)
			runProps += `<a:uFill>${genXmlColorSelection(opts.underline.color)}</a:uFill>`
		if (opts.fontFace) {
			// Match how PowerPoint writes a font picked from the UI: the chosen typeface goes in the
			// Latin (`<a:latin>`) and complex-script (`<a:cs>`) slots. The East Asian slot (`<a:ea>`) is
			// only written when an EA face is explicitly chosen (`fontFaceEA`); otherwise it inherits the
			// theme. Forcing a Latin-only font into `<a:ea>` — especially with the bogus charset values
			// PowerPoint never emits on ea/cs — duplicates/ghosts text in Office 365.
			// NOTE: order must be latin, ea, cs per CT_TextCharacterProperties.
			runProps += `<a:latin typeface="${opts.fontFace}" pitchFamily="34" charset="0"/>`
			if (opts.fontFaceEA) runProps += `<a:ea typeface="${opts.fontFaceEA}"/>`
			runProps += `<a:cs typeface="${opts.fontFace}"/>`
		}
	}

	// Hyperlink support
	if (opts.hyperlink) {
		if (typeof opts.hyperlink !== 'object')
			throw new Error("ERROR: text `hyperlink` option should be an object. Ex: `hyperlink:{url:'https://github.com'}` ")
		else if (!opts.hyperlink.url && !opts.hyperlink.slide)
			throw new Error("ERROR: 'hyperlink requires either `url` or `slide`'")
		else if (opts.hyperlink.url) {
			// runProps += '<a:uFill>'+ genXmlColorSelection('0000FF') +'</a:uFill>'; // Breaks PPT2010!
			runProps += `<a:hlinkClick r:id="rId${opts.hyperlink._rId}" invalidUrl="" action="" tgtFrame="" tooltip="${
				opts.hyperlink.tooltip ? encodeXmlEntities(opts.hyperlink.tooltip) : ''
			}" history="1" highlightClick="0" endSnd="0"${opts.color ? '>' : '/>'}`
		} else if (opts.hyperlink.slide) {
			runProps += `<a:hlinkClick r:id="rId${opts.hyperlink._rId}" action="ppaction://hlinksldjump" tooltip="${
				opts.hyperlink.tooltip ? encodeXmlEntities(opts.hyperlink.tooltip) : ''
			}"${opts.color ? '>' : '/>'}`
		}
		if (opts.color) {
			runProps += ' <a:extLst>'
			runProps += '  <a:ext uri="{A12FA001-AC4F-418D-AE19-62706E023703}">'
			runProps +=
				'   <ahyp:hlinkClr xmlns:ahyp="http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor" val="tx"/>'
			runProps += '  </a:ext>'
			runProps += ' </a:extLst>'
			runProps += '</a:hlinkClick>'
		}
	}

	// END runProperties
	runProps += `</${runPropsTag}>`

	return runProps
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
	return `<a:r>${genXmlTextRunProperties(textObj.options ?? {}, false)}<a:t>${encodeXmlEntities(String(textObj.text))}</a:t></a:r>`
}

/**
 * Builds `<a:normAutofit>` with explicit fontScale/lnSpcReduction for "shrink text on overflow"
 * @param {TextFitShrinkProps} fit - shrink fit options
 * @return {string} XML string (`<a:normAutofit .../>`)
 * @see ECMA-376 CT_TextNormAutofit (attributes in 1000ths of a percent)
 */
export function genXmlNormAutofit(fit: TextFitShrinkProps): string {
	let attrs = ''

	// NOTE: fontScale/lnSpcReduction are authored as a percent (0-100); OOXML stores them in 1000ths of a percent.
	const pct = (val: number | undefined, name: string): number | null => {
		if (val === undefined || val === null) return null
		if (typeof val !== 'number' || isNaN(val) || val < 0 || val > 100) {
			warn(`fit.${name} must be a number between 0 and 100 (percent); received ${String(val)} - attribute ignored.`)
			return null
		}
		return Math.round(val * FIXED_PCT_PER_PERCENT)
	}

	const fontScale = pct(fit.fontScale, 'fontScale')
	if (fontScale !== null) attrs += ` fontScale="${fontScale}"`
	const lnSpcReduction = pct(fit.lnSpcReduction, 'lnSpcReduction')
	if (lnSpcReduction !== null) attrs += ` lnSpcReduction="${lnSpcReduction}"`

	return `<a:normAutofit${attrs}/>`
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

		// A: Start paragraph, add paraProps
		strSlideXml += '<a:p>'
		// NOTE: `rtlMode` is like other opts, its propagated up to each text:options, so just check the 1st one
		let paragraphPropXml = `<a:pPr ${line[0]?.options?.rtlMode ? ' rtl="1" ' : ''}`
		let paragraphPropEmitted = false

		// B: Start paragraph, loop over lines and add text runs
		line.forEach((textObj, idx) => {
			// A: Set line index
			textObj.options._lineIdx = idx

			// A.1: Add soft break if not the first run of the line.
			if (idx > 0 && textObj.options.softBreakBefore) {
				strSlideXml += '<a:br/>'
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
					strSlideXml += cleaned
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
				strSlideXml += genXmlInlineMath(_textRunObj.math)
			} else if (!isEmptyBreakArtifact) {
				strSlideXml += genXmlTextRun(_textRunObj)
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
		if (slideObj._type === SlideObjectType.tablecell && (opts.fontSize || opts.fontFace)) {
			if (opts.fontFace) {
				strSlideXml +=
					`<a:endParaRPr lang="${opts.lang || 'en-US'}"` +
					(opts.fontSize ? ` sz="${clampFontSizeSz(opts.fontSize)}"` : '') +
					' dirty="0">'
				// Mirror genXmlTextRunProperties: Latin + complex-script slots carry the face; East Asian slot
				// inherits the theme unless `fontFaceEA` is set.
				strSlideXml += `<a:latin typeface="${opts.fontFace}" charset="0"/>`
				if (opts.fontFaceEA) strSlideXml += `<a:ea typeface="${opts.fontFaceEA}"/>`
				strSlideXml += `<a:cs typeface="${opts.fontFace}"/>`
				strSlideXml += '</a:endParaRPr>'
			} else {
				strSlideXml +=
					`<a:endParaRPr lang="${opts.lang || 'en-US'}"` +
					(opts.fontSize ? ` sz="${clampFontSizeSz(opts.fontSize)}"` : '') +
					' dirty="0"/>'
			}
		} else if (reqsClosingFontSize) {
			// Empty [lineBreak] lines should not contain runProp, however, they need to specify fontSize in `endParaRPr`
			strSlideXml +=
				`<a:endParaRPr lang="${opts.lang || 'en-US'}"` +
				(opts.fontSize ? ` sz="${clampFontSizeSz(opts.fontSize)}"` : '') +
				' dirty="0"/>'
		} else {
			strSlideXml += `<a:endParaRPr lang="${opts.lang || 'en-US'}" dirty="0"/>` // Added 20180101 to address PPT-2007 issues
		}

		// D: End paragraph
		strSlideXml += '</a:p>'
	})
	return strSlideXml
}
