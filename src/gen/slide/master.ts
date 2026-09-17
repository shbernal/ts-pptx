/**
 * ts-pptx: slide-master parts
 *
 * The slide master (`slideMaster1.xml`) and its `.rels`, plus the master text
 * styles (`<p:txStyles>`): per-level default values mirroring the built-in
 * Office master, with any `MasterTextStyleProps` overrides layered on top.
 */

import { CRLF, LAYOUT_IDX_SERIES_BASE, LEVEL_MARGINS_EMU, XML_DECL } from '../../constants-internal.js'
import type { MasterBulletProps, MasterTextStyleLevel, MasterTextStyleProps } from '../../types/index.js'
import type { SlideLayoutInternal, SlideMasterInternal } from '../../types/internal.js'
import { createColorElement, rejectEmptyColor } from '../drawingml/color.js'
import { lvlPPr, themeFontDefRPr } from '../drawingml/list-style.js'
import { buAutoNumEl, buCharEl } from '../drawingml/bullet.js'
import { clampFontSizeSz, clampParaIndentInchesEmu, clampParaMarginInchesEmu } from '../drawingml/clamp.js'
import { EMU_PER_INCH, HUNDREDTHS_PER_POINT } from '../../units.js'
import { warn } from '../../diagnostics.js'
import { el, raw, voidEl, type XmlAttrs } from '../oxml/el.js'
import { defaultRelIdStart, slideObjectRelationsToXml, slideObjectToXml } from './object.js'
import type { RendererTable } from './objects/shared.js'
import { PML_ROOT_NS } from '../../ooxml/namespaces.js'
import { SLIDE_LAYOUT_REL, THEME_REL } from '../../ooxml/rel-types.js'
import { DEFAULT_COLOR_MAP } from '../../ooxml/st-enums.js'
import { textAlignToken } from '../../ooxml/text-align.js'
import { xsdBoolIfTrue } from '../../ooxml/xsd-boolean.js'
import { SLIDE_MASTER_PATH, slideLayoutPath, targetFromPptSubpart } from '../opc/part-paths.js'

// Default per-level values mirroring the built-in Office master (used as the base that
// `MasterTextStyleProps` overrides are layered onto). `bu` describes the level's default
// bullet: 'none' -> <a:buNone/>, undefined -> no bullet element (otherStyle), or a glyph.
interface MasterLevelDefault {
	marL: number // EMU
	indent?: number // EMU (omitted when undefined)
	algn: string // OOXML algn value
	spcBefPct?: number // <a:spcBef> percent (×1000); omitted when undefined
	bu?: 'none' | { char: string; font: string } // default bullet; undefined => emit no bullet element
	sz: number // <a:defRPr@sz>
	font: 'mj' | 'mn' // major (heading) vs minor (body) theme font family
}
const MASTER_TITLE_DEFAULT: MasterLevelDefault = {
	marL: 0,
	algn: 'ctr',
	spcBefPct: 0,
	bu: 'none',
	sz: 4400,
	font: 'mj',
}
const MASTER_BODY_DEFAULTS: MasterLevelDefault[] = [
	{
		marL: 342900,
		indent: -342900,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 3200,
		font: 'mn',
	},
	{
		marL: 742950,
		indent: -285750,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '–', font: 'Arial' },
		sz: 2800,
		font: 'mn',
	},
	{
		marL: 1143000,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 2400,
		font: 'mn',
	},
	{
		marL: 1600200,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '–', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
	{
		marL: 2057400,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '»', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
	{
		marL: 2514600,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
	{
		marL: 2971800,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
	{
		marL: 3429000,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
	{
		marL: 3886200,
		indent: -228600,
		algn: 'l',
		spcBefPct: 20000,
		bu: { char: '•', font: 'Arial' },
		sz: 2000,
		font: 'mn',
	},
]
const MASTER_OTHER_DEFAULTS: MasterLevelDefault[] = LEVEL_MARGINS_EMU.map((marL) => ({
	marL,
	algn: 'l',
	sz: 1800,
	font: 'mn' as const,
}))

function masterAlignAttr(align: MasterTextStyleLevel['align']): string {
	// No `align` on the level: `''` omits `@algn` entirely so the level inherits from the theme,
	// which is NOT the same as pinning it to `l`. A value the write API does not name does the same.
	return align === undefined ? '' : (textAlignToken(align) ?? '')
}

/** Build the bullet element for a master level: caller override wins over the level default. */
function masterBulletXml(
	bulletOverride: boolean | MasterBulletProps | undefined,
	base: MasterLevelDefault['bu']
): string {
	// Explicit override
	if (bulletOverride === false) return voidEl('a:buNone')
	if (bulletOverride && typeof bulletOverride === 'object') {
		const font = bulletOverride.fontFace ? voidEl('a:buFont', { typeface: bulletOverride.fontFace }) : ''
		if (bulletOverride.type === 'number') {
			const startAt = typeof bulletOverride.numberStartAt === 'number' ? bulletOverride.numberStartAt : undefined
			return font + buAutoNumEl(bulletOverride.numberType || 'arabicPeriod', startAt, 'master textStyles bullet')
		}
		// character bullet (default)
		const buFont = bulletOverride.fontFace
			? font
			: voidEl('a:buFont', { typeface: 'Arial', pitchFamily: 34, charset: 0 })
		return buFont + buCharEl(bulletOverride.characterCode, '•', 'master textStyles bullet')
	}
	// No override (undefined / true): keep the level's default bullet
	if (base === 'none') return voidEl('a:buNone')
	if (base && typeof base === 'object')
		return (
			voidEl('a:buFont', { typeface: base.font, pitchFamily: 34, charset: 0 }) +
			buCharEl(undefined, base.char, 'master textStyles bullet')
		)
	return '' // otherStyle: no bullet element by default
}

/**
 * One length-valued override on a master text-style level, in EMU, or the level's default when the
 * caller stated something we cannot use.
 *
 * `marginLeft` and `indent` used to reach their clamps only when they were finite numbers and fall
 * back in silence otherwise, so `NaN`, `Infinity` and a string such as `'0.5in'` all became the
 * Office default with nothing said -- while `fontSize` on the very same level warned for exactly
 * those inputs. Silent coercion of invalid input is a footgun: the deck comes out wrong and the
 * caller has nothing to go on.
 *
 * `undefined` is the caller stating nothing, which is not a problem and is not reported.
 * @param stated - the caller's value, in inches
 * @param option - the option's name, as `MasterTextStyleLevel` spells it
 * @param fallback - the level's default, already in EMU
 * @param resolve - the clamp that converts a usable value to EMU
 */
function masterLevelLength<T extends number | undefined>(
	stated: number | undefined,
	option: string,
	fallback: T,
	resolve: (inches: number) => number
): number | T {
	if (stated === undefined) return fallback
	if (typeof stated === 'number' && Number.isFinite(stated)) return resolve(stated)
	// A level with no default of its own omits the attribute; say so rather than quoting `NaN` in.
	const kept = fallback === undefined ? 'the level default' : `${fallback / EMU_PER_INCH}in`
	warn(
		'master/invalid-text-style-length',
		`master textStyles ${option} "${String(stated)}" is invalid; keeping ${kept}.`
	)
	return fallback
}

/** Serialize one `<a:lvlNpPr>` from its default, layering an optional caller override. */
function masterLevelXml(levelNum: number, base: MasterLevelDefault, levelOverride: MasterTextStyleLevel = {}): string {
	const marL = masterLevelLength(levelOverride.marginLeft, 'marginLeft', base.marL, (inches) =>
		clampParaMarginInchesEmu(inches, 'master textStyles marginLeft')
	)
	const indentEmu = masterLevelLength(levelOverride.indent, 'indent', base.indent, (inches) =>
		clampParaIndentInchesEmu(inches, 'master textStyles indent')
	)
	const algn = (levelOverride.align && masterAlignAttr(levelOverride.align)) || base.algn

	let sz = base.sz
	// `!== undefined` rather than `typeof === 'number'`: a string such as `'18pt'` is as much a
	// stated value we cannot use as `NaN` is, and it used to fall through to the default unreported.
	if (levelOverride.fontSize !== undefined) {
		if (
			typeof levelOverride.fontSize !== 'number' ||
			!Number.isFinite(levelOverride.fontSize) ||
			levelOverride.fontSize <= 0
		)
			warn(
				'master/invalid-text-style-font-size',
				`master textStyles fontSize "${String(levelOverride.fontSize)}" is invalid; keeping default ${base.sz / HUNDREDTHS_PER_POINT}pt.`
			)
		else sz = clampFontSizeSz(levelOverride.fontSize, 'master textStyles fontSize')
	}
	// `''` takes the same `tx1` an omitted colour takes; it is reported rather than silently
	// resolved, because it is the caller's missing value and not a way to ask for the default.
	rejectEmptyColor(levelOverride.color, 'textStyles color')
	const colorXml = levelOverride.color ? createColorElement(levelOverride.color) : voidEl('a:schemeClr', { val: 'tx1' })
	const latinXml = levelOverride.fontFace
		? voidEl('a:latin', { typeface: levelOverride.fontFace })
		: voidEl('a:latin', { typeface: `+${base.font}-lt` })

	return lvlPPr(levelNum, { marL, indent: indentEmu, algn }, [
		typeof base.spcBefPct === 'number'
			? raw(el('a:spcBef', null, raw(voidEl('a:spcPct', { val: base.spcBefPct }))))
			: null,
		raw(masterBulletXml(levelOverride.bullet, base.bu)),
		raw(
			themeFontDefRPr(
				base.font,
				{ sz, b: xsdBoolIfTrue(levelOverride.bold), i: xsdBoolIfTrue(levelOverride.italic), kern: 1200 },
				colorXml,
				latinXml
			)
		),
	])
}

/** Clamp a caller-provided per-level override array to the 9 valid list levels, warning on overflow. */
function masterLevelOverrides(levels: MasterTextStyleLevel[] | undefined, group: string): MasterTextStyleLevel[] {
	if (!Array.isArray(levels)) return []
	if (levels.length > 9)
		warn(
			'master/too-many-text-style-levels',
			`master textStyles.${group} has ${levels.length} levels; only the first 9 are used.`
		)
	return levels.slice(0, 9)
}

/**
 * Build the `<p:txStyles>` block from caller overrides layered onto the Office master defaults.
 * Only invoked when `defineSlideMaster({ textStyles })` was set; the unconfigured deck keeps the
 * verbatim default literal in `makeXmlMaster` for byte-identical output.
 */
function makeXmlMasterTxStyles(textStyles: MasterTextStyleProps): string {
	const title = masterLevelXml(1, MASTER_TITLE_DEFAULT, textStyles.title)
	const bodyOverrides = masterLevelOverrides(textStyles.body, 'body')
	const body = MASTER_BODY_DEFAULTS.map((base, i) => masterLevelXml(i + 1, base, bodyOverrides[i])).join('')
	const otherOverrides = masterLevelOverrides(textStyles.other, 'other')
	const other = MASTER_OTHER_DEFAULTS.map((base, i) => masterLevelXml(i + 1, base, otherOverrides[i])).join('')
	return el('p:txStyles', null, [
		raw(el('p:titleStyle', null, raw(title))),
		raw(el('p:bodyStyle', null, raw(body))),
		raw(el('p:otherStyle', null, [raw(el('a:defPPr', null, raw(voidEl('a:defRPr', { lang: 'en-US' })))), raw(other)])),
	])
}

/**
 * The built-in Office master's `<p:txStyles>`, used verbatim when `defineSlideMaster({ textStyles })`
 * was never called. Deliberately NOT built via `masterLevelXml`: a configured title level always
 * carries `marL="0"` (from `MASTER_TITLE_DEFAULT.marL`), but the true unconfigured default titleStyle
 * has no `marL`/`indent` attribute at all — a pre-existing asymmetry between the two paths, preserved
 * exactly rather than unified.
 */
function makeXmlMasterDefaultTxStyles(): string {
	const defaultLevel = (
		n: number,
		attrs: XmlAttrs,
		bullet: string,
		sz: number,
		font: 'mj' | 'mn',
		spcBefPct?: number
	): string =>
		lvlPPr(n, attrs, [
			typeof spcBefPct === 'number' ? raw(el('a:spcBef', null, raw(voidEl('a:spcPct', { val: spcBefPct })))) : null,
			raw(bullet),
			raw(themeFontDefRPr(font, { sz, kern: 1200 })),
		])

	const title = defaultLevel(1, { algn: 'ctr' }, voidEl('a:buNone'), 4400, 'mj', 0)

	const bodyLevels = MASTER_BODY_DEFAULTS.map((base, i) => {
		const bu = base.bu as { char: string; font: string }
		return defaultLevel(
			i + 1,
			{ marL: base.marL, indent: base.indent, algn: base.algn },
			voidEl('a:buFont', { typeface: bu.font, pitchFamily: 34, charset: 0 }) + voidEl('a:buChar', { char: bu.char }),
			base.sz,
			'mn',
			base.spcBefPct
		)
	})

	const otherLevels = MASTER_OTHER_DEFAULTS.map((base, i) =>
		defaultLevel(i + 1, { marL: base.marL, algn: base.algn }, '', base.sz, 'mn')
	)

	// NOTE: the source template's own indentation reaches the file here — each style block's
	// children are preceded by two spaces and its own closing tag by one (see `fmt` below) —
	// same class of quirk as `genXmlPlaceholder` in text-body.ts. Preserved, not reformatted.
	return el(
		'p:txStyles',
		null,
		[
			raw(el('p:titleStyle', null, raw(title), { childPrefix: '  ', closePrefix: ' ' })),
			raw(el('p:bodyStyle', null, bodyLevels.map(raw), { childPrefix: '  ', closePrefix: ' ' })),
			raw(
				el(
					'p:otherStyle',
					null,
					[raw(el('a:defPPr', null, raw(voidEl('a:defRPr', { lang: 'en-US' })))), ...otherLevels.map(raw)],
					{ childPrefix: '  ', closePrefix: ' ' }
				)
			),
		],
		{ childPrefix: ' ' }
	)
}

/**
 * Creates Slide Master 1 (`ppt/slideMasters/slideMaster1.xml`)
 * @param slide - slide object that represents master slide layout
 * @param layouts - slide layouts
 * @param renderers - which renderer emits each shape family, passed through to the shape walk
 * @return XML
 */
export function makeXmlMaster(
	slide: SlideMasterInternal,
	layouts: SlideLayoutInternal[],
	renderers: RendererTable
): string {
	// The layouts are the first of the master's default relationships (`makeXmlMasterRel`), so their
	// ids start where that part numbers default relationships from.
	const layoutRidStart = defaultRelIdStart(slide)
	const layoutDefs = layouts
		.map((_layoutDef, idx) =>
			voidEl('p:sldLayoutId', { id: LAYOUT_IDX_SERIES_BASE + idx, 'r:id': `rId${layoutRidStart + idx}` })
		)
		.join('')

	const clrMap = voidEl('p:clrMap', DEFAULT_COLOR_MAP)

	// CT_HeaderFooter/@sldNum defaults to true (ECMA-376). When a slide-number placeholder is
	// defined on the master we must NOT disable it here, otherwise slides that PowerPoint inserts
	// from this master inherit sldNum="0" and the master slide number disappears.
	const hf = voidEl('p:hf', { sldNum: slide._slideNumberProps ? null : 0, hdr: 0, ftr: 0, dt: 0 })

	const txStyles = slide._txStyles ? makeXmlMasterTxStyles(slide._txStyles) : makeXmlMasterDefaultTxStyles()

	return (
		XML_DECL +
		CRLF +
		el(
			'p:sldMaster',
			{
				...PML_ROOT_NS,
			},
			[
				raw(slideObjectToXml(slide, renderers)),
				raw(clrMap),
				raw(el('p:sldLayoutIdLst', null, raw(layoutDefs))),
				raw(hf),
				raw(txStyles),
			]
		)
	)
}

/**
 * Creates `ppt/slideMasters/_rels/slideMaster1.xml.rels`
 * @param {SlideMasterInternal} masterSlide - the deck's slide master
 * @param {SlideLayoutInternal[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
export function makeXmlMasterRel(masterSlide: SlideMasterInternal, slideLayouts: SlideLayoutInternal[]): string {
	const defaultRels = slideLayouts.map((_layoutDef, idx) => ({
		target: targetFromPptSubpart(slideLayoutPath(idx + 1)),
		type: SLIDE_LAYOUT_REL,
	}))
	defaultRels.push({
		target: '../theme/theme1.xml',
		type: THEME_REL,
	})

	return slideObjectRelationsToXml(masterSlide, defaultRels, SLIDE_MASTER_PATH)
}
