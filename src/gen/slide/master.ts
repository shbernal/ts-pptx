/**
 * ts-pptx: slide-master parts
 *
 * The slide master (`slideMaster1.xml`) and its `.rels`, plus the master text
 * styles (`<p:txStyles>`): per-level default values mirroring the built-in
 * Office master, with any `MasterTextStyleProps` overrides layered on top.
 */

import { CRLF, LAYOUT_IDX_SERIES_BASE, XML_DECL } from '../../constants-internal.js'
import type { MasterBulletProps, MasterTextStyleLevel, MasterTextStyleProps } from '../../types/index.js'
import type { PresSlideInternal, SlideLayoutInternal } from '../../types/internal.js'
import { createColorElement } from '../drawingml/color.js'
import { inch2Emu } from '../../units-internal.js'
import { HUNDREDTHS_PER_POINT, ptToHundredths } from '../../units.js'
import { warn } from '../../diagnostics.js'
import { el, raw, voidEl, type XmlAttrs } from '../oxml/el.js'
import { slideObjectRelationsToXml, slideObjectToXml } from './object.js'
import { PML_ROOT_NS } from '../../ooxml/namespaces.js'

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
const MASTER_OTHER_DEFAULTS: MasterLevelDefault[] = [
	0, 457200, 914400, 1371600, 1828800, 2286000, 2743200, 3200400, 3657600,
].map((marL) => ({ marL, algn: 'l', sz: 1800, font: 'mn' as const }))

function masterAlignAttr(align: MasterTextStyleLevel['align']): string {
	switch (align) {
		case 'left':
			return 'l'
		case 'right':
			return 'r'
		case 'center':
			return 'ctr'
		case 'justify':
			return 'just'
		case undefined:
			// The only unmatched member: no `align` on the level. `''` omits `@algn` entirely so the
			// level inherits from the theme, which is NOT the same as pinning it to `l`.
			return ''
	}
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
			const type = bulletOverride.numberType || 'arabicPeriod'
			const startAt = typeof bulletOverride.numberStartAt === 'number' ? Math.round(bulletOverride.numberStartAt) : null
			return font + voidEl('a:buAutoNum', { type, startAt })
		}
		// character bullet (default). NOTE: `char` is a pre-escaped numeric char ref (e.g.
		// `&#x25AA;` from `characterCode`) — the same quirk `<a:buChar>` carries in text-run.ts.
		// The builder would double-escape the `&`, so this one attribute stays a raw template;
		// see text-run.ts's `<a:buChar>` note for the fuller rationale.
		const char = bulletOverride.characterCode ? `&#x${bulletOverride.characterCode};` : '•'
		const buFont = bulletOverride.fontFace
			? font
			: voidEl('a:buFont', { typeface: 'Arial', pitchFamily: 34, charset: 0 })
		return `${buFont}<a:buChar char="${char}"/>`
	}
	// No override (undefined / true): keep the level's default bullet
	if (base === 'none') return voidEl('a:buNone')
	if (base && typeof base === 'object')
		return voidEl('a:buFont', { typeface: base.font, pitchFamily: 34, charset: 0 }) + `<a:buChar char="${base.char}"/>`
	return '' // otherStyle: no bullet element by default
}

/** Serialize one `<a:lvlNpPr>` from its default, layering an optional caller override. */
function masterLevelXml(levelNum: number, base: MasterLevelDefault, levelOverride: MasterTextStyleLevel = {}): string {
	const marL =
		typeof levelOverride.marginLeft === 'number' && !isNaN(levelOverride.marginLeft)
			? inch2Emu(levelOverride.marginLeft)
			: base.marL
	const indentEmu =
		typeof levelOverride.indent === 'number' && !isNaN(levelOverride.indent)
			? inch2Emu(levelOverride.indent)
			: base.indent
	const algn = (levelOverride.align && masterAlignAttr(levelOverride.align)) || base.algn

	let sz = base.sz
	if (typeof levelOverride.fontSize === 'number') {
		if (isNaN(levelOverride.fontSize) || levelOverride.fontSize <= 0)
			warn(
				'master/invalid-text-style-font-size',
				`master textStyles fontSize "${levelOverride.fontSize}" is invalid; keeping default ${base.sz / HUNDREDTHS_PER_POINT}pt.`
			)
		else sz = ptToHundredths(levelOverride.fontSize)
	}
	const colorXml = levelOverride.color ? createColorElement(levelOverride.color) : voidEl('a:schemeClr', { val: 'tx1' })
	const latinXml = levelOverride.fontFace
		? voidEl('a:latin', { typeface: levelOverride.fontFace })
		: voidEl('a:latin', { typeface: `+${base.font}-lt` })

	return el(
		`a:lvl${levelNum}pPr`,
		{ marL, indent: indentEmu, algn, defTabSz: 914400, rtl: 0, eaLnBrk: 1, latinLnBrk: 0, hangingPunct: 1 },
		[
			typeof base.spcBefPct === 'number'
				? raw(el('a:spcBef', null, raw(voidEl('a:spcPct', { val: base.spcBefPct }))))
				: null,
			raw(masterBulletXml(levelOverride.bullet, base.bu)),
			raw(
				el('a:defRPr', { sz, b: levelOverride.bold ? '1' : null, i: levelOverride.italic ? '1' : null, kern: 1200 }, [
					raw(el('a:solidFill', null, raw(colorXml))),
					raw(latinXml),
					raw(voidEl('a:ea', { typeface: `+${base.font}-ea` })),
					raw(voidEl('a:cs', { typeface: `+${base.font}-cs` })),
				])
			),
		]
	)
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
		el(`a:lvl${n}pPr`, { ...attrs, defTabSz: 914400, rtl: 0, eaLnBrk: 1, latinLnBrk: 0, hangingPunct: 1 }, [
			typeof spcBefPct === 'number' ? raw(el('a:spcBef', null, raw(voidEl('a:spcPct', { val: spcBefPct })))) : null,
			raw(bullet),
			raw(
				el('a:defRPr', { sz, kern: 1200 }, [
					raw(el('a:solidFill', null, raw(voidEl('a:schemeClr', { val: 'tx1' })))),
					raw(voidEl('a:latin', { typeface: `+${font}-lt` })),
					raw(voidEl('a:ea', { typeface: `+${font}-ea` })),
					raw(voidEl('a:cs', { typeface: `+${font}-cs` })),
				])
			),
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
 * @param {PresSlideInternal} slide - slide object that represents master slide layout
 * @param {SlideLayoutInternal[]} layouts - slide layouts
 * @return {string} XML
 */
export function makeXmlMaster(slide: PresSlideInternal, layouts: SlideLayoutInternal[]): string {
	// NOTE: Pass layouts as static rels because they are not referenced any time
	const layoutDefs = layouts
		.map((_layoutDef, idx) =>
			voidEl('p:sldLayoutId', { id: LAYOUT_IDX_SERIES_BASE + idx, 'r:id': `rId${slide._rels.length + idx + 1}` })
		)
		.join('')

	const clrMap = voidEl('p:clrMap', {
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
	})

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
				raw(slideObjectToXml(slide)),
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
 * @param {PresSlideInternal} masterSlide - Slide object
 * @param {SlideLayoutInternal[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
export function makeXmlMasterRel(masterSlide: PresSlideInternal, slideLayouts: SlideLayoutInternal[]): string {
	const defaultRels = slideLayouts.map((_layoutDef, idx) => ({
		target: `../slideLayouts/slideLayout${idx + 1}.xml`,
		type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
	}))
	defaultRels.push({
		target: '../theme/theme1.xml',
		type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
	})

	return slideObjectRelationsToXml(masterSlide, defaultRels)
}
