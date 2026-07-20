/**
 * Presentation theme types — the `<a:clrScheme>` overrides and font/theme selection.
 *
 * Re-exported by `../core-interfaces.js`, which is the import site for the rest of `src/`.
 */
import type { HexColor } from './core.js'

/**
 * Theme color scheme overrides. Each slot maps to one `<a:clrScheme>` entry in `theme1.xml`;
 * any slot left unset keeps the default Office value. Provide 6-digit hex (no `#`).
 *
 * Slot names use the OOXML scheme names. The PowerPoint UI labels them as:
 * `dk1`=Text/Dark 1, `lt1`=Background/Light 1, `dk2`=Text/Dark 2, `lt2`=Background/Light 2,
 * `accent1`-`accent6`=Accent 1-6, `hlink`=Hyperlink, `folHlink`=Followed Hyperlink.
 * @example { accent1: 'C00000', dk2: '1F3864', hlink: '0070C0' }
 */
export interface ThemeColorScheme {
	/** Text/Dark 1 (default Office: black via `windowText`) */
	dk1?: HexColor
	/** Background/Light 1 (default Office: white via `window`) */
	lt1?: HexColor
	/** Text/Dark 2 (default Office: `44546A`) */
	dk2?: HexColor
	/** Background/Light 2 (default Office: `E7E6E6`) */
	lt2?: HexColor
	/** Accent 1 (default Office: `4472C4`) */
	accent1?: HexColor
	/** Accent 2 (default Office: `ED7D31`) */
	accent2?: HexColor
	/** Accent 3 (default Office: `A5A5A5`) */
	accent3?: HexColor
	/** Accent 4 (default Office: `FFC000`) */
	accent4?: HexColor
	/** Accent 5 (default Office: `5B9BD5`) */
	accent5?: HexColor
	/** Accent 6 (default Office: `70AD47`) */
	accent6?: HexColor
	/** Hyperlink (default Office: `0563C1`) */
	hlink?: HexColor
	/** Followed Hyperlink (default Office: `954F72`) */
	folHlink?: HexColor
}
export interface ThemeProps {
	/**
	 * Headings font face name
	 * @example 'Arial Narrow'
	 * @default 'Calibri Light'
	 */
	headFontFace?: string
	/**
	 * Body font face name
	 * @example 'Arial'
	 * @default 'Calibri'
	 */
	bodyFontFace?: string
	/**
	 * Headings East Asian font face — theme `<a:ea>` slot of the major font.
	 * Used for CJK (Chinese/Japanese/Korean) runs that fall back to the theme font.
	 * - leave unset to keep PowerPoint's empty default (auto-resolved per script)
	 * @example 'Yu Gothic'
	 */
	headFontFaceEA?: string
	/**
	 * Body East Asian font face — theme `<a:ea>` slot of the minor font.
	 * Used for CJK (Chinese/Japanese/Korean) runs that fall back to the theme font.
	 * - leave unset to keep PowerPoint's empty default (auto-resolved per script)
	 * @example 'Yu Gothic'
	 */
	bodyFontFaceEA?: string
	/**
	 * Headings complex-script font face — theme `<a:cs>` slot of the major font.
	 * Used for complex scripts such as Arabic, Hebrew, Thai, and Devanagari.
	 * - leave unset to keep PowerPoint's empty default (auto-resolved per script)
	 * @example 'Arial'
	 */
	headFontFaceCS?: string
	/**
	 * Body complex-script font face — theme `<a:cs>` slot of the minor font.
	 * Used for complex scripts such as Arabic, Hebrew, Thai, and Devanagari.
	 * - leave unset to keep PowerPoint's empty default (auto-resolved per script)
	 * @example 'Arial'
	 */
	bodyFontFaceCS?: string
	/**
	 * Theme color scheme overrides written to `ppt/theme/theme1.xml`.
	 * - any unset slot keeps its default Office value
	 * - references such as `pptx.SchemeColor.accent1` resolve against these values
	 * @example { accent1: 'C00000', accent2: '00B050', hlink: '0070C0' }
	 */
	colorScheme?: ThemeColorScheme
}
