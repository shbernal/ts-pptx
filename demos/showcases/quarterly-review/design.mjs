/**
 * Design system for the Kestrel quarterly review deck: palette, type scale, and the
 * five slide masters every slide is built on.
 *
 * The point of keeping this in one file is that the slides in `index.mjs` never name a
 * raw hex value or a magic font size. A deck that hard-codes `color: 'B08D3F'` in forty
 * places is not a design system, it is forty coincidences — and it is exactly what makes
 * a generated deck look generated.
 */
import { ShapeType } from 'pptx-ts'
import { WIDE } from '../lib/layout.mjs'

export const BRAND = {
	ink: '0C1D13',
	forest: '14301F',
	pine: '3E6B4A',
	ash: '6F7269',
	haze: 'D9DED1',
	bone: 'F3F1EA',
	white: 'FFFFFF',
	brass: 'B08D3F',
	// `brass` darkened until small caps hold up on `bone`. The fills and rules use the
	// lighter one; anything set as *text* on a light ground uses this.
	brassInk: '7E6224',
	clay: 'C4784A',
	rust: 'A75A3C',
	lichen: '8FA98A',
	fern: '6E9E72',
}

/** Series colours for charts, ordered so adjacent series stay legible side by side. */
export const CHART_COLORS = [BRAND.brass, BRAND.forest, BRAND.clay, BRAND.lichen, BRAND.rust, BRAND.fern]

export const FONT = {
	head: 'Segoe UI Semibold',
	body: 'Segoe UI',
	mono: 'Consolas',
}

/** Type scale. Sizes are points, matching PowerPoint's own units. */
export const TYPE = {
	display: 44,
	title: 32,
	section: 40,
	heading: 22,
	subhead: 16,
	body: 14,
	caption: 11,
	metric: 40,
}

export const MASTER = {
	title: 'KESTREL_TITLE',
	section: 'KESTREL_SECTION',
	content: 'KESTREL_CONTENT',
	data: 'KESTREL_DATA',
	closing: 'KESTREL_CLOSING',
}

const FOOTER_TEXT = 'Kestrel Analytics  ·  Q3 FY26 Business Review  ·  Internal'
const FOOTER_Y = 6.95
const FOOTER_H = 0.55

/**
 * The footer band shared by the two body masters: a forest bar, the deck's running
 * title on the left, and room on the right for the slide number the master places.
 */
function footerObjects() {
	return [
		{ rect: { x: 0, y: FOOTER_Y, w: '100%', h: FOOTER_H, fill: { color: BRAND.forest } } },
		{
			text: {
				text: FOOTER_TEXT,
				options: {
					x: 0.75,
					y: FOOTER_Y,
					w: 8,
					h: FOOTER_H,
					margin: 0,
					valign: 'middle',
					fontFace: FONT.body,
					fontSize: 9,
					color: BRAND.haze,
					charSpacing: 0.6,
				},
			},
		},
	]
}

const SLIDE_NUMBER = {
	x: WIDE.w - 1.15,
	y: FOOTER_Y,
	w: 0.5,
	h: FOOTER_H,
	align: 'right',
	valign: 'middle',
	fontFace: FONT.body,
	fontSize: 10,
	color: BRAND.white,
	bold: true,
}

/**
 * Apply the Kestrel theme and register every master.
 *
 * `pptx.theme.colorScheme` matters more than it looks: it rewrites `theme1.xml`, so a
 * chart series left on `accent1`, a hyperlink, and a table style region all pick up the
 * brand palette without being told about it individually.
 */
export function applyDesign(pptx) {
	pptx.layout = 'LAYOUT_WIDE'

	pptx.theme = {
		headFontFace: FONT.head,
		bodyFontFace: FONT.body,
		colorScheme: {
			dk1: BRAND.ink,
			lt1: BRAND.white,
			dk2: BRAND.forest,
			lt2: BRAND.bone,
			accent1: BRAND.brass,
			accent2: BRAND.forest,
			accent3: BRAND.clay,
			accent4: BRAND.lichen,
			accent5: BRAND.rust,
			accent6: BRAND.fern,
			hlink: BRAND.brassInk,
			folHlink: BRAND.ash,
		},
	}

	// TITLE — a native linear gradient rather than a background image, so the cover
	// scales to any layout and adds nothing to the package size.
	pptx.defineSlideMaster({
		title: MASTER.title,
		background: {
			type: 'gradient',
			gradient: {
				kind: 'linear',
				angle: 135,
				stops: [
					{ position: 0, color: BRAND.ink },
					{ position: 55, color: BRAND.forest },
					{ position: 100, color: BRAND.pine },
				],
			},
		},
		objects: [
			{ rect: { x: 0, y: 0, w: 0.22, h: '100%', fill: { color: BRAND.brass } } },
			{ rect: { x: 0.95, y: 4.62, w: 2.4, h: 0.06, fill: { color: BRAND.brass } } },
		],
	})

	// SECTION — flat ink, with an oversized ghosted numeral slot on the right.
	pptx.defineSlideMaster({
		title: MASTER.section,
		background: { color: BRAND.ink },
		objects: [
			{ rect: { x: 0, y: 0, w: 0.22, h: '100%', fill: { color: BRAND.clay } } },
			{ rect: { x: 0.95, y: 4.35, w: 1.5, h: 0.06, fill: { color: BRAND.clay } } },
		],
	})

	// CONTENT — the workhorse. Paper ground, hairline rule under the title area.
	pptx.defineSlideMaster({
		title: MASTER.content,
		background: { color: BRAND.bone },
		objects: [
			{ rect: { x: 0, y: 0, w: '100%', h: 0.1, fill: { color: BRAND.brass } } },
			{ rect: { x: 0.75, y: 1.28, w: WIDE.w - 1.5, h: 0.015, fill: { color: BRAND.haze } } },
			...footerObjects(),
		],
		slideNumber: SLIDE_NUMBER,
	})

	// DATA — same chrome, white ground. Charts and tables read better without the
	// paper tint behind them, and PowerPoint's own chart area default is white.
	pptx.defineSlideMaster({
		title: MASTER.data,
		background: { color: BRAND.white },
		objects: [
			{ rect: { x: 0, y: 0, w: '100%', h: 0.1, fill: { color: BRAND.clay } } },
			{ rect: { x: 0.75, y: 1.28, w: WIDE.w - 1.5, h: 0.015, fill: { color: BRAND.haze } } },
			...footerObjects(),
		],
		slideNumber: SLIDE_NUMBER,
	})

	// CLOSING — the title gradient mirrored, so the deck opens and shuts on the same note.
	pptx.defineSlideMaster({
		title: MASTER.closing,
		background: {
			type: 'gradient',
			gradient: {
				kind: 'linear',
				angle: 315,
				stops: [
					{ position: 0, color: BRAND.pine },
					{ position: 45, color: BRAND.forest },
					{ position: 100, color: BRAND.ink },
				],
			},
		},
		objects: [{ rect: { x: 0, y: 0, w: '100%', h: 0.1, fill: { color: BRAND.brass } } }],
	})
}

/** Slide title + optional kicker, positioned against the master's hairline rule. */
export function slideTitle(slide, title, kicker) {
	if (kicker) {
		slide.addText(kicker.toUpperCase(), {
			x: 0.75,
			y: 0.34,
			w: 8,
			h: 0.3,
			margin: 0,
			fontFace: FONT.body,
			fontSize: TYPE.caption,
			color: BRAND.brassInk,
			bold: true,
			charSpacing: 1.6,
		})
	}
	slide.addText(title, {
		x: 0.75,
		y: kicker ? 0.62 : 0.5,
		w: WIDE.w - 1.5,
		h: 0.6,
		margin: 0,
		valign: 'middle',
		fontFace: FONT.head,
		fontSize: TYPE.title,
		color: BRAND.ink,
	})
}

/**
 * The children of one KPI card: panel, accent cap, value, label, and a tinted delta chip.
 *
 * Returns the child array rather than a built object so the caller decides the grouping —
 * `slide.addGroup(kpiCard(...))` makes each card a single selectable object in PowerPoint,
 * which is what someone editing the deck afterwards actually wants.
 */
export function kpiCard({ x, y, w, h, label, value, delta, positive }) {
	const accent = positive ? BRAND.brass : BRAND.rust
	return [
		{
			roundRect: {
				x,
				y,
				w,
				h,
				rectRadius: 0.06,
				fill: { color: BRAND.white },
				line: { color: BRAND.haze, width: 1 },
			},
		},
		{ rect: { x: x + 0.06, y, w: w - 0.12, h: 0.07, fill: { color: accent } } },
		{
			text: {
				text: value,
				options: {
					x: x + 0.24,
					y: y + 0.28,
					w: w - 0.48,
					h: 0.85,
					margin: 0,
					valign: 'middle',
					fontFace: FONT.head,
					fontSize: TYPE.metric,
					color: BRAND.ink,
				},
			},
		},
		{
			text: {
				text: label.toUpperCase(),
				options: {
					x: x + 0.24,
					y: y + 1.14,
					w: w - 0.48,
					h: 0.3,
					margin: 0,
					fontFace: FONT.body,
					fontSize: 10,
					color: BRAND.ash,
					charSpacing: 1.2,
				},
			},
		},
		{
			text: {
				text: delta,
				options: {
					x: x + 0.24,
					y: y + 1.52,
					w: 1.25,
					h: 0.34,
					margin: 0,
					align: 'center',
					valign: 'middle',
					shape: ShapeType.roundRect,
					rectRadius: 0.16,
					fill: { color: accent, transparency: 86 },
					fontFace: FONT.head,
					fontSize: 11,
					color: accent,
				},
			},
		},
	]
}
