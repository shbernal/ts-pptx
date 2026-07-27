/**
 * Design system for the "Field Notes" deck: an editorial, image-led counterpart to the
 * quarterly review.
 *
 * Where the review deck is a grid of panels on a light ground, this one is dark, and every
 * spread is built on a photograph. That difference is the point of having two showcases —
 * the same library has to be able to produce both, and the techniques barely overlap.
 */
import { WIDE } from "../lib/layout.mjs";

export const BRAND = {
	ink: "14100E",
	charcoal: "241D19",
	stone: "4A403A",
	ash: "8C7F76",
	sand: "D9CFC4",
	bone: "F4EFE9",
	white: "FFFFFF",
	amber: "D98A2B",
	rust: "A8452C",
	sage: "6E8B6A",
};

export const FONT = {
	head: "Georgia",
	body: "Segoe UI",
	mono: "Consolas",
};

export const TYPE = {
	display: 54,
	title: 34,
	standfirst: 17,
	heading: 21,
	body: 13,
	caption: 10,
	quote: 24,
};

export const MASTER = {
	cover: "FIELD_COVER",
	plate: "FIELD_PLATE",
	editorial: "FIELD_EDITORIAL",
	colophon: "FIELD_COLOPHON",
};

/**
 * A dark-to-transparent scrim.
 *
 * Full-bleed photographs and white type do not reliably coexist: any given photo has a
 * bright patch somewhere, and that patch is where the headline becomes unreadable. Laying
 * a gradient that fades from opaque ink to fully transparent over the image guarantees
 * contrast at the type end regardless of what the photograph is doing underneath — the
 * standard editorial fix, and one native `a:gradFill` handles without a second bitmap.
 */
export function scrim({ x, y, w, h, angle = 0, color = BRAND.ink, from = 12, to = 100 }) {
	return {
		x,
		y,
		w,
		h,
		fill: {
			type: "gradient",
			gradient: {
				kind: "linear",
				angle,
				stops: [
					{ position: 0, color, transparency: from },
					{ position: 65, color, transparency: 72 },
					{ position: 100, color, transparency: to },
				],
			},
		},
		line: { type: "none" },
	};
}

export function applyDesign(pptx) {
	pptx.layout = "LAYOUT_WIDE";

	pptx.theme = {
		headFontFace: FONT.head,
		bodyFontFace: FONT.body,
		colorScheme: {
			dk1: BRAND.ink,
			lt1: BRAND.bone,
			dk2: BRAND.charcoal,
			lt2: BRAND.sand,
			accent1: BRAND.amber,
			accent2: BRAND.rust,
			accent3: BRAND.sage,
			accent4: BRAND.stone,
			accent5: BRAND.ash,
			accent6: BRAND.sand,
			hlink: BRAND.amber,
			folHlink: BRAND.ash,
		},
	};

	// COVER — the photograph is added per-slide, so the master carries only the ink ground
	// it sits on and the rule the wordmark aligns to.
	pptx.defineSlideMaster({
		title: MASTER.cover,
		background: { color: BRAND.ink },
	});

	// PLATE — a full-bleed image slide. Ink ground, and a thin amber rule bottom-left that
	// every plate's caption hangs from, so the captions line up across the deck.
	pptx.defineSlideMaster({
		title: MASTER.plate,
		background: { color: BRAND.ink },
		objects: [{ rect: { x: 0.85, y: 6.62, w: 0.75, h: 0.035, fill: { color: BRAND.amber } } }],
	});

	// EDITORIAL — text-forward pages on a warm dark ground, with a running foot.
	pptx.defineSlideMaster({
		title: MASTER.editorial,
		background: { color: BRAND.charcoal },
		objects: [
			{
				text: {
					text: "FIELD NOTES  ·  FOUR CITIES AFTER DARK",
					options: {
						x: 0.85,
						y: 6.85,
						w: 8,
						h: 0.3,
						margin: 0,
						valign: "middle",
						fontFace: FONT.body,
						fontSize: 8,
						color: BRAND.stone,
						charSpacing: 2.4,
					},
				},
			},
		],
		slideNumber: {
			x: WIDE.w - 1.35,
			y: 6.85,
			w: 0.5,
			h: 0.3,
			align: "right",
			valign: "middle",
			fontFace: FONT.body,
			fontSize: 9,
			color: BRAND.ash,
		},
	});

	// COLOPHON — a radial gradient, so the back cover reads as a different kind of page
	// from the linear-gradient front. Same API, visibly different result.
	pptx.defineSlideMaster({
		title: MASTER.colophon,
		background: {
			type: "gradient",
			gradient: {
				kind: "radial",
				center: { x: 32, y: 40 },
				stops: [
					{ position: 0, color: BRAND.stone },
					{ position: 55, color: BRAND.charcoal },
					{ position: 100, color: BRAND.ink },
				],
			},
		},
	});
}

/** The deck's wordmark, used on the cover and the colophon. */
export function wordmark(slide, { x, y, color = BRAND.bone }) {
	slide.addText("FIELD NOTES", {
		x,
		y,
		w: 5,
		h: 0.3,
		margin: 0,
		fontFace: FONT.body,
		fontSize: 11,
		color,
		bold: true,
		charSpacing: 5,
	});
}
