// Deck definitions shared by the browser harness and the Node side of the same test.
//
// Every scenario here exists to reach one of the three `RuntimeAdapter` loaders that
// `quarterly-review` never touches, because it draws every asset rather than loading one:
// `loadMedia`, `createSvgPngPreview`, `loadFontData` (src/runtime/browser.ts).
//
// The point of defining them once is that most of these decks get built *twice* — once in
// Chromium against `dist/browser.js` and once in Node against `dist/node.js` — and the
// assertion is that the two packages agree. If the deck were written out in both places
// a divergence in the fixture would read as a divergence in the runtime.
//
// `assets` is whatever the runtime under test consumes for a media/font source: absolute
// file paths on Node, URLs in the browser. That difference is the whole subject of the
// test, so it is a parameter and nothing else here is.

/** Silkscreen is a wide pixel font, so this overflows a 3in box and forces a real shrink. */
const OVERFLOW = 'The quick brown fox jumps over the lazy dog. '.repeat(6).trim()

/**
 * Every image deck pins its alt text, and has to.
 *
 * `addImage` defaults `descr` to the image *source* — which is a filesystem path on Node
 * and a URL in the browser, because that is the input the two runtimes take. So the
 * default makes the emitted XML differ for a reason that has nothing to do with the
 * loaders: the library faithfully recorded two different strings a caller handed it.
 * Pinning it removes the only confound and leaves the comparison about the bytes.
 */
const ALT_TEXT = 'harness fixture'

/**
 * @typedef {object} HarnessAssets
 * @property {string} png a real raster image
 * @property {string} svg a real SVG with an intrinsic size
 * @property {string} font a real TTF (Silkscreen, OFL)
 * @property {string} missingPng a path/URL that does not exist
 * @property {string} missingFont a path/URL that does not exist
 * @property {string} brokenSvg a `.svg` that fetches fine and does not decode
 * @property {string} zeroSizeSvg a `.svg` whose intrinsic width and height are both 0
 */

/**
 * Deck builders, keyed by the name the harness dispatches on. Each takes a fresh
 * presentation and the asset table, and leaves the deck ready to write.
 *
 * @type {Record<string, (pres: any, assets: HarnessAssets) => Promise<void>>}
 */
export const DECKS = {
	/**
	 * `loadMedia`, happy path. A raster image by path/URL is the plain case: Node reads
	 * bytes off disk and base64s them, the browser fetches a Blob and runs it through a
	 * `FileReader` into a data URI. Two different string shapes for the same bytes — this
	 * deck is what proves they land in the package as the same bytes anyway.
	 *
	 * Width and height are left off deliberately: a path image defers sizing to export
	 * time and derives it from the loaded payload, so `_szAuto` only resolves if the
	 * runtime's data URI is readable by the image sizer. Pinning w/h here would skip that.
	 */
	async raster(pres, assets) {
		const slide = pres.addSlide()
		slide.addImage({ path: assets.png, x: 1, y: 1, altText: ALT_TEXT })
	},

	/**
	 * `loadMedia` + `createSvgPngPreview`. An SVG consumes two rels: the SVG itself and a
	 * PNG fallback for renderers that will not draw SVG. Node has no rasterizer, so its
	 * `createSvgPngPreview` writes a fixed placeholder; the browser draws the SVG to a
	 * `<canvas>` and reads back a real PNG. This is the one deck the two runtimes are
	 * *expected* to disagree on, and the browser lane asserts the shape of the disagreement.
	 */
	async svg(pres, assets) {
		const slide = pres.addSlide()
		slide.addImage({ path: assets.svg, x: 1, y: 1, w: 2, h: 2, altText: ALT_TEXT })
	},

	/**
	 * `loadFontData`, both of its call sites in one deck: `registerFontMetrics` (bytes →
	 * opentype.js → a baked `fontScale`) and `embedFont` (bytes → an `/ppt/fonts/` part).
	 * The first proves the bytes parsed, the second proves they arrived intact, and the
	 * cross-runtime comparison proves a URL fetch and an `fs.readFile` produce the same
	 * ones.
	 */
	async fonts(pres, assets) {
		await pres.registerFontMetrics('Silkscreen', assets.font)
		await pres.embedFont({ path: assets.font, typeface: 'Silkscreen' })
		const slide = pres.addSlide()
		slide.addText(OVERFLOW, { x: 1, y: 1, w: 3, h: 1, fontFace: 'Silkscreen', fontSize: 18, fit: 'shrink' })
	},

	/** `loadMedia`'s `!response.ok` arm — a 404 under a served prefix, not a stubbed fetch. */
	async missingImage(pres, assets) {
		const slide = pres.addSlide()
		slide.addImage({ path: assets.missingPng, x: 1, y: 1, w: 1, h: 1 })
	},

	/** `loadFontData`'s `!response.ok` arm. Throws from `registerFontMetrics`, not at export. */
	async missingFont(pres, assets) {
		await pres.registerFontMetrics('Silkscreen', assets.missingFont)
	},

	/**
	 * `createSvgPngPreview`'s `image.onerror` arm: the fetch succeeds, so `loadMedia`
	 * hands back a well-formed `image/svg+xml` data URI, and the decode is what fails.
	 * Browser-only in meaning — Node's `createSvgPngPreview` cannot fail, so the same deck
	 * exports cleanly there with a placeholder in the PNG rel.
	 */
	async brokenSvg(pres, assets) {
		const slide = pres.addSlide()
		slide.addImage({ path: assets.brokenSvg, x: 1, y: 1, w: 1, h: 1 })
	},

	/**
	 * `createSvgPngPreview`'s `image.width + image.height === 0` guard. A 0×0 SVG decodes
	 * without error and then has nothing to draw; without the guard `canvas.toDataURL`
	 * would return a valid-looking data URI for an empty image and the deck would ship a
	 * blank fallback instead of failing. Browser-only, as above.
	 */
	async zeroSizeSvg(pres, assets) {
		const slide = pres.addSlide()
		slide.addImage({ path: assets.zeroSizeSvg, x: 1, y: 1, w: 1, h: 1 })
	},
}

/**
 * Build one named deck and return the package as base64.
 *
 * base64 rather than bytes because this same call is made through `page.evaluate`, whose
 * return value has to survive structured serialization — and because it keeps the Node
 * and browser call sites character-for-character identical.
 *
 * @param {any} pres a fresh presentation from the entry under test
 * @param {string} name a key of {@link DECKS}
 * @param {HarnessAssets} assets
 * @returns {Promise<string>}
 */
export async function buildDeckBase64(pres, name, assets) {
	const build = DECKS[name]
	if (!build) throw new Error(`unknown harness deck: ${name}`)
	await build(pres, assets)
	return /** @type {string} */ (await pres.write({ outputType: 'base64' }))
}
