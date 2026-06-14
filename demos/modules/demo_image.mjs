/**
 * NAME: demo_images.mjs
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by maintained demos (./demos/node, ./demos/vite-demo).
 * VER.: 3.12.0
 * BLD.: 20230319
 */

/**
 * NOTES:
 * - Images can be pre-encoded into base64, so they do not have to be on the webserver etc. (saves generation time and resources!)
 * - This also has the benefit of being able to be any type (path:images can only be exported as PNG)
 * - Image source: either `data` or `path` is required
 */

import {
	IMAGE_PATHS,
	BASE_TABLE_OPTS,
	BASE_TEXT_OPTS_L,
	BASE_TEXT_OPTS_R,
	BASE_CODE_OPTS,
	BKGD_LTGRAY,
	BKGD_LRGRAY,
	COLOR_BLUE,
	CODE_STYLE,
	TITLE_STYLE,
} from "./enums.mjs";
import { FEDIVERSE_TREE, HYPERLINK_SVG, KRITA_SPLASHSCREEN, SVG_MASTODON_LOGO_BASE64, UNITE_PNG } from "./media.mjs";

export function genSlides_Image(pptx) {
	pptx.addSection({ title: "Images" });

	genSlide01(pptx);
	genSlide02(pptx);
	genSlide03(pptx);
	genSlide04(pptx);
	genSlide05(pptx);
	genSlide06(pptx);
}

/**
 * SLIDE 1:
 * @param {PptxGenJS} pptx
 */
function genSlide01(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Images" });

	slide.addTable([[{ text: "Image Examples: Image Types", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-images.html");
	slide.slideNumber = { x: "50%", y: "95%", color: COLOR_BLUE };

	// TOP
	{
		// TOP: 1
		slide.addText("Type: GIF (animated)", {
			x: 0.5,
			y: 0.6,
			w: 2.75,
			h: 2.5,
			margin: 4,
			fill: { color: BKGD_LTGRAY },
			fontSize: 12,
			fontFace: "Segoe UI",
			color: COLOR_BLUE,
			valign: "top",
			align: "center",
		});
		slide.addImage({ x: 1.13, y: 1.1, w: 1.5, h: 1.5, path: IMAGE_PATHS.gifAnimTrippy.path, objectName: "animated gif" });
		slide.addText("(legacy PPT only animates in slide show)", {
			x: 0.5,
			y: 2.79,
			w: 2.75,
			h: 0.3,
			align: "center",
			fontSize: 10,
			color: "BF9000",
			fill: { color: "FFFCCC" },
		});

		// TOP: 2
		slide.addText("Type: GIF", {
			x: 3.7,
			y: 0.6,
			w: 2.75,
			h: 2.5,
			margin: 4,
			fill: { color: BKGD_LTGRAY },
			fontSize: 12,
			fontFace: "Segoe UI",
			color: COLOR_BLUE,
			valign: "top",
			align: "center",
		});
		slide.addImage({ x: 4.16, y: 1.1, w: 1.88, h: 1.5, path: IMAGE_PATHS.ccDjGif.path, altText: "this is a gif" });

		// TOP: 3
		slide.addText("Type: PNG (base64)", {
			x: 6.9,
			y: 0.6,
			w: 2.75,
			h: 2.5,
			margin: 4,
			fill: { color: BKGD_LTGRAY },
			fontSize: 12,
			fontFace: "Segoe UI",
			color: COLOR_BLUE,
			valign: "top",
			align: "center",
		});
		slide.addImage({ x: 7.53, y: 1.1, w: 1.5, h: 1.5, data: UNITE_PNG });

		// TOP: 4
		slide.addText("Hyperlink Image", {
			x: 10.1,
			y: 0.6,
			w: 2.75,
			h: 2.5,
			margin: 4,
			fill: { color: BKGD_LTGRAY },
			fontSize: 12,
			fontFace: "Segoe UI",
			color: COLOR_BLUE,
			valign: "top",
			align: "center",
		});
		slide.addImage({
			x: 10.8,
			y: 1.1,
			w: 1.36,
			h: 1.5,
			data: HYPERLINK_SVG,
			hyperlink: { url: "https://github.com/gitbrent/pptxgenjs", tooltip: "Visit Homepage" },
		});
	}

	// BTM
	{
		// BOTTOM-LEFT:
		slide.addText("Type: JPG", {
			x: 0.5,
			y: 3.5,
			w: 3.5,
			h: 3.5,
			margin: 4,
			fill: { color: BKGD_LTGRAY },
			fontSize: 12,
			fontFace: "Segoe UI",
			color: COLOR_BLUE,
			valign: "top",
			align: "center",
		});
		//slide.addImage({ path: IMAGE_PATHS.fediverse_tree.path, x: 0.72, y: 3.81, w: 3.06, h: 3.06 });
		slide.addImage({ data: FEDIVERSE_TREE, x: 0.72, y: 3.81, w: 3.06, h: 3.06 });

		// BOTTOM-CENTER:
		// peace image via: https://www.vecteezy.com/vector-art/242684-peace-vector-design
		slide.addText("Type: PNG", {
			x: 4.93,
			y: 3.5,
			w: 3.5,
			h: 3.5,
			margin: 4,
			fill: { color: BKGD_LTGRAY },
			fontSize: 12,
			fontFace: "Segoe UI",
			color: COLOR_BLUE,
			valign: "top",
			align: "center",
		});
		slide.addImage({ path: IMAGE_PATHS.peace4.path, x: 5.2, y: 3.81, w: 3.0, h: 3.0 });

		// BOTTOM-RIGHT:
		slide.addText("Type: SVG", {
			x: 9.33,
			y: 3.5,
			w: 3.5,
			h: 3.5,
			margin: 4,
			fill: { color: BKGD_LTGRAY },
			fontSize: 12,
			fontFace: "Segoe UI",
			color: COLOR_BLUE,
			valign: "top",
			align: "center",
		});
		slide.addImage({ path: IMAGE_PATHS.pixelfed_logo_svg.path, x: 9.25, y: 3.43, w: 2.33, h: 2.33 }); // TEST: `path`
		slide.addImage({ data: SVG_MASTODON_LOGO_BASE64, x: 10.61, y: 4.77, w: 2.0, h: 2.0, transparency: 50 }); // TEST: `data`
	}

	// TEST: Ensure framework corrects for missing all header
	// (Please **DO NOT** pass base64 data without the header! This is a JUNK TEST!)
	//slide.addImage({ x:5.2, y:2.6, w:0.8, h:0.8, data:'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAAjcAAAI3AGf6F88AAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAANVQTFRF////JLaSIJ+AIKqKKa2FKLCIJq+IJa6HJa6JJa6IJa6IJa2IJa6IJa6IJa6IJa6IJa6IJa6IJq6IKK+JKK+KKrCLLrGNL7KOMrOPNrSRN7WSPLeVQrmYRLmZSrycTr2eUb6gUb+gWsKlY8Wqbsmwb8mwdcy0d8y1e863g9G7hdK8htK9i9TAjNTAjtXBktfEntvKoNzLquDRruHTtePWt+TYv+fcx+rhyOvh0e7m1e/o2fHq4PTu5PXx5vbx7Pj18fr49fv59/z7+Pz7+f38/P79/f7+dNHCUgAAABF0Uk5TAAcIGBktSYSXmMHI2uPy8/XVqDFbAAABB0lEQVQ4y42T13qDMAyFZUKMbebp3mmbrnTvlY60TXn/R+oFGAyYzz1Xx/wylmWJqBLjUkVpGinJGXXliwSVEuG3sBdkaCgLPJMPQnQUDmo+jGFRPKz2WzkQl//wQvQoLPII0KuAiMjP+gMyn4iEFU1eAQCCiCU2fpCfFBVjxG18f35VOk7Swndmt9pKUl2++fG4qL2iqMPXpi8r1SKitDDne/rT8vPbRh2d6oC7n6PCLNx/bsEM0Edc5DdLAHD9tWueF9VJjmdP68DZ77iRkDKuuT19Hx3mx82MpVmo1Yfv+WXrSrxZ6slpiyes77FKif88t7Nh3C3nbFp327sHxz167uHtH/8/eds7gGsUQbkAAAAASUVORK5CYII=' });
	// NEGATIVE-TEST:
	//slide.addImage({ data:'../common/images/brokenImage.gif',  x:0.5, y:0.5, w:1.0, h:1.0 });
}

/**
 * SLIDE 4: Image URLs
 * @param {PptxGenJS} pptx
 */
function genSlide02(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Images" });

	slide.addTable([[{ text: "Image Examples: Image URLs", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-images.html");
	slide.slideNumber = { x: "50%", y: "95%", color: COLOR_BLUE };

	// TOP-LEFT: jpg
	slide.addImage({ path: IMAGE_PATHS.ccLogo.path, x: 0.5, y: 0.6, h: 2.68, w: 3.58 });
	slide.addText([{ text: `path:"${IMAGE_PATHS.ccLogo.path}"` }], { ...BASE_CODE_OPTS, ...{ x: 0.5, y: 3.28, h: 0.7, w: 3.58 }, ...CODE_STYLE });

	// TOP-CENTER: png
	slide.addImage({ path: IMAGE_PATHS.wikimedia2.path, x: 4.6, y: 0.6, h: 2.64, w: 3.45 });
	slide.addText([{ text: `path:"${IMAGE_PATHS.wikimedia2.path}"` }], { ...BASE_CODE_OPTS, ...{ x: 4.6, y: 3.24, h: 0.7, w: 3.45 }, ...CODE_STYLE });

	// TOP-RIGHT: relative-path test
	// NOTE: Node will throw exception when using "/" path
	// FIXME:
	console.log(`${typeof window === "undefined" ? ".." : ""}${IMAGE_PATHS.ccLicenseComp.path}`);
	// WIP: ^^^
	slide.addImage({
		path: `${typeof window === "undefined" ? ".." : ""}${IMAGE_PATHS.ccLicenseComp.path}`,
		x: 8.57,
		y: 0.6,
		h: 2.52,
		w: 4.26,
	});
	slide.addText([{ text: "// Example: local path", options: { breakLine: true } }, { text: `path:"${IMAGE_PATHS.ccLicenseComp.path}"` }], {
		...BASE_CODE_OPTS,
		...{ x: 8.57, y: 3.12, h: 0.82, w: 4.26 },
		...CODE_STYLE,
	});

	// BOTTOM: wide, url-sourced
	slide.addImage({ path: IMAGE_PATHS.sydneyBridge.path, x: 0.5, y: 4.35, h: 1.8, w: 12.33 });
	slide.addText(
		[{ text: '// Example: URL variables, plus more than one ".jpg"', options: { breakLine: true } }, { text: `path:"${IMAGE_PATHS.sydneyBridge.path}"` }],
		{ ...BASE_CODE_OPTS, ...{ x: 0.5, y: 6.15, h: 0.8, w: 12.33 }, ...CODE_STYLE },
	);
}

/**
 * SLIDE 2: Image Sizing
 * @param {PptxGenJS} pptx
 */
function genSlide03(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Images" });

	slide.addTable([[{ text: "Image Examples: Image Sizing/Rounding", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-images.html");
	slide.slideNumber = { x: "50%", y: "95%", w: 1, h: 1, color: COLOR_BLUE };

	// TOP: 1
	slide.addText("Sizing: Orig `w:5, h:2.8`", { x: 0.5, y: 0.6, w: 3.0, h: 0.3, color: COLOR_BLUE });
	slide.addImage({ data: KRITA_SPLASHSCREEN, x: 0.6, y: 1.1, w: 5.0, h: 2.84 });

	// TOP: 2
	slide.addText("Sizing: `contain, w:3`", { x: 0.6, y: 4.25, w: 3.0, h: 0.3, color: COLOR_BLUE });
	slide.addShape(pptx.shapes.RECTANGLE, { x: 0.6, y: 4.65, w: 3, h: 2, fill: { color: BKGD_LRGRAY } });
	slide.addImage({ data: KRITA_SPLASHSCREEN, x: 0.6, y: 4.65, w: 5.0, h: 1.5, sizing: { type: "contain", w: 3, h: 2 } });

	// TOP: 3
	slide.addText("Sizing: `cover, w:3, h:2`", { x: 5.3, y: 4.25, w: 3.0, h: 0.3, color: COLOR_BLUE });
	slide.addShape(pptx.shapes.RECTANGLE, { x: 5.3, y: 4.65, w: 3, h: 2, fill: { color: BKGD_LRGRAY } });
	slide.addImage({ data: KRITA_SPLASHSCREEN, x: 5.3, y: 4.65, w: 3.0, h: 1.5, sizing: { type: "cover", w: 3, h: 2 } });

	// TOP: 4
	slide.addText("Sizing: `crop, w:3, h:2`", { x: 10.0, y: 4.25, w: 3.0, h: 0.3, color: COLOR_BLUE });
	slide.addShape(pptx.shapes.RECTANGLE, { x: 10, y: 4.65, w: 3, h: 1.5, fill: { color: BKGD_LRGRAY } });
	slide.addImage({ data: KRITA_SPLASHSCREEN, x: 10.0, y: 4.65, w: 5.0, h: 1.5, sizing: { type: "crop", w: 3, h: 1.0, x: 0.5, y: 0.25 } });

	// TOP-RIGHT:
	slide.addText("Rounding: `rounding:true`", { x: 7.0, y: 0.6, w: 3.0, h: 0.3, color: COLOR_BLUE });
	slide.addShape(pptx.shapes.RECTANGLE, { x: 7, y: 1.0, w: 6.0, h: 3.0, fill: { color: BKGD_LRGRAY } });
	slide.addImage({ path: IMAGE_PATHS.kritaSquare.path, x: 7.42, y: 1.28, w: 2.5, h: 2.5, rounding: false });
	slide.addImage({ path: IMAGE_PATHS.kritaSquare.path, x: 10.25, y: 1.28, w: 2.5, h: 2.5, rounding: true });
}

/**
 * SLIDE 3: Image Rotation
 * @param {PptxGenJS} pptx
 */
function genSlide04(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Images" });

	slide.addTable([[{ text: "Image Examples: Image Rotation", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-images.html");
	slide.slideNumber = { x: "50%", y: "95%", w: 1, h: 1, color: COLOR_BLUE };

	// EXAMPLES
	slide.addText("`rotate:45` ", { ...{ x: 0.5, y: 0.6, h: 0.4, w: 4.0 }, ...TITLE_STYLE });
	slide.addText("`rotate:180`", { ...{ x: 4.66, y: 0.6, h: 0.4, w: 4.0 }, ...TITLE_STYLE });
	slide.addText("`rotate:315`", { ...{ x: 8.82, y: 0.6, h: 0.4, w: 4.0 }, ...TITLE_STYLE });

	slide.addImage({ path: IMAGE_PATHS.nycSubway.path, x: 0.78, y: 2.46, w: 4.5, h: 3, rotate: 45 });
	slide.addImage({ path: IMAGE_PATHS.nycSubway.path, x: 4.42, y: 2.25, w: 4.5, h: 3, rotate: 180 });
	slide.addImage({ path: IMAGE_PATHS.nycSubway.path, x: 8.25, y: 2.84, w: 4.5, h: 3, rotate: 315 });
}

/**
 * SLIDE 5: Image Shadow
 * @param {PptxGenJS} pptx
 */
function genSlide05(pptx) {
	const slide = pptx.addSlide({ sectionTitle: "Images" });

	slide.addTable([[{ text: "Image Examples: Image Shadows", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-images.html");
	slide.slideNumber = { x: "50%", y: "95%", w: 1, h: 1, color: COLOR_BLUE };

	// EXAMPLES

	// type:none
	const shadow1 = { shadow: { type: "none" } };
	slide.addText("Shadow: `type:none`", { ...{ x: 0.5, y: 0.6, h: 0.4, w: 6.0 }, ...TITLE_STYLE });
	slide.addText([{ text: JSON.stringify(shadow1, "", 2) }], { ...BASE_CODE_OPTS, ...{ x: 0.5, y: 1.0, h: 1.0, w: 6 }, ...CODE_STYLE });
	slide.addImage({ ...{ path: IMAGE_PATHS.nycSubway.path, x: 7.0, y: 0.6, h: 1.4, w: 2.11 }, ...shadow1 });

	// type:inner
	const shadow2 = { shadow: { type: "inner", opacity: 0.5, blur: 20, color: "000000", offset: 20, angle: 320 } };
	slide.addText("Shadow: `type:inner`", { ...{ x: 0.5, y: 2.45, h: 0.4, w: 6.0 }, ...TITLE_STYLE });
	slide.addText([{ text: JSON.stringify(shadow2, "", 2) }], { ...BASE_CODE_OPTS, ...{ x: 0.5, y: 2.85, h: 1.7, w: 6 }, ...CODE_STYLE });
	slide.addImage({ ...{ path: IMAGE_PATHS.nycSubway.path, x: 7.0, y: 2.45, h: 2.1, w: 3.13 }, ...shadow2 });

	// type:outer
	const shadow3 = { shadow: { type: "outer", opacity: 0.35, blur: 20, color: "000000", offset: 20, angle: 320 } };
	slide.addText("Shadow: `type:outer`", { ...{ x: 0.5, y: 5.0, h: 0.4, w: 6.0 }, ...TITLE_STYLE });
	slide.addText([{ text: JSON.stringify(shadow3, "", 2) }], { ...BASE_CODE_OPTS, ...{ x: 0.5, y: 5.4, h: 1.7, w: 6 }, ...CODE_STYLE });
	slide.addImage({ ...{ path: IMAGE_PATHS.nycSubway.path, x: 7.0, y: 5.0, h: 2.1, w: 3.13 }, ...shadow3 });
}

/**
 * SLIDE 6: Image embedded in a shape (clip + source crop)
 * Demonstrates clipping a picture to a preset shape, a freeform `points` path, and the
 * "picture placeholder" composition: a freeform clip filled by a center-cropped photo
 * (`points` + `sizing:'cover'`). See docs/image-in-shape.md.
 * @param {PptxGenJS} pptx
 */
function genSlide06(pptx) {
	const slide = pptx.addSlide({ sectionTitle: "Images" });

	slide.addTable([[{ text: "Image Examples: Image Embedded In A Shape", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-images.html");
	slide.slideNumber = { x: "50%", y: "95%", w: 1, h: 1, color: COLOR_BLUE };

	const photo = IMAGE_PATHS.nycSubway.path;

	// 1) Preset clip: hexagon
	slide.addText("Preset clip: `shape:'hexagon'`", { x: 0.5, y: 0.6, w: 2.8, h: 0.3, fontSize: 11, color: COLOR_BLUE });
	slide.addImage({ path: photo, x: 0.5, y: 1.0, w: 2.5, h: 2.5, shape: "hexagon", sizing: { type: "cover", w: 2.5, h: 2.5 } });

	// 2) Preset clip: roundRect with corner radius
	slide.addText("Preset clip: `shape:'roundRect'`", { x: 3.4, y: 0.6, w: 2.8, h: 0.3, fontSize: 11, color: COLOR_BLUE });
	slide.addImage({ path: photo, x: 3.4, y: 1.0, w: 2.5, h: 2.5, shape: "roundRect", rectRadius: 0.35, sizing: { type: "cover", w: 2.5, h: 2.5 } });

	// 3) Freeform clip: triangle via `points`
	slide.addText("Freeform clip: `points` (triangle)", { x: 6.3, y: 0.6, w: 2.8, h: 0.3, fontSize: 11, color: COLOR_BLUE });
	slide.addImage({
		path: photo,
		x: 6.3,
		y: 1.0,
		w: 2.5,
		h: 2.5,
		points: [{ x: 1.25, y: 0 }, { x: 2.5, y: 2.5 }, { x: 0, y: 2.5 }, { close: true }],
		sizing: { type: "cover", w: 2.5, h: 2.5 },
	});

	// 4) The half-disc "D" cover: freeform `arcTo` clip + center-cropped photo
	slide.addText("Half-disc 'D' cover: `points`+`cover`", { x: 9.2, y: 0.6, w: 3.6, h: 0.3, fontSize: 11, color: COLOR_BLUE });
	const dW = 3.4,
		dH = 6.0,
		dx = 9.2,
		dy = 1.0;
	const fx = 0.3179 * dW; // x of the flat (right-flush) edge
	slide.addImage({
		path: photo,
		x: dx,
		y: dy,
		w: dW,
		h: dH,
		points: [
			{ x: fx, y: 0 },
			{ x: dW, y: 0 },
			{ x: dW, y: dH },
			{ x: fx, y: dH },
			{ x: 0, y: dH / 2, curve: { type: "arc", hR: dH / 2, wR: fx, stAng: 90, swAng: 180 } },
			{ close: true },
		],
		sizing: { type: "cover", w: dW, h: dH },
	});
}
