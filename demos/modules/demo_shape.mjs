/**
 * NAME: demo_shapes.mjs
 * AUTH: Brent Ely
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by maintained demos (./demos/node, ./demos/vite-demo).
 * VER.: 3.5.0
 * BLD.: 20210401
 */

/**
 * CUSTOM GEOMETRY:
 * Notes from the author [apresmoi](https://github.com/apresmoi):
 * I've implemented this by using a similar spec to the one used by `svg-points`.
 * The path or contour of the custom geometry is declared under the property points of the ShapeProps object.
 * With this implementation we are supporting all the custom geometry rules: moveTo, lnTo, arcTo, cubicBezTo, quadBezTo and close.
 *
 * A translation of an svg path to a custom geometry could be achieved by using the svg-points package and adding a custom translation between the arcs.
 * The svg arc is described by the variables x, y, rx, ry, xAxisRotation, largeArcFlag and sweepFlag.
 * On the other side the pptx freeform arc is described by hR, wR, stAng, swAng — and no end point, which
 * the renderer derives from the pen position, the radii and the sweep. So an svg->pptx arc translation
 * has to solve angles from the svg end point; there is no such solver here, and the DSL takes the OOXML
 * form directly.
 */

import { SchemeColor, ShapeType } from "@shbernal/ts-pptx";
import { BASE_TABLE_OPTS, BASE_TEXT_OPTS_L, BASE_TEXT_OPTS_R } from "./enums.mjs";

export function genSlides_Shape(pptx) {
	pptx.addSection({ title: "Shapes" });

	genSlide01(pptx);
	genSlide02(pptx);
	genSlide03(pptx);
}

/**
 * SLIDE 1: Misc Shape Types (no text)
 * @param {TsPptx} pptx
 */
function genSlide01(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Shapes" });

	slide.addTable([[{ text: "Shape Examples 1: Misc Shape Types (no text)", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	// TOP-ROW

	slide.addShape(ShapeType.rect, { x: 0.5, y: 0.8, w: 1.5, h: 3.0, fill: { color: SchemeColor.accent1 }, line: { type: "none" } });
	slide.addShape(ShapeType.ellipse, { x: 2.2, y: 0.8, w: 3.0, h: 1.5, fill: { type: "solid", color: SchemeColor.accent2 } });
	slide.addShape(ShapeType.custGeom, {
		x: 2.5,
		y: 2.6,
		w: 2.0,
		h: 1.0,
		fill: { color: SchemeColor.accent3 },
		line: { color: "151515", width: 1 },
		points: [
			{ x: 0.0, y: 0.0 },
			{ x: 0.5, y: 1.0 },
			{ x: 1.0, y: 0.8 },
			{ x: 1.5, y: 1.0 },
			{ x: 2.0, y: 0.0 },
			{ x: 0.0, y: 0.0, curve: { type: "quadratic", x1: 1.0, y1: 0.5 } },
			{ close: true },
		],
	});
	slide.addShape(ShapeType.rect, { x: 5.7, y: 0.8, w: 1.5, h: 3.0, fill: { color: SchemeColor.accent4 }, rotate: 45 });
	slide.addShape(ShapeType.ellipse, { x: 7.4, y: 1.5, w: 3.0, h: 1.5, fill: { color: SchemeColor.accent6 }, rotate: 90 }); // TEST: no type
	slide.addShape(ShapeType.roundRect, {
		x: 10,
		y: 0.8,
		w: 3.0,
		h: 1.5,
		rectRadius: 1,
		fill: { color: SchemeColor.accent5 },
		line: { color: "151515", width: 1 },
	});
	slide.addShape(ShapeType.arc, { x: 10.75, y: 2.45, w: 1.5, h: 1.45, fill: { color: SchemeColor.accent3 }, angleRange: [45, 315] });

	// BOTTOM ROW

	slide.addShape(ShapeType.line, { x: 4.2, y: 4.4, w: 5.0, h: 0.0, line: { color: SchemeColor.accent2, width: 1, dashType: "lgDash" } });
	slide.addShape(ShapeType.line, {
		x: 4.2,
		y: 4.8,
		w: 5.0,
		h: 0.0,
		line: { color: SchemeColor.accent2, width: 2, dashType: "dashDot", beginArrowType: "arrow" },
	});
	slide.addShape(ShapeType.line, { x: 4.2, y: 5.2, w: 5.0, h: 0.0, line: { color: SchemeColor.accent2, width: 3, endArrowType: "triangle" } });
	slide.addShape(ShapeType.line, {
		x: 4.2,
		y: 5.6,
		w: 5.0,
		h: 0.0,
		line: { color: SchemeColor.accent2, width: 4, beginArrowType: "diamond", endArrowType: "oval" },
	});

	slide.addShape(ShapeType.rtTriangle, {
		x: 0.4,
		y: 4.3,
		w: 6.0,
		h: 3.0,
		fill: { color: SchemeColor.accent5 },
		line: { color: SchemeColor.accent1, width: 3 },
		objectName: "First Right Triangle",
	});
	slide.addShape(ShapeType.rtTriangle, {
		x: 7.0,
		y: 4.3,
		w: 6.0,
		h: 3.0,
		fill: { color: SchemeColor.accent5 },
		line: { color: SchemeColor.accent1, width: 2 },
		flipH: true,
	});
}

/**
 * SLIDE 2: Misc Shape Types with Text
 * @param {TsPptx} pptx
 */
function genSlide02(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Shapes" });

	slide.addTable([[{ text: "Shape Examples 2: Misc Shape Types (with text)", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	slide.addText("RECTANGLE", {
		shape: ShapeType.rect,
		x: 0.5,
		y: 0.8,
		w: 1.5,
		h: 3.0,
		fill: { color: SchemeColor.accent1 },
		align: "center",
		fontSize: 14,
	});
	slide.addText("OVAL (transparency:50)", {
		shape: ShapeType.ellipse,
		x: 2.2,
		y: 0.8,
		w: 3.0,
		h: 1.5,
		fill: { type: "solid", color: SchemeColor.accent2, transparency: 50 },
		align: "center",
		fontSize: 14,
	});
	slide.addText("CUSTOM", {
		shape: ShapeType.custGeom,
		x: 2.5,
		y: 2.6,
		w: 2.0,
		h: 1.0,
		fill: { color: SchemeColor.accent3 },
		line: { color: "151515", width: 1 },
		points: [
			{ x: 0.0, y: 0.0 },
			{ x: 0.5, y: 1.0 },
			{ x: 1.0, y: 0.8 },
			{ x: 1.5, y: 1.0 },
			{ x: 2.0, y: 0.0 },
			{ x: 0.0, y: 0.0, curve: { type: "quadratic", x1: 1.0, y1: 0.5 } },
			{ close: true },
		],
		align: "center",
		fontSize: 14,
	});
	slide.addText("RECTANGLE (rotate:45)", {
		shape: ShapeType.rect,
		x: 5.7,
		y: 0.8,
		w: 1.5,
		h: 3.0,
		fill: { color: SchemeColor.accent4 },
		rotate: 45,
		align: "center",
		fontSize: 14,
	});
	// TEST: DEPRECATED: `alpha`
	slide.addText("OVAL (rotate:90, transparency:75)", {
		shape: ShapeType.ellipse,
		x: 7.4,
		y: 1.5,
		w: 3.0,
		h: 1.5,
		fill: { type: "solid", color: SchemeColor.accent6, transparency: 75 },
		rotate: 90,
		align: "center",
		fontSize: 14,
	});
	slide.addText("ROUNDED-RECTANGLE\ndashType:dash\nrectRadius:1", {
		shape: ShapeType.roundRect,
		x: 10,
		y: 0.8,
		w: 3.0,
		h: 1.5,
		fill: { color: SchemeColor.accent5 },
		align: "center",
		fontSize: 14,
		line: { color: "151515", size: 1, dashType: "dash" },
		rectRadius: 1,
	});
	slide.addText("ARC", {
		shape: ShapeType.arc,
		x: 10.75,
		y: 2.45,
		w: 1.5,
		h: 1.45,
		fill: { color: SchemeColor.accent3 },
		angleRange: [45, 315],
		line: { color: "151515", width: 1 },
		fontSize: 14,
	});
	//
	slide.addText("LINE size=1", {
		shape: ShapeType.line,
		align: "center",
		x: 4.15,
		y: 4.4,
		w: 5,
		h: 0,
		line: { color: SchemeColor.accent2, width: 1, dashType: "lgDash" },
	});
	slide.addText("LINE size=2", {
		shape: ShapeType.line,
		align: "left",
		x: 4.15,
		y: 4.8,
		w: 5,
		h: 0,
		line: { color: SchemeColor.accent2, width: 2, dashType: "dashDot", endArrowType: "arrow" },
	});
	slide.addText("LINE size=3", {
		shape: ShapeType.line,
		align: "right",
		x: 4.15,
		y: 5.2,
		w: 5,
		h: 0,
		line: { color: SchemeColor.accent2, width: 3, beginArrowType: "triangle" },
	});
	slide.addText("LINE size=4", {
		shape: ShapeType.line,
		x: 4.15,
		y: 5.6,
		w: 5,
		h: 0,
		line: { color: SchemeColor.accent2, width: 4, beginArrowType: "diamond", endArrowType: "oval", transparency: 50 },
	});
	//
	slide.addText("RIGHT-TRIANGLE", {
		shape: ShapeType.rtTriangle,
		align: "center",
		x: 0.4,
		y: 4.3,
		w: 6,
		h: 3,
		fill: { color: SchemeColor.accent5 },
		line: { color: "696969", width: 3 },
	});
	slide.addText("HYPERLINK-SHAPE", {
		shape: ShapeType.rtTriangle,
		align: "center",
		x: 7.0,
		y: 4.3,
		w: 6,
		h: 3,
		fill: { color: SchemeColor.accent5 },
		line: { color: "696969", width: 2 },
		flipH: true,
		hyperlink: { url: "https://example.com", tooltip: "Visit Homepage" },
	});
}

/**
 * SLIDE 3: Interactive action buttons (slide-show navigation via ppaction://hlinkshowjump)
 * @param {TsPptx} pptx
 */
function genSlide03(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Shapes" });

	slide.addTable([[{ text: "Shape Examples 3: Action buttons (navigation)", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	// Each action button wires an <a:hlinkClick action="ppaction://hlinkshowjump?jump=…"/> so it
	// actually navigates when the deck is run as a slide show (not just static geometry).
	const buttons = [
		{ shape: ShapeType.actionButtonBeginning, action: "firstslide", tooltip: "First slide" },
		{ shape: ShapeType.actionButtonBackPrevious, action: "previousslide", tooltip: "Previous slide" },
		{ shape: ShapeType.actionButtonForwardNext, action: "nextslide", tooltip: "Next slide" },
		{ shape: ShapeType.actionButtonEnd, action: "lastslide", tooltip: "Last slide" },
		{ shape: ShapeType.actionButtonReturn, action: "lastslideviewed", tooltip: "Return" },
	];
	buttons.forEach((btn, idx) => {
		slide.addShape(btn.shape, {
			x: 0.5 + idx * 1.2,
			y: 3.0,
			w: 1.0,
			h: 1.0,
			fill: { color: SchemeColor.accent1 },
			line: { color: "696969", width: 1 },
			hyperlink: { action: btn.action, tooltip: btn.tooltip },
		});
	});
}
