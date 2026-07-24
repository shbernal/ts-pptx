/**
 * NAME: demo_group.mjs
 * AUTH: ts-pptx
 * DESC: Grouping demos — addGroup() (build from descriptors) and groupObjects() (wrap existing objects)
 * DEPS: Used by maintained demos (./demos/node, ./demos/vite-demo).
 */

import { BASE_TABLE_OPTS, BASE_TEXT_OPTS_L, BASE_TEXT_OPTS_R } from "./enums.mjs";

export function genSlides_Group(pptx) {
	pptx.addSection({ title: "Groups" });

	genSlide01(pptx);
	genSlide02(pptx);
}

/**
 * SLIDE 1: addGroup() — flat group, nested group, and a rotated/flipped group
 * @param {TsPptx} pptx
 */
function genSlide01(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Groups" });

	slide.addTable(
		[[{ text: "Group Examples 1: addGroup() (build a group from child descriptors)", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]],
		BASE_TABLE_OPTS,
	);

	// A flat group: a rect + label + a second rect, wrapped as one selectable object.
	// Frame omitted -> auto-bounds (the bounding box of the children). Children keep
	// their slide-absolute coordinates; the group is visually a no-op.
	slide.addGroup(
		[
			{ rect: { x: 0.5, y: 1.0, w: 2.5, h: 1.2, fill: { color: "4472C4" }, line: { color: "2F528F", width: 1 } } },
			{
				text: {
					text: "Branding",
					options: { x: 0.5, y: 1.0, w: 2.5, h: 1.2, color: "FFFFFF", align: "center", valign: "middle", bold: true, fontSize: 18 },
				},
			},
			{ shape: { type: pptx.ShapeType.ellipse, options: { x: 3.1, y: 1.0, w: 1.2, h: 1.2, fill: { color: "ED7D31" } } } },
		],
		{ objectName: "FlatGroup" },
	);
	slide.addText("^ addGroup([rect, text, ellipse]) — one Selection Pane entry", { x: 0.5, y: 2.35, w: 5, h: 0.4, fontSize: 11, color: "606060" });

	// A nested group: an outer group whose members include another group. Names number
	// per slide and inside-out (the inner group takes the lower `Group N` index).
	slide.addGroup(
		[
			{ rect: { x: 5.5, y: 1.0, w: 1.5, h: 1.2, fill: { color: "70AD47" } } },
			{
				group: {
					children: [
						{ rect: { x: 7.2, y: 1.0, w: 1.5, h: 1.2, fill: { color: "FFC000" } } },
						{ text: { text: "Nested", options: { x: 7.2, y: 1.0, w: 1.5, h: 1.2, align: "center", valign: "middle", bold: true } } },
					],
					options: { objectName: "InnerGroup" },
				},
			},
		],
		{ objectName: "OuterGroup" },
	);
	slide.addText("^ a group nested inside a group", { x: 5.5, y: 2.35, w: 5, h: 0.4, fontSize: 11, color: "606060" });

	// A rotated + flipped group. rotate/flipH apply to the whole group (about its pivot);
	// the child space stays identity, so the children are not independently repositioned.
	slide.addGroup(
		[
			{ rect: { x: 2.5, y: 4.0, w: 3.0, h: 1.4, fill: { color: "5B9BD5" }, line: { color: "41719C", width: 1 } } },
			{
				text: {
					text: "rotate: 20, flipH",
					options: { x: 2.5, y: 4.0, w: 3.0, h: 1.4, color: "FFFFFF", align: "center", valign: "middle", bold: true },
				},
			},
		],
		{ rotate: 20, flipH: true, objectName: "RotatedGroup", altText: "A rotated, flipped group", objectLock: { noMove: true } },
	);
	slide.addText("^ rotate/flipH/objectLock/altText on the group as a whole", { x: 2.5, y: 5.5, w: 6, h: 0.4, fontSize: 11, color: "606060" });
}

/**
 * SLIDE 2: groupObjects() — group objects already added to the slide, by objectName
 * @param {TsPptx} pptx
 */
function genSlide02(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Groups" });

	slide.addTable(
		[[{ text: "Group Examples 2: groupObjects() (wrap objects already on the slide)", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]],
		BASE_TABLE_OPTS,
	);

	// Independent add* calls, each with an objectName. groupObjects() then folds two of
	// them into one group after the fact, leaving the third loose. Children keep their
	// existing z-order and geometry; the wrapper takes the topmost member's former slot.
	slide.addShape(pptx.ShapeType.rect, { x: 1.0, y: 1.2, w: 3.0, h: 1.0, fill: { color: "C00000" }, objectName: "Header" });
	slide.addText("Caption", {
		x: 1.2,
		y: 2.4,
		w: 2.6,
		h: 0.6,
		color: "FFFFFF",
		fill: { color: "7030A0" },
		align: "center",
		valign: "middle",
		objectName: "Caption",
	});
	slide.addShape(pptx.ShapeType.rect, { x: 6.0, y: 1.2, w: 2.0, h: 1.0, fill: { color: "00B050" }, objectName: "Loose" });

	slide.groupObjects(["Header", "Caption"], { objectName: "Banner" });

	slide.addText('groupObjects(["Header", "Caption"], { objectName: "Banner" }) — "Loose" stays top-level', {
		x: 1.0,
		y: 3.3,
		w: 10,
		h: 0.4,
		fontSize: 12,
		color: "606060",
	});

	// A connector can bind to a shape inside the group by its objectName — grouped
	// children are still addressable slide-wide.
	slide.addConnector({ type: "elbow", x1: 8.0, y1: 1.7, x2: 4.0, y2: 1.7, endShape: "Header", color: "404040", width: 2 });
	slide.addText("^ a connector bound to 'Header' (a shape now inside 'Banner')", { x: 1.0, y: 3.9, w: 10, h: 0.4, fontSize: 11, color: "606060" });
}
