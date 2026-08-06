/**
 * Showcase deck: "Field Notes — Four Cities After Dark".
 *
 * The visual counterpart to the quarterly review. Where that deck is charts and tables on a
 * light grid, this one is full-bleed photography, gradient scrims over images, duotone and
 * rounded picture treatments, grouped annotation callouts, an embedded video, an embedded 3D
 * model, and hyperlinks.
 *
 * Build it with `pnpm demos:build field-notes`.
 *
 * Unlike the quarterly review, this deck is Node-only by nature: it loads photographs, a video
 * and a `.glb` from `demos/common` by path, so it cannot run in a browser without those assets
 * being served. That is the honest split between the two showcases, not an oversight.
 */
import TsPptx, { ShapeType } from "@shbernal/ts-pptx";
import { image, imageDataUri, media } from "../lib/assets.mjs";
import { WIDE, centeredRow, columns } from "../lib/layout.mjs";
import { BRAND, FONT, MASTER, TYPE, applyDesign, scrim, wordmark } from "./design.mjs";

const CITIES = [
	{
		name: "Chicago",
		country: "United States",
		photo: "chicago_bean_bohne.jpg",
		note: "A polished bean that eats the skyline and hands it back curved.",
	},
	{
		name: "Tokyo",
		country: "Japan",
		photo: "tokyo-subway-route-map.jpg",
		note: "Thirteen operators, one map, and a colour system that never repeats.",
	},
	{
		name: "New York",
		country: "United States",
		photo: "nyc-subway.png",
		note: "The only network here that never closes, and looks it.",
	},
	{
		name: "Sydney",
		country: "Australia",
		photo: "sydney_harbour_bridge_night.jpg",
		note: "A bridge lit like an argument for staying out past midnight.",
	},
];

function addCover(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.cover, sectionTitle: "Front matter" });

	slide.addImage({
		path: image("sydney_harbour_bridge_night.jpg"),
		x: 0,
		y: 0,
		w: WIDE.w,
		h: WIDE.h,
		sizing: { type: "cover", w: WIDE.w, h: WIDE.h },
		altText: "Sydney Harbour Bridge lit at night",
	});
	// Scrim from the left, so type sits on ink and the photograph stays visible on the right.
	slide.addShape(ShapeType.rect, scrim({ x: 0, y: 0, w: 8.6, h: WIDE.h, angle: 0, from: 6, to: 100 }));

	wordmark(slide, { x: 0.85, y: 0.85 });
	slide.addText("Four Cities\nAfter Dark", {
		x: 0.85,
		y: 2.3,
		w: 7.5,
		h: 2.4,
		margin: 0,
		fontFace: FONT.head,
		fontSize: TYPE.display,
		color: BRAND.bone,
		lineSpacingMultiple: 1.02,
	});
	slide.addShape(ShapeType.rect, { x: 0.85, y: 4.85, w: 1.1, h: 0.04, fill: { color: BRAND.amber } });
	slide.addText("Four transit systems, photographed between last train and first light — and what each one\nsays about the city that built it.", {
		x: 0.85,
		y: 5.15,
		w: 7,
		h: 0.9,
		margin: 0,
		fontFace: FONT.body,
		fontSize: TYPE.standfirst,
		color: BRAND.sand,
		lineSpacingMultiple: 1.35,
	});
	slide.addText("Volume 04  ·  Autumn 2026", {
		x: 0.85,
		y: 6.5,
		w: 6,
		h: 0.3,
		margin: 0,
		fontFace: FONT.body,
		fontSize: TYPE.caption,
		color: BRAND.ash,
		charSpacing: 2,
	});

	slide.addNotes("Cover photograph: Sydney Harbour Bridge. Shot handheld at 1/15s, which is why the rigging is soft.");
}

function addContents(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.editorial, sectionTitle: "Front matter" });

	slide.addText("In this volume", {
		x: 0.85,
		y: 0.75,
		w: 8,
		h: 0.7,
		margin: 0,
		valign: "middle",
		fontFace: FONT.head,
		fontSize: TYPE.title,
		color: BRAND.bone,
	});

	const cardW = 2.72;
	const xs = centeredRow(CITIES.length, cardW, 0.38);
	CITIES.forEach((city, i) => {
		slide.addImage({
			path: image(city.photo),
			x: xs[i],
			y: 2.0,
			w: cardW,
			h: 2.0,
			sizing: { type: "cover", w: cardW, h: 2.0 },
			rounding: false,
			altText: `${city.name} — ${city.note}`,
		});
		slide.addShape(ShapeType.rect, { x: xs[i], y: 4.06, w: 0.55, h: 0.03, fill: { color: BRAND.amber } });
		slide.addText(String(i + 1).padStart(2, "0"), {
			x: xs[i],
			y: 1.6,
			w: 1,
			h: 0.32,
			margin: 0,
			fontFace: FONT.body,
			fontSize: TYPE.caption,
			color: BRAND.amber,
			bold: true,
			charSpacing: 2,
		});
		slide.addText(city.name, {
			x: xs[i],
			y: 4.22,
			w: cardW,
			h: 0.45,
			margin: 0,
			fontFace: FONT.head,
			fontSize: TYPE.heading,
			color: BRAND.bone,
		});
		slide.addText(city.note, {
			x: xs[i],
			y: 4.72,
			w: cardW,
			h: 1.1,
			margin: 0,
			fontFace: FONT.body,
			fontSize: TYPE.body,
			color: BRAND.ash,
			lineSpacingMultiple: 1.4,
		});
	});

	slide.addNotes("Order is west to east by first visit, not by page count.");
}

function addPlate(pptx, { city, kicker, headline, body, quote, attribution }) {
	const slide = pptx.addSlide({ masterTitle: MASTER.plate, sectionTitle: city.name });
	const photoW = 8.2;

	slide.addImage({
		path: image(city.photo),
		x: 0,
		y: 0,
		w: photoW,
		h: WIDE.h,
		sizing: { type: "cover", w: photoW, h: WIDE.h },
		altText: `${city.name}, ${city.country}`,
	});
	// A short scrim on the photo's right edge blends it into the text panel instead of
	// leaving a hard seam where the image stops.
	slide.addShape(ShapeType.rect, scrim({ x: photoW - 1.6, y: 0, w: 1.6, h: WIDE.h, angle: 0, from: 100, to: 0 }));

	const panelX = photoW;
	const panelW = WIDE.w - photoW;
	slide.addShape(ShapeType.rect, { x: panelX, y: 0, w: panelW, h: WIDE.h, fill: { color: BRAND.ink } });

	const textX = panelX + 0.6;
	const textW = panelW - 1.2;
	slide.addText(kicker.toUpperCase(), {
		x: textX,
		y: 0.95,
		w: textW,
		h: 0.3,
		margin: 0,
		fontFace: FONT.body,
		fontSize: TYPE.caption,
		color: BRAND.amber,
		bold: true,
		charSpacing: 2.4,
	});
	slide.addText(headline, {
		x: textX,
		y: 1.35,
		w: textW,
		h: 1.5,
		margin: 0,
		fontFace: FONT.head,
		fontSize: 28,
		color: BRAND.bone,
		lineSpacingMultiple: 1.12,
	});
	slide.addText(body, {
		x: textX,
		y: 3.0,
		w: textW,
		h: 2.1,
		margin: 0,
		fontFace: FONT.body,
		fontSize: TYPE.body,
		color: BRAND.sand,
		lineSpacingMultiple: 1.5,
	});
	slide.addText(`“${quote}”`, {
		x: textX,
		y: 5.25,
		w: textW,
		h: 1.0,
		margin: 0,
		fontFace: FONT.head,
		fontSize: 15,
		color: BRAND.amber,
		italic: true,
		lineSpacingMultiple: 1.3,
	});
	slide.addText(attribution, {
		x: textX,
		y: 6.3,
		w: textW,
		h: 0.3,
		margin: 0,
		fontFace: FONT.body,
		fontSize: 9,
		color: BRAND.stone,
		charSpacing: 1.4,
	});

	// The caption hangs from the amber rule the plate master places.
	slide.addText(`${city.name}, ${city.country}`, {
		x: 0.85,
		y: 6.72,
		w: 6,
		h: 0.3,
		margin: 0,
		fontFace: FONT.body,
		fontSize: 9,
		color: BRAND.sand,
		charSpacing: 1.4,
	});

	return slide;
}

function addAnnotatedMap(pptx) {
	const city = CITIES[1];
	const slide = pptx.addSlide({ masterTitle: MASTER.editorial, sectionTitle: city.name });

	slide.addText("Reading the Tokyo map", {
		x: 0.85,
		y: 0.7,
		w: 8,
		h: 0.6,
		margin: 0,
		valign: "middle",
		fontFace: FONT.head,
		fontSize: TYPE.title,
		color: BRAND.bone,
	});
	slide.addText("Thirteen operators share one diagram. The colour system is the only thing holding it together.", {
		x: 0.85,
		y: 1.35,
		w: 9,
		h: 0.4,
		margin: 0,
		fontFace: FONT.body,
		fontSize: TYPE.body,
		color: BRAND.ash,
	});

	// Duotone pushes the map into the deck's palette so it reads as a plate rather than a
	// screenshot — a picture effect PowerPoint applies natively, no pre-processing.
	slide.addImage({
		path: image(city.photo),
		x: 0.85,
		y: 1.95,
		w: 7.4,
		h: 4.5,
		sizing: { type: "contain", w: 7.4, h: 4.5 },
		duotone: { shadow: BRAND.ink, highlight: BRAND.sand },
		altText: "Tokyo subway route map",
	});

	const notes = [
		{
			n: "01",
			title: "Two networks, one fare gate",
			body: "Tokyo Metro and Toei run separately. The transfer surcharge is the map’s one invisible feature.",
		},
		{ n: "02", title: "The loop as an anchor", body: "Every other line is described by where it meets the Yamanote. The map is drawn outward from it." },
		{ n: "03", title: "Letters before colours", body: "Each line carries a letter and a number so the diagram survives being photocopied in grey." },
	];
	const noteX = 8.6;
	const noteW = WIDE.w - noteX - 0.85;
	notes.forEach((note, i) => {
		const y = 2.05 + i * 1.5;
		slide.addGroup([
			{
				text: {
					text: note.n,
					options: {
						x: noteX,
						y,
						w: 0.6,
						h: 0.3,
						margin: 0,
						fontFace: FONT.body,
						fontSize: TYPE.caption,
						color: BRAND.amber,
						bold: true,
						charSpacing: 1.6,
					},
				},
			},
			{
				text: {
					text: note.title,
					options: {
						x: noteX,
						y: y + 0.3,
						w: noteW,
						h: 0.34,
						margin: 0,
						valign: "middle",
						fontFace: FONT.head,
						fontSize: 15,
						color: BRAND.bone,
					},
				},
			},
			{
				text: {
					text: note.body,
					options: {
						x: noteX,
						y: y + 0.66,
						w: noteW,
						h: 0.8,
						margin: 0,
						fontFace: FONT.body,
						fontSize: 11,
						color: BRAND.ash,
						lineSpacingMultiple: 1.4,
					},
				},
			},
		]);
	});

	slide.addNotes("The duotone treatment here is a picture effect, not a pre-processed image — the source file is the plain map.");
}

function addGrid(pptx) {
	const city = CITIES[2];
	const slide = pptx.addSlide({ masterTitle: MASTER.editorial, sectionTitle: city.name });

	slide.addText("New York, in three frames", {
		x: 0.85,
		y: 0.7,
		w: 9,
		h: 0.6,
		margin: 0,
		valign: "middle",
		fontFace: FONT.head,
		fontSize: TYPE.title,
		color: BRAND.bone,
	});

	const { x: xs, w } = columns(3, { gap: 0.4, left: 0.85, width: WIDE.w - 1.7 });
	const frames = [
		{ photo: city.photo, caption: "The map that never sleeps" },
		{ photo: "nyc-subway.png", caption: "Express and local, same platform" },
		{ photo: "wiki-example.jpg", caption: "Above ground, briefly" },
	];
	frames.forEach((frame, i) => {
		slide.addImage({
			path: image(frame.photo),
			x: xs[i],
			y: 1.7,
			w,
			h: 3.4,
			sizing: { type: "cover", w, h: 3.4 },
			altText: frame.caption,
		});
		slide.addShape(ShapeType.rect, { x: xs[i], y: 5.22, w: 0.45, h: 0.03, fill: { color: BRAND.amber } });
		slide.addText(frame.caption, {
			x: xs[i],
			y: 5.35,
			w,
			h: 0.6,
			margin: 0,
			fontFace: FONT.body,
			fontSize: 11,
			color: BRAND.sand,
			lineSpacingMultiple: 1.35,
		});
	});

	slide.addNotes("Three frames rather than one: the point is the system, and no single photograph holds a system.");
}

async function addMotion(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.editorial, sectionTitle: "Motion" });

	slide.addText("Motion", {
		x: 0.85,
		y: 0.7,
		w: 8,
		h: 0.6,
		margin: 0,
		valign: "middle",
		fontFace: FONT.head,
		fontSize: TYPE.title,
		color: BRAND.bone,
	});
	slide.addText("Video embeds travel inside the .pptx, so the deck plays with no network and no external file.", {
		x: 0.85,
		y: 1.35,
		w: 9,
		h: 0.4,
		margin: 0,
		fontFace: FONT.body,
		fontSize: TYPE.body,
		color: BRAND.ash,
	});

	slide.addMedia({
		type: "video",
		path: media("sample.mp4"),
		// `cover` takes a base64 data URI, not a path — the one place in these decks where
		// an asset has to be inlined by hand rather than passed as a `path:`.
		cover: await imageDataUri("cover_video_16x9.png"),
		x: 0.85,
		y: 2.0,
		w: 7.6,
		h: 4.28,
	});

	const noteX = 8.85;
	const noteW = WIDE.w - noteX - 0.85;
	slide.addText(
		[
			{ text: "Embedded, not linked\n", options: { fontFace: FONT.head, fontSize: 16, color: BRAND.bone, breakLine: true } },
			{
				text: "The media bytes live in the package alongside the slides. A linked video would break the moment this file is emailed to anyone.",
				options: { fontSize: 11, color: BRAND.sand, breakLine: true },
			},
			{ text: "\n", options: { fontSize: 8, breakLine: true } },
			{
				text: "The poster frame is a separate image, so the slide reads correctly in thumbnails and print, where nothing can play.",
				options: { fontSize: 11, color: BRAND.ash },
			},
		],
		{
			x: noteX,
			y: 2.0,
			w: noteW,
			h: 3.0,
			margin: 0.18,
			shape: ShapeType.roundRect,
			rectRadius: 0.08,
			fill: { color: BRAND.stone, transparency: 60 },
			fontFace: FONT.body,
			lineSpacingMultiple: 1.45,
		},
	);

	slide.addNotes("If the video will not play in the room, the poster frame is doing its job — carry on and mention it.");
}

function addDimension(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.editorial, sectionTitle: "Motion" });

	slide.addText("Dimension", {
		x: 0.85,
		y: 0.7,
		w: 8,
		h: 0.6,
		margin: 0,
		valign: "middle",
		fontFace: FONT.head,
		fontSize: TYPE.title,
		color: BRAND.bone,
	});
	slide.addText("A glTF binary embedded in the package: PowerPoint 2019 and later renders it live and lets you orbit it.", {
		x: 0.85,
		y: 1.35,
		w: 9,
		h: 0.4,
		margin: 0,
		fontFace: FONT.body,
		fontSize: TYPE.body,
		color: BRAND.ash,
	});

	slide.addModel3d({
		path: media("cube.glb"),
		// A real render, not the gray placeholder — everything that is not PowerPoint 2019+ draws
		// this picture, including the thumbnail this deck is most often seen as.
		preview: { path: image("cube_3d_preview.png") },
		// The cube spans 2 model units, which is exactly what the default scale assumes; a model of
		// any other size needs `meterPerModelUnit: 1 / <largest bounding-box dimension>`.
		camera: { pos: { x: 1.3516, y: 1.0988, z: 1.9305 }, lookAt: { x: 0, y: 0, z: 0 }, fov: 45 },
		objectName: "Cube3D",
		altText: "A shaded cube, viewed from above one corner.",
		x: 1.6,
		y: 2.0,
		w: 6.1,
		h: 4.28,
	});

	const noteX = 8.85;
	const noteW = WIDE.w - noteX - 0.85;
	slide.addText(
		[
			{ text: "The camera is yours\n", options: { fontFace: FONT.head, fontSize: 16, color: BRAND.bone, breakLine: true } },
			{
				text: "PowerPoint frames a model from its bounding box. ts-pptx never opens the .glb, so it ships a fixed default and lets you place the camera — this one orbits to a corner.",
				options: { fontSize: 11, color: BRAND.sand, breakLine: true },
			},
			{ text: "\n", options: { fontSize: 8, breakLine: true } },
			{
				text: "The preview is a real PowerPoint render of the same model, so the slide reads correctly in print and in PDF.",
				options: { fontSize: 11, color: BRAND.ash },
			},
		],
		{
			x: noteX,
			y: 2.0,
			w: noteW,
			h: 3.0,
			margin: 0.18,
			shape: ShapeType.roundRect,
			rectRadius: 0.08,
			fill: { color: BRAND.stone, transparency: 60 },
			fontFace: FONT.body,
			lineSpacingMultiple: 1.45,
		},
	);

	slide.addNotes("Click the model in slide show and drag — it orbits. In anything older than PowerPoint 2019 it is a still picture.");
}

function addColophon(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.colophon, sectionTitle: "Back matter" });

	wordmark(slide, { x: 0.95, y: 1.4 });
	slide.addText("Colophon", {
		x: 0.95,
		y: 1.95,
		w: 8,
		h: 0.9,
		margin: 0,
		valign: "middle",
		fontFace: FONT.head,
		fontSize: 40,
		color: BRAND.bone,
	});
	slide.addShape(ShapeType.rect, { x: 0.95, y: 3.0, w: 1.1, h: 0.04, fill: { color: BRAND.amber } });

	const { x: xs, w } = columns(2, { gap: 0.7, left: 0.95, width: 8.4 });
	slide.addText(
		[
			{ text: "Photography\n", options: { color: BRAND.amber, fontSize: 10, bold: true, charSpacing: 1.6, breakLine: true } },
			{
				text: "Shared demo assets from the ts-pptx repository, reproduced here under their original licences.",
				options: { color: BRAND.sand, fontSize: 11 },
			},
		],
		{ x: xs[0], y: 3.4, w, h: 1.4, margin: 0, fontFace: FONT.body, lineSpacingMultiple: 1.45 },
	);
	slide.addText(
		[
			{ text: "Typesetting\n", options: { color: BRAND.amber, fontSize: 10, bold: true, charSpacing: 1.6, breakLine: true } },
			{
				text: "Set in Georgia and Segoe UI. Every rule, scrim, and picture effect on these pages is generated — no slide was touched in PowerPoint.",
				options: { color: BRAND.sand, fontSize: 11 },
			},
		],
		{ x: xs[1], y: 3.4, w, h: 1.4, margin: 0, fontFace: FONT.body, lineSpacingMultiple: 1.45 },
	);

	slide.addText(
		[
			{ text: "Built with ", options: { color: BRAND.ash, fontSize: 12 } },
			{
				text: "ts-pptx",
				options: {
					color: BRAND.amber,
					fontSize: 12,
					bold: true,
					hyperlink: { url: "https://github.com/shbernal/ts-pptx", tooltip: "ts-pptx on GitHub" },
				},
			},
			{ text: "  ·  regenerate this deck with  ", options: { color: BRAND.ash, fontSize: 12 } },
			{ text: "pnpm demos:build", options: { color: BRAND.sand, fontSize: 12, fontFace: FONT.mono } },
		],
		{ x: 0.95, y: 5.6, w: 10, h: 0.4, margin: 0, fontFace: FONT.body },
	);

	slide.addNotes("The hyperlink on this slide is a real relationship in the package, not styled text.");
}

/** Build the deck and write it to `outFile`. */
export async function build(outFile) {
	const pptx = new TsPptx();

	pptx.title = "Field Notes — Four Cities After Dark";
	pptx.subject = "Photo essay";
	pptx.author = "Field Notes";
	pptx.company = "Field Notes";

	applyDesign(pptx);
	// Sections must exist before a slide can be filed into one; `addSlide({ sectionTitle })`
	// warns rather than creating a missing section.
	for (const title of ["Front matter", ...CITIES.map((c) => c.name), "Motion", "Back matter"]) {
		pptx.addSection({ title });
	}

	addCover(pptx);
	addContents(pptx);
	addPlate(pptx, {
		city: CITIES[0],
		kicker: "Chicago",
		headline: "A city that decided to look at itself",
		body: "Millennium Park was supposed to be a lid over a rail yard. What it became is the one place in Chicago where the skyline is not above you but wrapped around you, upside down, at eye level — and where the crowd photographing it becomes part of the photograph.",
		quote: "You cannot take a picture of it without taking a picture of yourself taking it.",
		attribution: "FIELD NOTE  ·  22:40, NORTH LOOP",
	});
	addAnnotatedMap(pptx);
	addGrid(pptx);
	addPlate(pptx, {
		city: CITIES[3],
		kicker: "Sydney",
		headline: "Lit like an argument for staying out",
		body: "The bridge carries eight lanes, two rail tracks, a cycleway and a footpath, and at night it advertises none of them. What it advertises is the harbour underneath — which is the correct order of priorities for a piece of infrastructure that has been photographed more than it has been crossed.",
		quote: "Infrastructure that knows it is being looked at behaves differently after dark.",
		attribution: "FIELD NOTE  ·  23:15, MILSONS POINT",
	});
	await addMotion(pptx);
	addDimension(pptx);
	addColophon(pptx);

	return await pptx.writeFile({ fileName: outFile });
}

export const showcase = {
	slug: "field-notes",
	title: "Field Notes — Four Cities After Dark",
	description: "Visual flagship: full-bleed photography, gradient scrims, duotone, an embedded video, an embedded 3D model, hyperlinks.",
	fileName: "Field_Notes_Four_Cities.pptx",
	build,
};
