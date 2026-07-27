/**
 * Showcase deck: "Meridian Group — Q3 FY26 Business Review".
 *
 * The corporate flagship. It leans on the parts of ts-pptx a report deck actually needs:
 * a themed colour scheme, five slide masters, native gradients, grouped composite shapes,
 * charts (stacked column, doughnut, line), a hand-styled table, and speaker notes.
 *
 * Build it with `pnpm demos:build quarterly-review`.
 *
 * This module imports nothing from `node:` — no filesystem, no paths, and every asset it
 * needs is drawn rather than loaded. That is what lets `demos/vite-demo` import it and build
 * the same deck in a browser: the deck is the deck, and only the runner differs.
 */
import TsPptx, { ChartType, ShapeType } from "@shbernal/ts-pptx";
import { CONTENT_W, MARGIN, WIDE, centeredRow, columns, signedPct } from "../lib/layout.mjs";
import { AGENDA, KPIS, QUARTERS, REGIONS, RETENTION, REVENUE_BY_SEGMENT, REVENUE_MIX, ROADMAP, SEGMENTS } from "./data.mjs";
import { BRAND, CHART_COLORS, FONT, MASTER, TYPE, applyDesign, kpiCard, slideTitle } from "./design.mjs";

const TOTAL_Q3 = SEGMENTS.reduce((sum, s) => sum + s.value, 0);

/** Deck sections, declared up front — `addSlide({ sectionTitle })` only files a slide into
 *  a section that already exists, and warns (rather than creating one) if it does not. */
const SECTIONS = ["Opening", "Performance", "Operations", "Closing"];

function addCover(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.title, sectionTitle: "Opening" });

	slide.addText("MERIDIAN GROUP", {
		x: 0.95,
		y: 2.55,
		w: 8,
		h: 0.35,
		margin: 0,
		fontFace: FONT.body,
		fontSize: TYPE.caption,
		color: BRAND.teal,
		bold: true,
		charSpacing: 3,
	});
	slide.addText("Q3 FY26\nBusiness Review", {
		x: 0.95,
		y: 2.95,
		w: 9,
		h: 1.6,
		margin: 0,
		fontFace: FONT.head,
		fontSize: TYPE.display,
		color: BRAND.white,
		lineSpacingMultiple: 1.05,
	});
	slide.addText("Revenue up 13.7% on Platform strength — and the payback question\nwe need to settle before we sign off on Q4 hiring.", {
		x: 0.95,
		y: 4.85,
		w: 8.2,
		h: 0.9,
		margin: 0,
		fontFace: FONT.body,
		fontSize: TYPE.subhead,
		color: BRAND.mist,
		lineSpacingMultiple: 1.3,
	});
	slide.addText("Prepared for the board  ·  14 October 2026", {
		x: 0.95,
		y: 6.4,
		w: 8,
		h: 0.3,
		margin: 0,
		fontFace: FONT.body,
		fontSize: TYPE.caption,
		color: BRAND.slate,
		charSpacing: 1.2,
	});

	slide.addNotes(
		"Open on the payback number, not the revenue number. Revenue is the easy half and they will have read it; " +
			"CAC payback moving from 15 to 17 months is the decision on the table today.",
	);
}

function addAgenda(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.content, sectionTitle: "Opening" });
	slideTitle(slide, "What we will cover", "Agenda");

	const rowH = 1.18;
	const top = 1.75;
	AGENDA.forEach((item, i) => {
		const y = top + i * rowH;
		slide.addGroup([
			{ rect: { x: MARGIN, y, w: CONTENT_W, h: 0.012, fill: { color: BRAND.mist } } },
			{
				text: {
					text: String(i + 1).padStart(2, "0"),
					options: {
						x: MARGIN,
						y: y + 0.16,
						w: 0.95,
						h: 0.7,
						margin: 0,
						valign: "middle",
						fontFace: FONT.head,
						fontSize: 30,
						color: BRAND.mist,
					},
				},
			},
			{
				text: {
					text: item.title,
					options: {
						x: MARGIN + 1.05,
						y: y + 0.14,
						w: 5.4,
						h: 0.42,
						margin: 0,
						valign: "middle",
						fontFace: FONT.head,
						fontSize: TYPE.heading,
						color: BRAND.ink,
					},
				},
			},
			{
				text: {
					text: item.detail,
					options: {
						x: MARGIN + 1.05,
						y: y + 0.58,
						w: 9.5,
						h: 0.36,
						margin: 0,
						valign: "middle",
						fontFace: FONT.body,
						fontSize: TYPE.body,
						color: BRAND.slate,
					},
				},
			},
		]);
	});

	slide.addNotes("Four items, twenty minutes. Hold questions on the roadmap until the last section.");
}

function addSectionDivider(pptx, number, title, standfirst) {
	const slide = pptx.addSlide({ masterTitle: MASTER.section, sectionTitle: title });

	slide.addText(number, {
		x: 8.6,
		y: 1.1,
		w: 4,
		h: 4.2,
		margin: 0,
		align: "right",
		valign: "middle",
		fontFace: FONT.head,
		fontSize: 200,
		color: BRAND.navy,
	});
	slide.addText(title, {
		x: 0.95,
		y: 3.45,
		w: 7.5,
		h: 0.85,
		margin: 0,
		valign: "middle",
		fontFace: FONT.head,
		fontSize: TYPE.section,
		color: BRAND.white,
	});
	slide.addText(standfirst, {
		x: 0.95,
		y: 4.6,
		w: 7,
		h: 0.8,
		margin: 0,
		fontFace: FONT.body,
		fontSize: TYPE.subhead,
		color: BRAND.mist,
		lineSpacingMultiple: 1.3,
	});
	return slide;
}

function addKpiBand(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.content, sectionTitle: "Performance" });
	slideTitle(slide, "The quarter in four numbers", "Performance");

	const cardW = 2.75;
	const cardH = 2.05;
	const xs = centeredRow(KPIS.length, cardW, 0.4);
	KPIS.forEach((kpi, i) => {
		slide.addGroup(kpiCard({ x: xs[i], y: 1.95, w: cardW, h: cardH, ...kpi }));
	});

	slide.addText(
		[
			{ text: "Read together: ", options: { bold: true, color: BRAND.ink } },
			{
				text: "we bought this quarter’s growth with a longer payback. Three of four move the right way, and the fourth is the price of the other three.",
				options: { color: BRAND.slate },
			},
		],
		{
			x: MARGIN,
			y: 4.5,
			w: CONTENT_W,
			h: 0.6,
			margin: 0,
			fontFace: FONT.body,
			fontSize: TYPE.subhead,
			lineSpacingMultiple: 1.3,
		},
	);

	slide.addShape(ShapeType.rect, { x: MARGIN, y: 5.35, w: 0.08, h: 1.2, fill: { color: BRAND.amber } });
	slide.addText(
		"The EMEA expansion carried 16.8% quarter-on-quarter growth on 208 people. The same expansion is\nwhat pushed blended CAC payback out by two months. Both facts are the same decision.",
		{
			x: MARGIN + 0.3,
			y: 5.35,
			w: CONTENT_W - 0.3,
			h: 1.2,
			margin: 0,
			valign: "middle",
			fontFace: FONT.body,
			fontSize: TYPE.body,
			color: BRAND.steel,
			lineSpacingMultiple: 1.4,
		},
	);

	slide.addNotes("If the board only remembers one slide, make it this one. Do not defend the payback number yet.");
}

function addRevenueChart(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.data, sectionTitle: "Performance" });
	slideTitle(slide, "Revenue by segment", "Performance");

	slide.addChart(REVENUE_BY_SEGMENT, {
		x: MARGIN,
		y: 1.6,
		w: 8.3,
		h: 5.1,
		type: ChartType.bar,
		barDir: "col",
		barGrouping: "stacked",
		barGapWidthPct: 65,
		chartColors: CHART_COLORS,
		showLegend: true,
		legendPos: "b",
		legendColor: BRAND.slate,
		legendFontSize: 11,
		showValue: true,
		dataLabelColor: BRAND.white,
		dataLabelFontSize: 10,
		dataLabelFormatCode: "$0.0",
		catAxisLabelColor: BRAND.slate,
		catAxisLabelFontSize: 11,
		catAxisLineShow: false,
		valAxisLabelColor: BRAND.slate,
		valAxisLabelFontSize: 11,
		valAxisLabelFormatCode: '$0"M"',
		valAxisLineShow: false,
		valGridLine: { color: BRAND.mist, size: 1, style: "solid" },
	});

	const noteX = MARGIN + 8.3 + 0.45;
	const noteW = WIDE.w - noteX - MARGIN;
	slide.addText("WHAT MOVED", {
		x: noteX,
		y: 1.7,
		w: noteW,
		h: 0.3,
		margin: 0,
		fontFace: FONT.body,
		fontSize: TYPE.caption,
		color: BRAND.amber,
		bold: true,
		charSpacing: 1.6,
	});
	slide.addText(
		[
			{
				text: "Platform is the whole story.",
				options: { bold: true, color: BRAND.ink, breakLine: true, fontSize: TYPE.body },
			},
			{
				text: "It added $3.2M this quarter — more than Services and Licensing combined have added all year.",
				options: { color: BRAND.slate, breakLine: true, fontSize: TYPE.body },
			},
			{ text: "\n", options: { fontSize: 8, breakLine: true } },
			{ text: "Licensing turned over.", options: { bold: true, color: BRAND.ink, breakLine: true, fontSize: TYPE.body } },
			{
				text: "First down quarter in nine. Two renewals slipped into Q4; neither is at risk, but the trend line is now flat and we should stop planning against growth here.",
				options: { color: BRAND.slate, fontSize: TYPE.body },
			},
		],
		{
			x: noteX,
			y: 2.1,
			w: noteW,
			h: 4.2,
			margin: 0,
			fontFace: FONT.body,
			lineSpacingMultiple: 1.35,
		},
	);

	slide.addNotes(
		`Totals by quarter: ${QUARTERS.map((q, i) => `${q} $${REVENUE_BY_SEGMENT.reduce((sum, s) => sum + s.values[i], 0).toFixed(1)}M`).join(", ")}.`,
	);
}

function addRevenueMix(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.data, sectionTitle: "Performance" });
	slideTitle(slide, "Where the $46.4M came from", "Performance");

	slide.addChart(REVENUE_MIX, {
		x: 0.4,
		y: 1.55,
		w: 6.2,
		h: 5.2,
		type: ChartType.doughnut,
		holeSize: 62,
		chartColors: CHART_COLORS,
		showLegend: false,
		showValue: true,
		showPercent: true,
		dataLabelColor: BRAND.white,
		dataLabelFontSize: 12,
		dataLabelPosition: "ctr",
	});
	// The doughnut hole is a real hole, so a text box centred on the chart shows through it.
	slide.addText(
		[
			{ text: "$46.4M\n", options: { fontSize: 26, color: BRAND.ink, fontFace: FONT.head } },
			{ text: "total Q3 revenue", options: { fontSize: 11, color: BRAND.slate } },
		],
		{
			x: 2.35,
			y: 3.65,
			w: 2.3,
			h: 1,
			margin: 0,
			align: "center",
			valign: "middle",
			fontFace: FONT.body,
		},
	);

	const { x: colX, w: colW } = columns(1, { left: 7.05, width: WIDE.w - 7.05 - MARGIN });
	SEGMENTS.forEach(({ label, value, note }, i) => {
		const y = 1.95 + i * 1.45;
		slide.addGroup([
			{ rect: { x: colX[0], y, w: 0.08, h: 1.05, fill: { color: CHART_COLORS[i] } } },
			{
				text: {
					text: label,
					options: {
						x: colX[0] + 0.28,
						y,
						w: colW - 0.28,
						h: 0.36,
						margin: 0,
						valign: "middle",
						fontFace: FONT.head,
						fontSize: TYPE.heading,
						color: BRAND.ink,
					},
				},
			},
			{
				text: {
					text: `$${value.toFixed(1)}M  ·  ${((value / TOTAL_Q3) * 100).toFixed(0)}% of revenue`,
					options: {
						x: colX[0] + 0.28,
						y: y + 0.38,
						w: colW - 0.28,
						h: 0.32,
						margin: 0,
						valign: "middle",
						fontFace: FONT.body,
						fontSize: TYPE.body,
						color: BRAND.slate,
					},
				},
			},
			{
				text: {
					text: note,
					options: {
						x: colX[0] + 0.28,
						y: y + 0.7,
						w: colW - 0.28,
						h: 0.32,
						margin: 0,
						valign: "middle",
						fontFace: FONT.body,
						fontSize: TYPE.caption,
						color: BRAND.steel,
						italic: true,
					},
				},
			},
		]);
	});

	slide.addNotes("Platform crossing 55% of revenue is the number that changes how we should be valued. Say it out loud.");
}

function addRegionTable(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.data, sectionTitle: "Operations" });
	slideTitle(slide, "Regional performance", "Operations");

	const headerCell = {
		fill: { color: BRAND.navy },
		color: BRAND.white,
		bold: true,
		fontSize: 12,
		fontFace: FONT.body,
		valign: "middle",
		margin: [0.12, 0.16, 0.12, 0.16],
	};
	const bodyCell = (i) => ({
		fill: { color: i % 2 === 0 ? BRAND.white : BRAND.paper },
		color: BRAND.ink,
		fontSize: 13,
		fontFace: FONT.body,
		valign: "middle",
		margin: [0.12, 0.16, 0.12, 0.16],
		border: [{ type: "none" }, { type: "none" }, { type: "solid", color: BRAND.mist, width: 1 }, { type: "none" }],
	});

	const header = ["Region", "Revenue", "QoQ growth", "New logos", "Headcount"].map((text, i) => ({
		text,
		options: { ...headerCell, align: i === 0 ? "left" : "right" },
	}));

	const rows = REGIONS.map((r, i) => [
		{ text: r.region, options: { ...bodyCell(i), bold: true } },
		{ text: `$${r.revenue.toFixed(1)}M`, options: { ...bodyCell(i), align: "right" } },
		{
			text: signedPct(r.growth),
			options: { ...bodyCell(i), align: "right", bold: true, color: r.growth >= 0 ? BRAND.teal : BRAND.rose },
		},
		{ text: `${r.logos}%`, options: { ...bodyCell(i), align: "right" } },
		{ text: String(r.headcount), options: { ...bodyCell(i), align: "right" } },
	]);

	const totals = {
		revenue: REGIONS.reduce((s, r) => s + r.revenue, 0),
		logos: REGIONS.reduce((s, r) => s + r.logos, 0),
		headcount: REGIONS.reduce((s, r) => s + r.headcount, 0),
	};
	const totalCell = {
		fill: { color: BRAND.mist },
		color: BRAND.ink,
		bold: true,
		fontSize: 13,
		fontFace: FONT.body,
		valign: "middle",
		margin: [0.12, 0.16, 0.12, 0.16],
	};
	rows.push([
		{ text: "Total", options: totalCell },
		{ text: `$${totals.revenue.toFixed(1)}M`, options: { ...totalCell, align: "right" } },
		{ text: signedPct(13.7), options: { ...totalCell, align: "right", color: BRAND.teal } },
		{ text: `${totals.logos}%`, options: { ...totalCell, align: "right" } },
		{ text: String(totals.headcount), options: { ...totalCell, align: "right" } },
	]);

	slide.addTable([header, ...rows], {
		x: MARGIN,
		y: 1.75,
		w: CONTENT_W,
		colW: [4.3, 2.0, 2.1, 1.9, 1.53],
		rowH: 0.52,
		autoPage: false,
	});

	slide.addText(
		"LATAM is the only region shrinking, on a base small enough that one churned account explains all of it. " + "We are not reacting to it this quarter.",
		{
			x: MARGIN,
			y: 5.4,
			w: CONTENT_W,
			h: 0.7,
			margin: 0,
			fontFace: FONT.body,
			fontSize: TYPE.body,
			color: BRAND.slate,
			lineSpacingMultiple: 1.35,
		},
	);

	slide.addNotes("If asked about LATAM: the churned account was Grupo Alvear, non-renewal on a 2023 pilot contract.");
}

function addRetentionChart(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.data, sectionTitle: "Operations" });
	slideTitle(slide, "Retention held while we grew", "Operations");

	slide.addChart(RETENTION, {
		x: MARGIN,
		y: 1.7,
		w: 8.6,
		h: 4.9,
		type: ChartType.line,
		chartColors: [BRAND.teal, BRAND.navy],
		lineSize: 3,
		lineDataSymbol: "circle",
		lineDataSymbolSize: 9,
		lineSmooth: false,
		showLegend: true,
		legendPos: "b",
		legendColor: BRAND.slate,
		legendFontSize: 11,
		catAxisLabelColor: BRAND.slate,
		catAxisLabelFontSize: 12,
		catAxisLineShow: false,
		valAxisLabelColor: BRAND.slate,
		valAxisLabelFontSize: 11,
		valAxisMinVal: 85,
		valAxisMaxVal: 120,
		valAxisMajorUnit: 5,
		valAxisLabelFormatCode: '0"%"',
		valAxisLineShow: false,
		valGridLine: { color: BRAND.mist, size: 1, style: "solid" },
	});

	const calloutX = MARGIN + 8.6 + 0.4;
	slide.addText(
		[
			{ text: "114%\n", options: { fontSize: 34, color: BRAND.teal, fontFace: FONT.head, breakLine: true } },
			{
				text: "net revenue retention in September — the highest since we started measuring it, and it happened in the same quarter we onboarded the most new logos.",
				options: { fontSize: TYPE.body, color: BRAND.steel },
			},
		],
		{
			x: calloutX,
			y: 1.9,
			w: WIDE.w - calloutX - MARGIN,
			h: 2.6,
			margin: 0.18,
			shape: ShapeType.roundRect,
			rectRadius: 0.1,
			fill: { color: BRAND.teal, transparency: 92 },
			line: { color: BRAND.teal, width: 1 },
			fontFace: FONT.body,
			lineSpacingMultiple: 1.3,
		},
	);

	slide.addNotes("Gross retention at 96% is the one to watch — it is the number that would tell us the product is slipping.");
}

function addRoadmap(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.content, sectionTitle: "Operations" });
	slideTitle(slide, "What we are committing to", "Next quarter");

	const { x: xs, w } = columns(ROADMAP.length, { gap: 0.3 });
	ROADMAP.forEach((item, i) => {
		const accent = [BRAND.teal, BRAND.navy, BRAND.amber, BRAND.violet][i];
		slide.addGroup([
			{
				shape: {
					type: ShapeType.chevron,
					options: { x: xs[i], y: 1.85, w, h: 0.62, fill: { color: accent }, line: { color: accent } },
				},
			},
			{
				text: {
					text: item.when.toUpperCase(),
					options: {
						x: xs[i],
						y: 1.85,
						w: w - 0.25,
						h: 0.62,
						margin: 0,
						align: "center",
						valign: "middle",
						fontFace: FONT.head,
						fontSize: 13,
						color: BRAND.white,
						charSpacing: 2,
					},
				},
			},
			{
				roundRect: {
					x: xs[i],
					y: 2.65,
					w,
					h: 2.5,
					rectRadius: 0.06,
					fill: { color: BRAND.white },
					line: { color: BRAND.mist, width: 1 },
				},
			},
			{
				text: {
					text: item.title,
					options: {
						x: xs[i] + 0.22,
						y: 2.9,
						w: w - 0.44,
						h: 0.9,
						margin: 0,
						fontFace: FONT.head,
						fontSize: TYPE.subhead,
						color: BRAND.ink,
						lineSpacingMultiple: 1.2,
					},
				},
			},
			{
				text: {
					text: item.note,
					options: {
						x: xs[i] + 0.22,
						y: 3.85,
						w: w - 0.44,
						h: 1.1,
						margin: 0,
						fontFace: FONT.body,
						fontSize: TYPE.body,
						color: BRAND.slate,
						lineSpacingMultiple: 1.35,
					},
				},
			},
		]);
	});

	slide.addText(
		[
			{ text: "And what we are dropping: ", options: { bold: true, color: BRAND.ink } },
			{
				text: "the in-app analytics rebuild slips to FY27. It is the only way the four above land on time, and it is a better trade than shipping five things late.",
				options: { color: BRAND.slate },
			},
		],
		{
			x: MARGIN,
			y: 5.55,
			w: CONTENT_W,
			h: 0.8,
			margin: 0.18,
			shape: ShapeType.roundRect,
			rectRadius: 0.08,
			fill: { color: BRAND.amber, transparency: 90 },
			fontFace: FONT.body,
			fontSize: TYPE.body,
			lineSpacingMultiple: 1.35,
		},
	);

	slide.addNotes("Expect pushback on dropping the analytics rebuild. The trade is four on time versus five late.");
}

function addClosing(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.closing, sectionTitle: "Closing" });

	slide.addText("Questions", {
		x: 0.95,
		y: 2.7,
		w: 9,
		h: 1,
		margin: 0,
		valign: "middle",
		fontFace: FONT.head,
		fontSize: TYPE.display,
		color: BRAND.white,
	});
	slide.addShape(ShapeType.rect, { x: 0.95, y: 3.95, w: 2.4, h: 0.06, fill: { color: BRAND.teal } });
	slide.addText(
		[
			{ text: "Meridian Group  ·  Office of the CFO\n", options: { color: BRAND.mist, breakLine: true } },
			{ text: "Full data appendix and the Q3 model are in the board folder.", options: { color: BRAND.slate } },
		],
		{
			x: 0.95,
			y: 4.35,
			w: 8,
			h: 0.9,
			margin: 0,
			fontFace: FONT.body,
			fontSize: TYPE.body,
			lineSpacingMultiple: 1.4,
		},
	);

	slide.addNotes("Close by restating the one decision needed today: approve or defer the Q4 EMEA hiring plan.");
}

/** Build the deck and write it to `outFile`. */
export async function build(outFile) {
	const pptx = new TsPptx();

	pptx.title = "Meridian Group — Q3 FY26 Business Review";
	pptx.subject = "Quarterly business review";
	pptx.author = "Meridian Group";
	pptx.company = "Meridian Group";

	applyDesign(pptx);
	for (const title of SECTIONS) pptx.addSection({ title });

	addCover(pptx);
	addAgenda(pptx);
	addSectionDivider(pptx, "01", "Performance", "How the quarter landed against the plan we set in June.");
	addKpiBand(pptx);
	addRevenueChart(pptx);
	addRevenueMix(pptx);
	addSectionDivider(pptx, "02", "Operations", "What it cost to get there, region by region.");
	addRegionTable(pptx);
	addRetentionChart(pptx);
	addRoadmap(pptx);
	addClosing(pptx);

	return await pptx.writeFile({ fileName: outFile });
}

export const showcase = {
	slug: "quarterly-review",
	title: "Meridian Q3 FY26 Business Review",
	description: "Corporate flagship: themed masters, KPI cards, charts, a styled table, and speaker notes.",
	fileName: "Meridian_Q3_Business_Review.pptx",
	build,
};
