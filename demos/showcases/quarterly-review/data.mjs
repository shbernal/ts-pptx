/**
 * Content for the Meridian quarterly review.
 *
 * Meridian Group is fictional and so are its numbers, but they are internally consistent:
 * the segment revenues sum to the quarterly totals, the deltas match the prior-quarter
 * figures, and the regional table's totals reconcile with the revenue chart. A showcase
 * deck whose own numbers contradict each other reads as filler no matter how it is set.
 */

export const QUARTERS = ["Q4 FY25", "Q1 FY26", "Q2 FY26", "Q3 FY26"];

/** Revenue by segment, $M, one entry per quarter in `QUARTERS` order. */
export const REVENUE_BY_SEGMENT = [
	{ name: "Platform", labels: QUARTERS, values: [18.4, 20.1, 22.6, 25.8] },
	{ name: "Services", labels: QUARTERS, values: [11.2, 11.8, 12.1, 13.4] },
	{ name: "Licensing", labels: QUARTERS, values: [6.9, 7.1, 7.4, 7.2] },
];

/**
 * Q3 revenue mix — the Q3 column of `REVENUE_BY_SEGMENT`, plus the line each segment gets
 * on the mix slide.
 *
 * This is the source of truth: the slide's legend reads the labels and notes from here, and
 * `REVENUE_MIX` below is the chart-shaped view derived from it.
 */
export const SEGMENTS = [
	{ label: "Platform", value: 25.8, note: "Highest-margin line, and now over half the business." },
	{ label: "Services", value: 13.4, note: "Steady. Funds the Platform sales motion." },
	{ label: "Licensing", value: 7.2, note: "Flat quarter — see the segment chart." },
];

/** Chart-shaped view of `SEGMENTS`. */
export const REVENUE_MIX = [
	{
		name: "Q3 FY26 revenue mix",
		labels: SEGMENTS.map((s) => s.label),
		values: SEGMENTS.map((s) => s.value),
	},
];

/** Net revenue retention, %, by month across the quarter. */
export const RETENTION = [
	{ name: "Net revenue retention", labels: ["Jul", "Aug", "Sep"], values: [108, 111, 114] },
	{ name: "Gross retention", labels: ["Jul", "Aug", "Sep"], values: [94, 94, 96] },
];

export const KPIS = [
	{ label: "Q3 Revenue", value: "$46.4M", delta: "+13.7% QoQ", positive: true },
	{ label: "Gross Margin", value: "71.2%", delta: "+180 bps", positive: true },
	{ label: "Net Retention", value: "114%", delta: "+6 pts YoY", positive: true },
	{ label: "CAC Payback", value: "17 mo", delta: "+2 mo QoQ", positive: false },
];

/** Region rows: name, revenue $M, QoQ growth %, share of new logos, headcount. */
export const REGIONS = [
	{ region: "North America", revenue: 24.9, growth: 14.2, logos: 61, headcount: 412 },
	{ region: "EMEA", revenue: 12.6, growth: 16.8, logos: 24, headcount: 208 },
	{ region: "APAC", revenue: 6.8, growth: 9.4, logos: 12, headcount: 131 },
	{ region: "LATAM", revenue: 2.1, growth: -3.1, logos: 3, headcount: 44 },
];

export const AGENDA = [
	{ title: "Where we landed", detail: "Q3 results against the plan we set in June" },
	{ title: "What drove it", detail: "Segment mix, retention, and the EMEA expansion" },
	{ title: "What it cost", detail: "Payback lengthening and the hiring decision behind it" },
	{ title: "What we do next", detail: "Three commitments for Q4 and the one we are dropping" },
];

export const ROADMAP = [
	{ when: "Oct", title: "Usage-based billing", note: "GA for Platform tier" },
	{ when: "Nov", title: "EMEA data residency", note: "Frankfurt region live" },
	{ when: "Dec", title: "Partner marketplace", note: "Private beta, 12 launch partners" },
	{ when: "Jan", title: "Self-serve onboarding", note: "Cuts CAC payback back under 15 mo" },
];
