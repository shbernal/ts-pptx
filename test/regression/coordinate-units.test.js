import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// Coordinates resolve through a single boundary (coordToEmu) with NO magnitude guessing:
// a bare number is ALWAYS inches; other units use an explicit string suffix. This replaces the
// old "a number >= 100 must already be EMU" heuristic that silently mis-rendered values near the
// threshold. Slide is the default LAYOUT_16x9 → 10in wide (9144000 EMU), 5.625in tall (5143500).
const IN = 914400 // EMU per inch

async function offExtFor(opts) {
	const { zip } = await build((p) => {
		p.addSlide().addShape('rect', { x: opts.x, y: opts.y, w: opts.w, h: opts.h })
	})
	const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
	const sp = xml.slice(xml.indexOf('<p:sp>'))
	const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(sp)
	const ext = /<a:ext cx="(-?\d+)" cy="(-?\d+)"\/>/.exec(sp)
	assert(off && ext, 'expected <a:off>/<a:ext> in shape; got: ' + sp.slice(0, 300))
	return { x: +off[1], y: +off[2], cx: +ext[1], cy: +ext[2] }
}

defineRegressionSuite('Coordinate units', [
	{
		name: 'bare number is inches',
		fn: async () => {
			const r = await offExtFor({ x: 1, y: 2, w: 3, h: 0.5 })
			assert(r.x === IN && r.y === 2 * IN && r.cx === 3 * IN && r.cy === 0.5 * IN, JSON.stringify(r))
		},
	},
	{
		name: 'percent resolves against the slide axis (width for x/w, height for y/h)',
		fn: async () => {
			const r = await offExtFor({ x: '50%', y: '20%', w: '25%', h: '40%' })
			// 50% of 10in width = 5in; 20% of 5.625in height; 25% width; 40% height
			assert(r.x === Math.round(0.5 * 10 * IN), `x ${r.x}`)
			assert(r.cx === Math.round(0.25 * 10 * IN), `cx ${r.cx}`)
			assert(r.y === Math.round(0.2 * 5.625 * IN), `y ${r.y}`)
			assert(r.cy === Math.round(0.4 * 5.625 * IN), `cy ${r.cy}`)
		},
	},
	{
		name: 'explicit "in" suffix equals a bare number',
		fn: async () => {
			const r = await offExtFor({ x: '1.5in', y: '2in', w: '3in', h: '1in' })
			assert(r.x === 1.5 * IN && r.y === 2 * IN && r.cx === 3 * IN && r.cy === IN, JSON.stringify(r))
		},
	},
	{
		name: 'points: 72pt == 1in',
		fn: async () => {
			const r = await offExtFor({ x: '72pt', y: '36pt', w: '144pt', h: '72pt' })
			assert(r.x === IN && r.y === IN / 2 && r.cx === 2 * IN && r.cy === IN, JSON.stringify(r))
		},
	},
	{
		name: 'pixels: 96px == 1in at the default 96 DPI',
		fn: async () => {
			const r = await offExtFor({ x: '96px', y: '48px', w: '960px', h: '192px' })
			assert(r.x === IN && r.y === IN / 2 && r.cx === 10 * IN && r.cy === 2 * IN, JSON.stringify(r))
		},
	},
	{
		name: 'raw EMU escape hatch passes through exactly',
		fn: async () => {
			const r = await offExtFor({ x: '914400emu', y: '457200emu', w: '1828800emu', h: '12700emu' })
			assert(r.x === 914400 && r.y === 457200 && r.cx === 1828800 && r.cy === 12700, JSON.stringify(r))
		},
	},
	{
		name: 'a bare number >= 100 is inches now (no EMU passthrough)',
		fn: async () => {
			// Under the old heuristic this 120 would have been emitted as 120 EMU (a collapsed dot).
			const r = await offExtFor({ x: 120, y: 1, w: 1, h: 1 })
			assert(r.x === 120 * IN, `expected ${120 * IN} (120in), got ${r.x}`)
		},
	},
])
