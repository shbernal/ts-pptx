import { assert, assertEqual, captureDiagnostics, defineRegressionSuite, slideXml } from '../../helpers.js'

// Which fill kind a props object asks for used to be answered in seven places, and they
// disagreed. `fill: { gradient }` painted a black `<a:solidFill>` and warned that the caller's
// (absent) colour string was invalid; `fill: { pattern }` did the same; a background authored as
// `{ gradient }` or `{ pattern }` emitted no `<p:bg>` at all. Meanwhile the stroke emitter
// inferred `gradient` from its sub-object and three `define/` modules each carried their own
// copy of the image half, so the same spelling behaved differently depending on where it landed.
//
// `resolveFillKind` is the one answer now, and these pin it from every entry point: the shape
// interior, the stroke (through both of the two `define/` rebuilds that normalise a line), and
// the slide background. The rule they encode is that a sub-object selects its kind on its own,
// and an explicit `type` beats a sub-object that disagrees.

const BOX = { x: 1, y: 1, w: 2, h: 1 }
const GRADIENT = {
	kind: 'linear',
	angle: 0,
	stops: [
		{ position: 0, color: '0088CC' },
		{ position: 100, color: 'FF0000' },
	],
}
const PATTERN = { preset: 'diagCross', fgColor: '003366', bgColor: 'FFFFFF' }
/** A 1x1 transparent PNG — the smallest thing that resolves to a real media relationship. */
const PNG =
	'image/png;base64,' +
	'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

/** The fill-group elements inside `block`, in document order, by local name. */
function fillKinds(block) {
	return [...block.matchAll(/<a:(solidFill|gradFill|pattFill|blipFill|noFill)\b/g)].map((m) => m[1])
}

/** The `<p:spPr>` of the part's first shape — its interior fill, without the `<a:ln>` stroke. */
function shapeFill(xml) {
	const spPr = xml.slice(xml.indexOf('<p:spPr>'), xml.indexOf('<a:ln'))
	assert(spPr, 'expected a <p:spPr> in:\n' + xml)
	return fillKinds(spPr)
}

/** The `<a:ln>` of the part's first shape. */
function lineFill(xml) {
	const open = xml.indexOf('<a:ln')
	assert(open >= 0, 'expected an <a:ln> in:\n' + xml)
	return fillKinds(xml.slice(open, xml.indexOf('</a:ln>', open)))
}

function backgroundBlock(xml) {
	const open = xml.indexOf('<p:bg>')
	return open < 0 ? null : xml.slice(open, xml.indexOf('</p:bg>') + '</p:bg>'.length)
}

defineRegressionSuite('Fill kind inference', [
	{
		name: 'a fill sub-object selects its kind without a `type` (gradient, pattern, image)',
		fn: async () => {
			for (const [label, fill, expected] of [
				['gradient', { gradient: GRADIENT }, 'gradFill'],
				['pattern', { pattern: PATTERN }, 'pattFill'],
				['image', { image: { data: PNG } }, 'blipFill'],
			]) {
				// The bug this pins is specifically silent: the shape came out black, and the only
				// diagnostic named a colour the caller never wrote. So assert on both.
				const { result: xml, codes } = await captureDiagnostics(() =>
					slideXml((p) => p.addSlide().addShape('rect', { ...BOX, fill }))
				)
				assertEqual(shapeFill(xml).join(','), expected, `fill: { ${label} } emits <a:${expected}>`)
				assertEqual(codes.join(','), '', `fill: { ${label} } reports nothing`)
			}
		},
	},
	{
		name: 'a fill sub-object and its explicit `type` spelling emit the same bytes',
		fn: async () => {
			for (const [type, sub] of /** @type {[string, Record<string, unknown>][]} */ ([
				['gradient', { gradient: GRADIENT }],
				['pattern', { pattern: PATTERN }],
				['image', { image: { data: PNG } }],
			])) {
				const inferred = await slideXml((p) => p.addSlide().addShape('rect', { ...BOX, fill: sub }))
				const explicit = await slideXml((p) => p.addSlide().addShape('rect', { ...BOX, fill: { type, ...sub } }))
				assertEqual(inferred, explicit, `fill: { ${type} } matches fill: { type: '${type}', … }`)
			}
		},
	},
	{
		name: 'an explicit `type` wins over a sub-object that disagrees',
		fn: async () => {
			// The rule exists so that `type: 'none'` can still mean transparent when a gradient is
			// left on the object — the case where losing the argument silently repaints the shape.
			const none = await slideXml((p) =>
				p.addSlide().addShape('rect', { ...BOX, fill: { type: 'none', gradient: GRADIENT } })
			)
			assertEqual(shapeFill(none).join(','), 'noFill', "type: 'none' beats a gradient sub-object")

			const solid = await slideXml((p) =>
				p.addSlide().addShape('rect', { ...BOX, fill: { type: 'solid', color: '00FF00', gradient: GRADIENT } })
			)
			assertEqual(shapeFill(solid).join(','), 'solidFill', "type: 'solid' beats a gradient sub-object")
			assert(solid.includes('<a:srgbClr val="00FF00"/>'), `expected the solid colour; got:\n${solid}`)
		},
	},
	{
		name: 'a line sub-object selects its stroke kind, and a bare line still inherits its colour',
		fn: async () => {
			for (const [label, line, expected] of [
				['gradient', { width: 2, gradient: GRADIENT }, 'gradFill'],
				['pattern', { width: 2, pattern: PATTERN }, 'pattFill'],
				['color', { width: 2, color: 'FF0000' }, 'solidFill'],
				// No paint named at all: the default line colour is applied at define time, so this
				// stays a solid stroke rather than becoming an inherited one.
				['nothing', { width: 2 }, 'solidFill'],
				['none', { width: 2, type: 'none' }, 'noFill'],
			]) {
				const xml = await slideXml((p) => p.addSlide().addShape('rect', { ...BOX, line }))
				assertEqual(lineFill(xml).join(','), expected, `line: { ${label} } emits <a:${expected}>`)
			}
		},
	},
	{
		name: 'a picture stroke is refused rather than painted as nothing',
		fn: async () => {
			// `<a:ln>`'s paint child is EG_LineFillProperties (noFill/solidFill/gradFill/pattFill),
			// so a bitmap stroke has no OOXML expression at all. `ShapeLineProps` subtracts `image`
			// from the fill props it inherits, and this is the runtime half of that subtraction:
			// the spelling was reachable from JS, registered no media, and painted nothing while
			// warning about a rel it was never going to be given.
			for (const line of [
				{ width: 2, image: { data: PNG } },
				{ width: 2, type: 'image', image: { data: PNG } },
			]) {
				let err
				try {
					await slideXml((p) => p.addSlide().addShape('rect', { ...BOX, line }))
				} catch (e) {
					err = e
				}
				assertEqual(err?.code, 'line/image-fill-unsupported', `expected a refusal for line: ${JSON.stringify(line)}`)
			}
			// The interior is unaffected: a picture fill is a shape fill, and still paints.
			const xml = await slideXml((p) => p.addSlide().addShape('rect', { ...BOX, fill: { image: { data: PNG } } }))
			assertEqual(shapeFill(xml).join(','), 'blipFill', 'a picture *fill* still paints')
		},
	},
	{
		name: "a line `type: 'inherit'` emits no paint child at all",
		fn: async () => {
			const xml = await slideXml((p) => p.addSlide().addShape('rect', { ...BOX, line: { width: 2, type: 'inherit' } }))
			assertEqual(lineFill(xml).join(','), '', 'an inherited stroke names no paint')
		},
	},
	{
		name: 'the addText `shape: line` rebuild resolves the same kinds as the addShape one',
		fn: async () => {
			// define/text.ts carries a second, near-duplicate ShapeLineProps rebuild. It stamped
			// `type: 'solid'` on unconditionally and defaulted the colour with it, so a pattern
			// stroke on this path came out a default-black solid line.
			for (const [label, line, expected] of [
				['gradient', { width: 2, gradient: GRADIENT }, 'gradFill'],
				['pattern', { width: 2, pattern: PATTERN }, 'pattFill'],
			]) {
				const xml = await slideXml((p) => p.addSlide().addText('x', { shape: 'line', x: 1, y: 1, w: 4, h: 0, line }))
				assertEqual(lineFill(xml).join(','), expected, `addText line: { ${label} } emits <a:${expected}>`)
			}
		},
	},
	{
		name: 'a background sub-object selects its kind, and an empty background still emits no <p:bg>',
		fn: async () => {
			for (const [label, background, expected] of [
				['gradient', { gradient: GRADIENT }, 'gradFill'],
				['pattern', { pattern: PATTERN }, 'pattFill'],
				['type + pattern', { type: 'pattern', pattern: PATTERN }, 'pattFill'],
				['color', { color: 'FF0000' }, 'solidFill'],
				['none', { type: 'none' }, 'noFill'],
			]) {
				const xml = await slideXml((p) => {
					p.addSlide().background = background
				})
				const bg = backgroundBlock(xml)
				assert(bg, `background: { ${label} } should emit a <p:bg>; got:\n${xml}`)
				assertEqual(fillKinds(bg).join(','), expected, `background: { ${label} } emits <a:${expected}>`)
			}

			// The two spellings of silence stay silent: `<p:bgPr>` requires a fill child, so
			// emitting the element for a background that names no paint would be invalid as well
			// as wrong.
			for (const [label, background] of [
				['empty', {}],
				['inherit', { type: 'inherit' }],
			]) {
				const xml = await slideXml((p) => {
					p.addSlide().background = background
				})
				assertEqual(backgroundBlock(xml), null, `background: { ${label} } emits no <p:bg>`)
			}
		},
	},
])
