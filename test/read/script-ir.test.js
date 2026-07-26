// The deck IR — `ts-pptx/script`'s read half.
//
// Two kinds of assertion live here, and the split matters.
//
// The *invariants* run over the whole fixture corpus: every deck converts without
// throwing, the result is JSON-serializable, it holds no `undefined`, and converting twice
// gives the same answer. Those are the properties the IR's contract rests on — an
// `undefined` anywhere would give "absent" two spellings and make two IRs compare unequal
// for no reason, which is exactly the phantom failure a round-trip check must not produce.
//
// The *specifics* pin conversions that are easy to get wrong and silent when wrong:
// path-unit scaling, group child spaces, and the three constructs the read model cannot
// see at all (theme line width, embedded media, equations). Each of those was a real bug
// caught by running this against the corpus rather than by reading the types.

import { readFile, readdir } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import { Presentation, isAutoShape } from '../../dist/read.js'
import { readModelToIr } from '../../dist/script.js'
import { assert, assertEqual } from '../helpers.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const FIXTURES = path.join(__dirname, 'fixtures')

const fixtureNames = (await readdir(FIXTURES)).filter((name) => name.endsWith('.pptx')).sort()

async function irFor(name) {
	return readModelToIr(await Presentation.load(await readFile(path.join(FIXTURES, name))))
}

/** Every call across every slide, flattened. */
function allCalls(ir) {
	return ir.slides.flatMap((slide) => slide.calls)
}

/** Note constructs recorded anywhere in the deck. */
function constructs(ir) {
	return new Set(ir.fidelity.map((note) => note.construct))
}

/** Walk every value in the IR, so an invariant can be asserted over all of them. */
function* walk(value, trail = '$') {
	yield [trail, value]
	if (Array.isArray(value)) {
		for (const [index, item] of value.entries()) yield* walk(item, `${trail}[${index}]`)
	} else if (value && typeof value === 'object' && !(value instanceof Uint8Array)) {
		for (const [key, item] of Object.entries(value)) yield* walk(item, `${trail}.${key}`)
	}
}

describe('deck IR — corpus invariants', () => {
	test('every fixture converts, and produces slides matching the source', async () => {
		for (const name of fixtureNames) {
			const source = await Presentation.load(await readFile(path.join(FIXTURES, name)))
			const ir = await irFor(name)
			assertEqual(ir.slides.length, source.slides.length, `${name}: slide count`)
			assertEqual(
				ir.slides.map((slide) => slide.number).join(','),
				source.slides.map((_, index) => index + 1).join(','),
				`${name}: slide numbers are 1-based and contiguous`
			)
		}
	})

	test('the IR is JSON-serializable, so it can be diffed and cached', async () => {
		for (const name of fixtureNames) {
			const ir = await irFor(name)
			// Asset bytes are the one non-JSON member and are carried out-of-band by design.
			const { assets, ...rest } = ir
			JSON.parse(JSON.stringify(rest))
			for (const asset of assets) {
				assert(asset.bytes instanceof Uint8Array, `${name}: asset ${asset.name} carries raw bytes`)
				assert(asset.name.length > 0 && asset.contentType.length > 0, `${name}: asset is named and typed`)
			}
		}
	})

	test('no value is `undefined`, so "absent" has exactly one spelling', async () => {
		for (const name of fixtureNames) {
			const ir = await irFor(name)
			for (const [trail, value] of walk(ir)) {
				assert(value !== undefined, `${name}: ${trail} is undefined; absent keys must be omitted, not set to undefined`)
			}
		}
	})

	test('conversion is deterministic — the same deck twice gives the same IR', async () => {
		// Not a formality: asset names are assigned in first-reference order, so any
		// nondeterminism in the walk would surface here as renamed images.
		for (const name of fixtureNames) {
			const [first, second] = [await irFor(name), await irFor(name)]
			const strip = (ir) =>
				JSON.stringify({ ...ir, assets: ir.assets.map((a) => [a.name, a.contentType, a.bytes.length]) })
			assertEqual(strip(second), strip(first), `${name}: two conversions differ`)
		}
	})

	test('every call names a real write-API method and carries its arguments', async () => {
		const arity = { addText: 2, addShape: 2, addImage: 1, addTable: 2, addChart: 3, addConnector: 1, addGroup: 2 }
		for (const name of fixtureNames) {
			for (const call of allCalls(await irFor(name))) {
				assert(call.method in arity, `${name}: unknown method ${call.method}`)
				assertEqual(call.args.length, arity[call.method], `${name}: ${call.method} argument count`)
			}
		}
	})

	test('every fidelity note is fully populated and names a cause', async () => {
		for (const name of fixtureNames) {
			for (const note of (await irFor(name)).fidelity) {
				assert(
					['dropped', 'flattened', 'approximated'].includes(note.disposition),
					`${name}: ${note.construct} has disposition ${note.disposition}`
				)
				assert(['unread', 'unwritable', 'unsupported'].includes(note.cause), `${name}: ${note.construct} cause`)
				assert(/^[a-z][A-Za-z]*(\.[a-zA-Z]+)+$/.test(note.construct), `${name}: ${note.construct} is not a dotted key`)
				assert(note.detail.length > 20, `${name}: ${note.construct} detail is too thin to act on`)
			}
		}
	})
})

describe('deck IR — geometry', () => {
	test('position is raw EMU, so it survives the round-trip exactly', async () => {
		const ir = await irFor('mixed.pptx')
		const positioned = allCalls(ir)
			.map((call) => call.args.find((arg) => arg && typeof arg === 'object' && 'x' in arg))
			.filter(Boolean)
		assert(positioned.length > 0, 'mixed.pptx should produce positioned calls')
		for (const options of positioned) {
			for (const key of ['x', 'y', 'w', 'h']) {
				if (options[key] === undefined) continue
				assert(
					/^-?\d+emu$/.test(options[key]),
					`${key} should be an exact EMU string, got ${JSON.stringify(options[key])}`
				)
			}
		}
	})

	test('custGeom path points land exactly where the source path puts them', async () => {
		const source = await Presentation.load(await readFile(path.join(FIXTURES, 'custgeom.pptx')))
		const ir = await irFor('custgeom.pptx')
		const shapes = allCalls(ir).filter((call) => call.args[0] === 'custGeom')
		assert(shapes.length > 0, 'custgeom.pptx should produce custGeom shapes')

		for (const shape of shapes) {
			const origin = source.slides
				.flatMap((slide) => slide.shapes)
				.filter(isAutoShape)
				.find((candidate) => candidate.name === shape.sourceName)
			const [firstPath] = origin.customGeometry.paths
			const frame = origin.absoluteFrame
			const expected = firstPath.commands
				.filter((command) => 'x' in command)
				.map((command) => Math.round((command.x * frame.width) / firstPath.w))
			const actual = shape.args[1].points
				.filter((point) => point.x !== undefined)
				.slice(0, expected.length)
				.map((point) => Number(String(point.x).replace('emu', '')))
			assertEqual(actual.join(','), expected.join(','), `${shape.sourceName}: path x coordinates`)
		}
	})

	test('a path viewport that differs from the shape box is scaled, not passed through', async () => {
		// The corpus cannot exercise this on its own: every fixture path — and every path
		// ts-pptx itself writes — sets `a:path/@w` to the shape width, so the scale factor is
		// 1 and a version that skipped scaling entirely would still pass. Halving the
		// viewport through the documented raw hatch produces the case that tells them apart.
		const deck = await Presentation.load(await readFile(path.join(FIXTURES, 'custgeom.pptx')))
		const shape = deck.slides
			.flatMap((slide) => slide.shapes)
			.filter(isAutoShape)
			.find((candidate) => candidate.customGeometry)
		const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
		const pathEl = shape.element_.getElementsByTagNameNS(A_NS, 'path').item(0)
		const original = Number(pathEl.getAttribute('w'))
		pathEl.setAttribute('w', String(original * 2))
		shape.markDirty()

		const rescaled = readModelToIr(await Presentation.load(await deck.save()))
		const call = allCalls(rescaled).find((candidate) => candidate.sourceName === shape.name)
		const baseline = allCalls(await irFor('custgeom.pptx')).find((candidate) => candidate.sourceName === shape.name)

		const spanOf = (one) => {
			const xs = one.args[1].points
				.filter((point) => point.x !== undefined)
				.map((p) => Number(String(p.x).replace('emu', '')))
			return Math.max(...xs) - Math.min(...xs)
		}
		// Twice the viewport over the same box means every coordinate maps to half the width.
		assertEqual(spanOf(call), Math.round(spanOf(baseline) / 2), 'doubling a:path/@w should halve the emitted span')
	})
})

describe('deck IR — groups', () => {
	test('a nested group nests rather than flattening', async () => {
		const ir = await irFor('group-transform.pptx')
		const nested = allCalls(ir).find((call) => call.sourceName === 'nested-rot-in-scale')
		assert(nested, 'group-transform.pptx has a nested-rot-in-scale group')
		const kinds = nested.args[0].map((child) => Object.keys(child)[0])
		assert(kinds.includes('group'), `nested group should emit a group child, got ${kinds.join(',')}`)
	})

	test('group children use the key-tagged GroupChildProps shape', async () => {
		const allowed = new Set(['image', 'line', 'rect', 'roundRect', 'shape', 'text', 'group'])
		for (const name of fixtureNames) {
			for (const call of allCalls(await irFor(name))) {
				if (call.method !== 'addGroup') continue
				for (const child of call.args[0]) {
					const keys = Object.keys(child)
					assertEqual(keys.length, 1, `${name}: a group child must be a single-key descriptor`)
					assert(allowed.has(keys[0]), `${name}: ${keys[0]} is not a GroupChildProps variant`)
				}
			}
		}
	})

	test('only a group that actually scales its children is noted', async () => {
		// A rotation or flip leaves the child space alone; only a differing chOff/chExt
		// changes it. The fixture names say which is which, so this is checkable.
		const ir = await irFor('group-transform.pptx')
		const noted = new Set(
			ir.fidelity.filter((note) => note.construct === 'group.childSpace').map((note) => note.shapeName)
		)
		assert(noted.has('scale-rot'), 'a scaling group should be noted')
		assert(!noted.has('rot30'), 'a pure rotation does not change the child space and must not be noted')
		assert(!noted.has('flipH'), 'a pure flip does not change the child space and must not be noted')
	})
})

describe('deck IR — losses the read model cannot see', () => {
	// These three are the reason fidelity notes exist. In each case the construct is
	// invisible to every read accessor, so without an explicit check the conversion looks
	// clean while silently changing the deck.

	test('a theme-referenced outline reports its lost width', async () => {
		const ir = await irFor('mixed.pptx')
		const noted = ir.fidelity.filter((note) => note.construct === 'line.width')
		assert(noted.length > 0, 'mixed.pptx has shapes styled from the theme line list')
		assertEqual(noted[0].cause, 'unread', 'the cause is a missing accessor, not a missing write option')
	})

	test('embedded audio/video is reported, not silently emitted as a still image', async () => {
		const ir = await irFor('av-media.pptx')
		assert(constructs(ir).has('media.audioVideo'), 'av-media.pptx should report its media as flattened')
		// It still emits something, so the slide is not left blank.
		assert(allCalls(ir).length > 0, 'the poster frame should still be emitted')
	})

	test('an OMML equation is reported rather than emitted as an empty box', async () => {
		const ir = await irFor('math-omml.pptx')
		const noted = ir.fidelity.find((note) => note.construct === 'text.equation')
		assert(noted, 'math-omml.pptx holds an equation no accessor exposes')
		assertEqual(noted.cause, 'unread', 'the write API can author equations; the read side cannot see them')
	})
})

describe('deck IR — slide sourcing', () => {
	test('a slide is authored unless something on it has no write-API expression', async () => {
		for (const name of fixtureNames) {
			const ir = await irFor(name)
			for (const slide of ir.slides) {
				if (slide.source === 'authored') continue
				assertEqual(slide.calls.length, 0, `${name}: a carried slide must emit no calls`)
				assert(
					ir.fidelity.some((note) => note.slideNumber === slide.number && note.construct === 'slide.carried'),
					`${name}: slide ${slide.number} is carried without a note saying why`
				)
			}
		}
	})
})
