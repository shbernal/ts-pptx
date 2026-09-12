// Drawing ids stay unique, and connector bindings keep naming their shapes, when shapes are
// carried from one shape tree into another.
//
// `p:cNvPr/@id` is what an animation's `spid` and a connector's `a:stCxn`/`a:endCxn` name a shape
// by, so two shapes sharing an id make both ambiguous. `carryMasterGraphics` copied a layout's and
// a master's decorations onto the slide with their source ids, which routinely equal the slide's
// own. `importShape` renumbered a lifted subtree but left a connector binding inside it naming the
// old id, which on the host names a different shape.

import { describe, test } from 'vitest'
import { assert, assertEqual } from '../helpers.js'
import { openFixture } from './corpus.js'

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

/** Every `p:cNvPr/@id` under `root`, in document order. */
function drawingIds(root) {
	return [...root.getElementsByTagNameNS(P_NS, 'cNvPr')].map((el) => el.getAttribute('id'))
}

/**
 * Every connector binding under `root`, in document order, as the name of the shape inside `root`
 * its id names, or `null` when the id names nothing inside `root`.
 */
function bindingTargets(root) {
	const nameById = new Map(
		[...root.getElementsByTagNameNS(P_NS, 'cNvPr')].map((el) => [el.getAttribute('id'), el.getAttribute('name')])
	)
	return [...root.getElementsByTagName('*')]
		.filter((el) => el.namespaceURI === A_NS && (el.localName === 'stCxn' || el.localName === 'endCxn'))
		.map((el) => nameById.get(el.getAttribute('id')) ?? null)
}

describe('drawing ids across carried and imported shapes', () => {
	test.for(['gradient-fill', 'preset-geometry', 'theme-colors'])(
		'carryMasterGraphics leaves every drawing id on the %s slide unique',
		async (name) => {
			const target = await openFixture('empty')
			const source = await openFixture(name)
			const ownIds = drawingIds(source.slides[0].part.dom.documentElement)
			const imported = target.importSlide(source, 0, { theme: 'preserve', carryMasterGraphics: true })
			const ids = drawingIds(imported.part.dom.documentElement)
			assert(ids.length > ownIds.length, `decorations were carried onto the slide (${ids.length} ids, ${ownIds.length} of its own)`) // prettier-ignore
			const repeated = [...new Set(ids.filter((id, i) => ids.indexOf(id) !== i))]
			assertEqual(repeated.join(','), '', 'no drawing id is used twice')
		}
	)

	test('importShape keeps a grouped connector bound to the shape inside its group', async () => {
		// `mixed.pptx` slide 5, shape 4: a PowerPoint-authored group holding a rectangle and two
		// connectors whose `a:endCxn` both name that rectangle. Their `a:stCxn` name shapes outside
		// the group, which the import does not carry; those are covered by the next case.
		const source = await openFixture('mixed')
		const sourceGroup = source.slides[4].shapes[4].element_
		const inside = bindingTargets(sourceGroup).filter((name) => name !== null)
		assertEqual(inside.join(), 'Rectangle 3,Rectangle 3', 'the fixture group binds two connectors inside itself')

		const target = await openFixture('mixed')
		const imported = target.importShape(target.slides[0], source.slides[4], 4)
		assert(drawingIds(imported.element_).join() !== drawingIds(sourceGroup).join(), 'the imported group was renumbered')
		assertEqual(
			bindingTargets(imported.element_).join(),
			inside.join(),
			'each binding inside the group still names its shape'
		)
	})

	test('importShape drops a binding to a shape the import did not carry', async () => {
		// `mixed.pptx` slide 5, shape 3: a group of ids 8, 7 and 16 whose connector's `a:stCxn` names id 4,
		// a shape outside the group. Imported onto slide 1 the group is renumbered 4, 5 and 6, so the
		// source id, kept, bound the connector to its own enclosing group.
		const source = await openFixture('mixed')
		const sourceGroup = source.slides[4].shapes[3].element_
		const targets = bindingTargets(sourceGroup)
		assert(targets.includes(null), 'the fixture group binds a connector to a shape outside itself')

		const target = await openFixture('mixed')
		const imported = target.importShape(target.slides[0], source.slides[4], 3)
		assertEqual(
			bindingTargets(imported.element_).join(),
			targets.filter((name) => name !== null).join(),
			'only the bindings to shapes inside the group are carried'
		)
	})
})
