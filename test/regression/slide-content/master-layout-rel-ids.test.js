import { describe, expect, test } from 'vitest'
import { makeXmlMaster, makeXmlMasterRel } from '../../../src/gen/slide/master.ts'
import { composeFamilies } from '../../../src/families/shared.ts'
import { ALL_CONSTRUCT_FAMILIES } from '../../../src/entry-families.ts'

// The master's `<p:sldLayoutId r:id>` values name relationships its `.rels` part declares, and the
// two parts are written by different functions. `makeXmlMaster` derived the ids as
// `_rels.length + idx + 1` while the rels writer numbers them from the highest id across the
// master's hyperlink, chart and media relationships, so they agreed only while the master held
// no chart or media relationship. Nothing registers one today, which is why the corpus cannot
// tell: these inject one.

const RENDERERS = composeFamilies(ALL_CONSTRUCT_FAMILIES).renderers
const LAYOUT = { name: 'test', width: 9144000, height: 6858000 }

/** A part as `presentation.ts` constructs one, with `extra` on top. */
const part = (extra) => ({
	_presLayout: LAYOUT,
	_rels: [],
	_relsChart: [],
	_relsMedia: [],
	_slideNumberProps: null,
	_slideObjects: [],
	...extra,
})

const layouts = [
	part({ _name: 'First', _margin: [0.5, 0.5, 0.5, 0.5], _slide: null, _slideNum: 1000 }),
	part({ _name: 'Second', _margin: [0.5, 0.5, 0.5, 0.5], _slide: null, _slideNum: 1001 }),
]

/** One attribute of an element's open tag, in any attribute order. */
const attrOf = (tag, name) => new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1]

/** The ids the rels part gives the layouts, in document order. */
const layoutIdsInRels = (xml) =>
	[...xml.matchAll(/<Relationship\s[^>]*>/g)]
		.map(([tag]) => tag)
		.filter((tag) => attrOf(tag, 'Type')?.endsWith('/slideLayout'))
		.map((tag) => attrOf(tag, 'Id'))

/** The ids the master part names on its layout list, in document order. */
const layoutIdsInMaster = (xml) => [...xml.matchAll(/<p:sldLayoutId\s[^>]*>/g)].map(([tag]) => attrOf(tag, 'r:id'))

describe('the master names its layouts by the ids its relationships part gives them', () => {
	test('with no relationship of its own', () => {
		const master = part({ _slideNum: null })
		const rels = layoutIdsInRels(makeXmlMasterRel(master, layouts))
		expect(rels).toEqual(['rId1', 'rId2'])
		expect(layoutIdsInMaster(makeXmlMaster(master, layouts, RENDERERS))).toEqual(rels)
	})

	test('when the master carries a media relationship', () => {
		const master = part({
			_slideNum: null,
			_relsMedia: [{ rId: 3, type: 'image/png', Target: '../media/image1.png' }],
		})
		const rels = layoutIdsInRels(makeXmlMasterRel(master, layouts))
		expect(rels).toEqual(['rId4', 'rId5'])
		expect(layoutIdsInMaster(makeXmlMaster(master, layouts, RENDERERS))).toEqual(rels)
	})
})
