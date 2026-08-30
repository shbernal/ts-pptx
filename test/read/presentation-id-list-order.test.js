// Where a get-or-added `p:sldMasterIdLst` lands in `CT_Presentation`.
//
// `getOrAddChild(root, qname, successors)` inserts before the first sibling named in
// `successors`, so a name missing from that list is a sibling the new child gets placed
// *after*. The three registration paths in `read/api` each carried their own hand-written copy
// of the `CT_Presentation` suffix and the copies disagreed: `p:smartTags` was in one and
// missing from the other two. They are now sliced out of the one declared sequence in
// `src/ooxml/sequence.ts`.
//
// Reaching the hole takes a degenerate presentation.xml. `p:notesSz` is mandatory in
// `CT_Presentation` and sits immediately before `p:smartTags`, so wherever the deck is
// schema-valid the successor search stops at `p:notesSz` and the missing name never decides
// anything. The registration paths handle a degenerate part on purpose — that is why they pass
// a successor list instead of appending — so the fixture below takes `p:sldIdLst`, `p:sldSz`
// and `p:notesSz` out along with the id list, which is the one construction where the old list
// and the new one place the element differently.
//
// The notes-master list had the same omission and gets no case here: every public route to it
// (`importSlide({ importNotes: true })`, `appendSlides`, `Slide.addNotes`) needs the
// destination to have either a slide or a slide size, and `p:sldIdLst` and `p:sldSz` both
// precede `p:smartTags` — so no deck that can reach `registerNotesMaster` can reach the hole.

import JSZip from 'jszip'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert } from '../helpers.js'
import { readFixture } from './corpus.js'

/** Element children of `presentation.xml`'s root, in document order, as qnames. */
function childOrder(presentation) {
	const root = presentation.opc.part('/ppt/presentation.xml').dom.documentElement
	const out = []
	for (let n = root.firstChild; n; n = n.nextSibling) {
		if (n.nodeType === 1) out.push(n.nodeName)
	}
	return out
}

/**
 * `empty.pptx` with every child between `p:sldMasterIdLst` and `p:smartTags` removed and a
 * `p:smartTags` put in their place — the state in which the successor list is what decides
 * where a re-added `p:sldMasterIdLst` goes. Its `r:id` points at a tags part added alongside
 * it, so the relationship graph stays whole.
 */
async function deckWithSmartTagsAndNoMasterIdLst() {
	const zip = await JSZip.loadAsync(await readFixture('empty'))

	const presName = 'ppt/presentation.xml'
	const xml = (await zip.file(presName).async('string'))
		.replace('<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>', '')
		.replace('<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>', '')
		.replace('<p:sldSz cx="12192000" cy="6858000"/>', '')
		.replace('<p:notesSz cx="6858000" cy="9144000"/>', '<p:smartTags r:id="rId7"/>')
	assert(xml.includes('<p:smartTags '), 'the patched fixture carries p:smartTags')
	assert(!xml.includes('<p:sldMasterIdLst>'), 'the patched fixture carries no p:sldMasterIdLst')
	assert(!xml.includes('<p:notesSz'), 'the patched fixture carries nothing between the two')
	zip.file(presName, xml)

	const relsName = 'ppt/_rels/presentation.xml.rels'
	zip.file(
		relsName,
		(await zip.file(relsName).async('string')).replace(
			'</Relationships>',
			'<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="tags/tag1.xml"/></Relationships>'
		)
	)
	zip.file(
		'ppt/tags/tag1.xml',
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
	)
	zip.file(
		'[Content_Types].xml',
		(await zip.file('[Content_Types].xml').async('string')).replace(
			'</Types>',
			'<Override PartName="/ppt/tags/tag1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"/></Types>'
		)
	)

	return Presentation.load(await zip.generateAsync({ type: 'uint8array' }))
}

describe('CT_Presentation id-list insertion point', () => {
	test('a registered slide master lands before p:smartTags, not after it', async () => {
		const dest = await deckWithSmartTagsAndNoMasterIdLst()
		const source = await Presentation.load(await readFixture('image'))
		dest.importSlideMasters(source, { requireEqualSize: false })

		const order = childOrder(dest)
		assert(order.includes('p:sldMasterIdLst'), `p:sldMasterIdLst was registered (${order.join(', ')})`)
		assert(
			order.indexOf('p:sldMasterIdLst') < order.indexOf('p:smartTags'),
			`p:sldMasterIdLst precedes p:smartTags (${order.join(', ')})`
		)
	})
})
