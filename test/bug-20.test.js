import { build, readEntry, assert } from './helpers.js'

// Returns the byte index of `needle` in `haystack`, or -1 if not present.
// We use indexOf on the raw XML string to verify schema-canonical ordering of
// CT_Presentation child elements in `ppt/presentation.xml`.
function idxOf(xml, needle) {
	return xml.indexOf(needle)
}

export default [
	{
		name: 'ppt/presentation.xml: notesMasterIdLst appears AFTER sldMasterIdLst',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/presentation.xml')
			const iSldMaster = idxOf(xml, '<p:sldMasterIdLst>')
			const iNotesMaster = idxOf(xml, '<p:notesMasterIdLst>')
			assert(iSldMaster !== -1, 'expected <p:sldMasterIdLst> in presentation.xml; got: ' + xml)
			assert(iNotesMaster !== -1, 'expected <p:notesMasterIdLst> in presentation.xml; got: ' + xml)
			assert(
				iSldMaster < iNotesMaster,
				'expected <p:sldMasterIdLst> before <p:notesMasterIdLst>; got order sldMaster=' +
					iSldMaster +
					' notesMaster=' +
					iNotesMaster +
					' in: ' +
					xml
			)
		},
	},
	{
		name: 'ppt/presentation.xml: notesMasterIdLst appears BEFORE sldIdLst (CT_Presentation order)',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/presentation.xml')
			const iNotesMaster = idxOf(xml, '<p:notesMasterIdLst>')
			const iSldIdLst = idxOf(xml, '<p:sldIdLst>')
			assert(iNotesMaster !== -1, 'expected <p:notesMasterIdLst> in presentation.xml; got: ' + xml)
			assert(iSldIdLst !== -1, 'expected <p:sldIdLst> in presentation.xml; got: ' + xml)
			assert(
				iNotesMaster < iSldIdLst,
				'expected <p:notesMasterIdLst> before <p:sldIdLst> (CT_Presentation order); got order notesMaster=' +
					iNotesMaster +
					' sldIdLst=' +
					iSldIdLst +
					' in: ' +
					xml
			)
		},
	},
	{
		name: 'ppt/presentation.xml: sldIdLst appears BEFORE sldSz',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/presentation.xml')
			const iSldIdLst = idxOf(xml, '<p:sldIdLst>')
			const iSldSz = idxOf(xml, '<p:sldSz ')
			assert(iSldIdLst !== -1, 'expected <p:sldIdLst> in presentation.xml; got: ' + xml)
			assert(iSldSz !== -1, 'expected <p:sldSz ...> in presentation.xml; got: ' + xml)
			assert(
				iSldIdLst < iSldSz,
				'expected <p:sldIdLst> before <p:sldSz>; got order sldIdLst=' + iSldIdLst + ' sldSz=' + iSldSz + ' in: ' + xml
			)
		},
	},
	{
		name: 'ppt/presentation.xml: sldSz appears BEFORE notesSz',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/presentation.xml')
			const iSldSz = idxOf(xml, '<p:sldSz ')
			const iNotesSz = idxOf(xml, '<p:notesSz ')
			assert(iSldSz !== -1, 'expected <p:sldSz ...> in presentation.xml; got: ' + xml)
			assert(iNotesSz !== -1, 'expected <p:notesSz ...> in presentation.xml; got: ' + xml)
			assert(
				iSldSz < iNotesSz,
				'expected <p:sldSz> before <p:notesSz>; got order sldSz=' + iSldSz + ' notesSz=' + iNotesSz + ' in: ' + xml
			)
		},
	},
]
