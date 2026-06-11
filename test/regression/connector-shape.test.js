import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// Regression: slide.addConnector emits a PowerPoint connector (<p:cxnSp>) — not a plain line
// shape — with the correct connector preset, a min-corner origin + flip flags derived from the
// endpoints, and line styling/arrowheads (upstream gitbrent/PptxGenJS#1059).

async function slideXml(buildFn) {
	const { zip } = await build(buildFn)
	return readEntry(zip, 'ppt/slides/slide1.xml')
}

defineRegressionSuite('Connector shapes (upstream #1059)', [
	{
		name: 'straight connector emits cxnSp with straightConnector1 + styled line/arrow',
		fn: async () => {
			const xml = await slideXml((p) => {
				p.addSlide().addConnector({
					type: 'straight',
					x1: 1,
					y1: 1,
					x2: 4,
					y2: 3,
					color: 'FF0000',
					width: 2,
					endArrowType: 'triangle',
				})
			})
			const cxn = (xml.match(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g) || [])[0]
			assert(cxn, 'expected a <p:cxnSp> element')
			assert(cxn.includes('<a:prstGeom prst="straightConnector1">'), 'expected straightConnector1 preset')
			// origin = min corner (1in,1in = 914400 EMU); ext = |4-1|x|3-1| = 3in x 2in.
			assert(cxn.includes('<a:off x="914400" y="914400"/>'), 'expected origin at min corner')
			assert(cxn.includes('<a:ext cx="2743200" cy="1828800"/>'), 'expected ext from endpoint deltas')
			assert(cxn.includes('<a:srgbClr val="FF0000"/>'), 'expected red line color')
			assert(cxn.includes('w="25400"'), 'expected 2pt line width (25400 EMU)')
			assert(cxn.includes('<a:tailEnd type="triangle"/>'), 'expected end arrowhead')
			// A connector is NOT a <p:sp> and has no text body.
			assert(!cxn.includes('<p:txBody'), 'connector must not emit a text body')
		},
	},
	{
		name: 'reversed endpoints set flipH/flipV; elbow + curved map to bent/curved presets',
		fn: async () => {
			const xml = await slideXml((p) => {
				const s = p.addSlide()
				s.addConnector({ type: 'elbow', x1: 5, y1: 3, x2: 2, y2: 1, dashType: 'dash' }) // end is left/above start
				s.addConnector({ type: 'curved', x1: 1, y1: 1, x2: 3, y2: 4 })
			})
			const cxns = xml.match(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g) || []
			assert(cxns.length === 2, `expected 2 connectors; got ${cxns.length}`)
			assert(cxns[0].includes('<a:prstGeom prst="bentConnector3">'), 'elbow → bentConnector3')
			assert(/<a:xfrm flipH="1" flipV="1">/.test(cxns[0]), 'reversed endpoints must set flipH and flipV')
			assert(cxns[0].includes('<a:prstDash val="dash"/>'), 'expected dashed line')
			assert(cxns[1].includes('<a:prstGeom prst="curvedConnector3">'), 'curved → curvedConnector3')
			assert(!/flip[HV]="1"/.test(cxns[1]), 'forward endpoints must not flip')
		},
	},
	{
		name: 'missing endpoints throw',
		fn: async () => {
			let threw = false
			try {
				await build((p) => p.addSlide().addConnector({ x1: 1, y1: 1, x2: 4 }))
			} catch (ex) {
				threw = true
				assert(/x1, y1, x2, y2/.test(ex.message), `expected an endpoint-required error; got: ${ex.message}`)
			}
			assert(threw, 'addConnector without all endpoints must throw')
		},
	},
])
