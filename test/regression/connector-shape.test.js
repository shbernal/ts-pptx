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
		name: 'elbow adj sets the bentConnector3 jog guide; values map percent→1000ths',
		fn: async () => {
			const xml = await slideXml((p) => {
				p.addSlide().addConnector({ type: 'elbow', x1: 1, y1: 1, x2: 5, y2: 3, adj: 25 })
			})
			const cxn = (xml.match(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g) || [])[0]
			assert(cxn.includes('<a:prstGeom prst="bentConnector3">'), 'one bend → bentConnector3')
			// 25% → val 25000 on adj1 (the single bent-connector jog guide).
			assert(cxn.includes('<a:avLst><a:gd name="adj1" fmla="val 25000"/></a:avLst>'), 'expected adj1=25000')
		},
	},
	{
		name: 'bends selects bentConnector4/5 and adj array fills each jog (adj1..adjN)',
		fn: async () => {
			const xml = await slideXml((p) => {
				const s = p.addSlide()
				s.addConnector({ type: 'elbow', x1: 1, y1: 1, x2: 5, y2: 3, bends: 2, adj: [30, 70] })
				s.addConnector({ type: 'curved', x1: 1, y1: 4, x2: 5, y2: 6, bends: 3, adj: [10, 50, 90] })
			})
			const cxns = xml.match(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g) || []
			assert(cxns[0].includes('<a:prstGeom prst="bentConnector4">'), 'bends:2 → bentConnector4')
			assert(
				cxns[0].includes('<a:avLst><a:gd name="adj1" fmla="val 30000"/><a:gd name="adj2" fmla="val 70000"/></a:avLst>'),
				'expected two bent jogs'
			)
			assert(cxns[1].includes('<a:prstGeom prst="curvedConnector5">'), 'curved bends:3 → curvedConnector5')
			assert(
				cxns[1].includes(
					'<a:avLst><a:gd name="adj1" fmla="val 10000"/><a:gd name="adj2" fmla="val 50000"/><a:gd name="adj3" fmla="val 90000"/></a:avLst>'
				),
				'expected three curved jogs'
			)
		},
	},
	{
		name: 'adj array length must match bends; non-finite adj throws',
		fn: async () => {
			let threw = false
			try {
				await build((p) =>
					p.addSlide().addConnector({ type: 'elbow', x1: 1, y1: 1, x2: 5, y2: 3, bends: 2, adj: [50] })
				)
			} catch (ex) {
				threw = true
				assert(/must supply 2 value/.test(ex.message), `expected a length-mismatch error; got: ${ex.message}`)
			}
			assert(threw, 'mismatched adj length must throw')

			let threw2 = false
			try {
				await build((p) => p.addSlide().addConnector({ type: 'elbow', x1: 1, y1: 1, x2: 5, y2: 3, adj: Number.NaN }))
			} catch (ex) {
				threw2 = true
				assert(/finite number/.test(ex.message), `expected a finite-number error; got: ${ex.message}`)
			}
			assert(threw2, 'NaN adj must throw')
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
