import { ShapeType } from '../../dist/node.js'
import {
	setDiagnosticHandler,
	defineRegressionSuite,
	build,
	readEntry,
	assert,
	assertEqual,
	assertXmlOrder,
} from '../helpers.js'

defineRegressionSuite('Shape preset mapping', 'legacy bug-10', [
	{
		name: 'addShape("oval", ...) emits prst="ellipse" (not invalid "oval")',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape('oval', { x: 1, y: 1, w: 0.4, h: 0.4, fill: { color: '00B0B9' } })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:prstGeom\s+prst="ellipse"/.test(xml), 'expected prstGeom prst="ellipse" in slide1.xml; got: ' + xml)
			assert(!/<a:prstGeom\s+prst="oval"/.test(xml), 'invalid prst="oval" still present in slide1.xml')
		},
	},
	{
		name: 'addShape("roundedRectangle", ...) emits prst="roundRect" (not invalid "roundedRectangle")',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape('roundedRectangle', { x: 1, y: 1, w: 2, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				/<a:prstGeom\s+prst="roundRect"/.test(xml),
				'expected prstGeom prst="roundRect" in slide1.xml; got: ' + xml
			)
			assert(
				!/<a:prstGeom\s+prst="roundedRectangle"/.test(xml),
				'invalid prst="roundedRectangle" still present in slide1.xml'
			)
		},
	},
	{
		name: 'addShape("rectangle", ...) emits prst="rect" (not invalid "rectangle")',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape('rectangle', { x: 1, y: 1, w: 2, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:prstGeom\s+prst="rect"/.test(xml), 'expected prstGeom prst="rect" in slide1.xml; got: ' + xml)
			assert(!/<a:prstGeom\s+prst="rectangle"/.test(xml), 'invalid prst="rectangle" still present in slide1.xml')
		},
	},
	{
		name: 'enum-constant API still works (ShapeType.ellipse -> ellipse)',
		fn: async () => {
			const { zip, pres } = await build((p) => {
				const s = p.addSlide()
				s.addShape(ShapeType.ellipse, { x: 1, y: 1, w: 0.4, h: 0.4 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				/<a:prstGeom\s+prst="ellipse"/.test(xml),
				'expected prstGeom prst="ellipse" via ShapeType.ellipse; got: ' + xml
			)
			void pres
		},
	},
	{
		name: 'addShape with an unknown preset throws instead of emitting invalid prst',
		fn: async () => {
			let threw = false
			try {
				await build((p) => {
					const s = p.addSlide()
					s.addShape('hexgon', { x: 1, y: 1, w: 2, h: 1 }) // typo for "hexagon"
				})
			} catch (err) {
				threw = true
				assert(/Invalid shape "hexgon"/.test(String(err.message)), 'unexpected error message: ' + err.message)
			}
			assert(threw, 'expected addShape("hexgon") to throw')
		},
	},
	{
		name: 'ShapeType.foldedCorner emits the valid spec spelling prst="foldedCorner"',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape(ShapeType.foldedCorner, { x: 1, y: 1, w: 2, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:prstGeom\s+prst="foldedCorner"/.test(xml), 'expected prst="foldedCorner"; got: ' + xml)
			assert(!/prst="folderCorner"/.test(xml), 'invalid prst="folderCorner" still present')
		},
	},
	{
		name: 'a valid ST_ShapeType preset not exposed via shapes.* (straightConnector1) is accepted',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape('straightConnector1', { x: 1, y: 1, w: 2, h: 0 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:prstGeom\s+prst="straightConnector1"/.test(xml), 'expected prst="straightConnector1"; got: ' + xml)
		},
	},
	{
		name: 'addText with a valid shape preset emits that prstGeom',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hi', { shape: 'ellipse', x: 1, y: 1, w: 1, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:prstGeom\s+prst="ellipse"/.test(xml), 'expected prst="ellipse" via addText shape; got: ' + xml)
		},
	},
	{
		name: 'addText with an invalid shape preset throws (gen-xml safety net)',
		fn: async () => {
			let threw = false
			try {
				await build((p) => {
					const s = p.addSlide()
					s.addText('hi', { shape: 'ellipsee', x: 1, y: 1, w: 1, h: 1 }) // typo
				})
			} catch (err) {
				threw = true
				assert(/Invalid shape "ellipsee"/.test(String(err.message)), 'unexpected error message: ' + err.message)
			}
			assert(threw, 'expected addText with invalid shape to throw')
		},
	},
	{
		name: 'custGeom freeform shape is accepted (special-cased, not a prstGeom)',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape(ShapeType.custGeom, {
					x: 1,
					y: 1,
					w: 2,
					h: 2,
					points: [
						{ x: 0, y: 0 },
						{ x: 2, y: 0 },
						{ x: 1, y: 2, close: true },
					],
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:custGeom>/.test(xml), 'expected <a:custGeom> for CUSTOM_GEOMETRY; got: ' + xml)
		},
	},
	{
		name: 'custGeom with no guides/connectionSites/adjustHandles emits the empty sub-lists (byte-identity contract)',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape(ShapeType.custGeom, {
					x: 1,
					y: 1,
					w: 2,
					h: 2,
					points: [
						{ x: 0, y: 0 },
						{ x: 2, y: 0 },
						{ x: 1, y: 2, close: true },
					],
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			// Locks the exact empty-case bytes the byte-identity gate depends on.
			assert(/<a:avLst \/>/.test(xml), 'expected empty <a:avLst />; got: ' + xml)
			assert(/<a:gdLst><\/a:gdLst>/.test(xml), 'expected empty <a:gdLst></a:gdLst>; got: ' + xml)
			assert(/<a:ahLst \/>/.test(xml), 'expected empty <a:ahLst />; got: ' + xml)
			assert(/<a:cxnLst><\/a:cxnLst>/.test(xml), 'expected empty <a:cxnLst></a:cxnLst>; got: ' + xml)
		},
	},
	{
		name: 'custGeom populates gdLst/cxnLst/ahLst (XY + polar) in the correct child order',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape(ShapeType.custGeom, {
					x: 1,
					y: 1,
					w: 2,
					h: 2,
					points: [
						{ x: 0, y: 0 },
						{ x: 2, y: 0 },
						{ x: 1, y: 2, close: true },
					],
					guides: [{ name: 'w2', formula: '*/ w 1 2' }],
					connectionSites: [
						{ ang: 0, x: 1, y: 0 }, // 1in -> 914400 EMU; ang 0deg -> 0
						{ ang: 90, x: 'hc', y: 0 }, // guide-name x emitted verbatim; ang 90deg -> 5400000
					],
					adjustHandles: [
						{ x: 0, y: 0, gdRefX: 'w2', minX: 0, maxX: 1 },
						{ x: 1, y: 1, gdRefAng: 'a1', minAng: 0, maxAng: 90 },
					],
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')

			assert(
				xml.includes('<a:gdLst><a:gd name="w2" fmla="*/ w 1 2"/></a:gdLst>'),
				'expected populated gdLst; got: ' + xml
			)
			assert(
				xml.includes('<a:cxnLst><a:cxn ang="0"><a:pos x="914400" y="0"/></a:cxn>'),
				'expected first cxn with EMU-resolved pos; got: ' + xml
			)
			assert(
				xml.includes('<a:cxn ang="5400000"><a:pos x="hc" y="0"/></a:cxn></a:cxnLst>'),
				'expected second cxn with guide-name x emitted verbatim; got: ' + xml
			)
			assert(
				xml.includes('<a:ahLst><a:ahXY gdRefX="w2" minX="0" maxX="914400"><a:pos x="0" y="0"/></a:ahXY>'),
				'expected ahXY handle; got: ' + xml
			)
			assert(
				xml.includes(
					'<a:ahPolar gdRefAng="a1" minAng="0" maxAng="5400000"><a:pos x="914400" y="914400"/></a:ahPolar></a:ahLst>'
				),
				'expected ahPolar handle; got: ' + xml
			)

			// CT_CustomGeometry2D child order: avLst -> gdLst -> ahLst -> cxnLst.
			assertXmlOrder(xml, '<a:avLst', '<a:gdLst', 'custGeom child order')
			assertXmlOrder(xml, '<a:gdLst', '<a:ahLst', 'custGeom child order')
			assertXmlOrder(xml, '<a:ahLst', '<a:cxnLst', 'custGeom child order')
			assertXmlOrder(xml, '<a:cxnLst', '<a:rect', 'custGeom child order')
			assertXmlOrder(xml, '<a:rect', '<a:pathLst', 'custGeom child order')
		},
	},
	{
		name: 'custGeom invalid guide (empty name/formula) is dropped and warns',
		fn: async () => {
			const warnings = []
			setDiagnosticHandler((d) => warnings.push(d.message))
			let xml
			try {
				const { zip } = await build((p) => {
					const s = p.addSlide()
					s.addShape(ShapeType.custGeom, {
						x: 1,
						y: 1,
						w: 2,
						h: 2,
						points: [{ x: 0, y: 0 }],
						guides: [{ name: '', formula: '' }],
					})
				})
				xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			} finally {
				setDiagnosticHandler(null)
			}
			assert(
				warnings.some((w) => w.includes('guide entry') && w.includes('was ignored')),
				'expected a warning for the invalid guide; got: ' + warnings.join(' | ')
			)
			// The degenerate guide is dropped, so the list falls back to the empty-case bytes.
			assert(/<a:gdLst><\/a:gdLst>/.test(xml), 'expected empty <a:gdLst></a:gdLst> after dropping guide; got: ' + xml)
			assert(!/<a:gd\b/.test(xml), 'no <a:gd> should be emitted for the invalid guide; got: ' + xml)
		},
	},
	{
		name: 'custGeom guide with an unknown formula operation is dropped and warns',
		fn: async () => {
			const warnings = []
			setDiagnosticHandler((d) => warnings.push(d.message))
			let xml
			try {
				const { zip } = await build((p) => {
					const s = p.addSlide()
					s.addShape(ShapeType.custGeom, {
						x: 1,
						y: 1,
						w: 2,
						h: 2,
						points: [{ x: 0, y: 0 }],
						guides: [
							{ name: 'w2', formula: '*/ w 1 2' }, // valid op, kept
							{ name: 'bad', formula: 'bogus 1 2' }, // unknown op, dropped
						],
					})
				})
				xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			} finally {
				setDiagnosticHandler(null)
			}
			assert(
				warnings.some((w) => w.includes('unknown operation "bogus"') && w.includes('was ignored')),
				'expected a warning naming the unknown operation; got: ' + warnings.join(' | ')
			)
			assert(
				xml.includes('<a:gdLst><a:gd name="w2" fmla="*/ w 1 2"/></a:gdLst>'),
				'the valid guide should survive alone; got: ' + xml
			)
			assert(!xml.includes('name="bad"'), 'no <a:gd> should be emitted for the unknown-op guide; got: ' + xml)
		},
	},
	{
		name: 'custGeom accepts every ECMA-376 guide formula operation',
		fn: async () => {
			// The closed set from ECMA-376 Part 1 §20.1.9.11; a regression here means the
			// validator started rejecting a legitimate formula, which is worse than the
			// footgun it guards against.
			const ops = [
				'*/',
				'+-',
				'+/',
				'?:',
				'abs',
				'at2',
				'cat2',
				'cos',
				'max',
				'min',
				'mod',
				'pin',
				'sat2',
				'sin',
				'sqrt',
				'tan',
				'val',
			]
			const warnings = []
			setDiagnosticHandler((d) => warnings.push(d.message))
			let xml
			try {
				const { zip } = await build((p) => {
					const s = p.addSlide()
					s.addShape(ShapeType.custGeom, {
						x: 1,
						y: 1,
						w: 2,
						h: 2,
						points: [{ x: 0, y: 0 }],
						guides: ops.map((op, i) => ({ name: `g${i}`, formula: `${op} w 1 2` })),
					})
				})
				xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			} finally {
				setDiagnosticHandler(null)
			}
			assertEqual(warnings.length, 0, 'no operation should warn; got: ' + warnings.join(' | '))
			ops.forEach((op, i) => {
				assert(xml.includes(`name="g${i}"`), `guide for op "${op}" should be emitted; got: ` + xml)
			})
		},
	},
	{
		name: 'custGeom invalid connectionSite (non-finite ang) is dropped and warns',
		fn: async () => {
			const warnings = []
			setDiagnosticHandler((d) => warnings.push(d.message))
			let xml
			try {
				const { zip } = await build((p) => {
					const s = p.addSlide()
					s.addShape(ShapeType.custGeom, {
						x: 1,
						y: 1,
						w: 2,
						h: 2,
						points: [{ x: 0, y: 0 }],
						connectionSites: [
							{ ang: NaN, x: 0, y: 0 }, // invalid: dropped + warned
							{ ang: 0, x: 1, y: 1 }, // valid: still emitted
						],
					})
				})
				xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			} finally {
				setDiagnosticHandler(null)
			}
			assert(
				warnings.some((w) => w.includes('connectionSite entry') && w.includes('was ignored')),
				'expected a warning for the invalid connectionSite; got: ' + warnings.join(' | ')
			)
			// Only the one valid site survives.
			assert(
				xml.includes('<a:cxnLst><a:cxn ang="0"><a:pos x="914400" y="914400"/></a:cxn></a:cxnLst>'),
				'expected only the valid connection site to be emitted; got: ' + xml
			)
		},
	},
	{
		// Connector presets are typed on the public `SHAPE_NAME` union (derived from
		// EXTRA_SHAPE_PRESETS), so a string-literal `addShape('bentConnector2', ...)` is a
		// valid typed call AND serializes as a static prstGeom connector geometry. This test
		// runs under tsconfig.test.json, so it also fails to compile if the union regresses.
		name: 'addShape with a connector preset is typed and emits its prstGeom',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape('bentConnector2', { x: 1, y: 1, w: 2, h: 1 })
				s.addShape('curvedConnector4', { x: 1, y: 3, w: 2, h: 1 })
				s.addShape('straightConnector1', { x: 1, y: 5, w: 2, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:prstGeom\s+prst="bentConnector2"/.test(xml), 'expected prst="bentConnector2"; got: ' + xml)
			assert(/<a:prstGeom\s+prst="curvedConnector4"/.test(xml), 'expected prst="curvedConnector4"')
			assert(/<a:prstGeom\s+prst="straightConnector1"/.test(xml), 'expected prst="straightConnector1"')
		},
	},
])
