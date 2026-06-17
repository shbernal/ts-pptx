import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

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
		name: 'enum-constant API still works (pres.shapes.OVAL -> ellipse)',
		fn: async () => {
			const { zip, pres } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.OVAL, { x: 1, y: 1, w: 0.4, h: 0.4 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				/<a:prstGeom\s+prst="ellipse"/.test(xml),
				'expected prstGeom prst="ellipse" via pres.shapes.OVAL; got: ' + xml
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
		name: 'pres.shapes.FOLDED_CORNER emits the valid spec spelling prst="foldedCorner"',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.FOLDED_CORNER, { x: 1, y: 1, w: 2, h: 1 })
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
				s.addShape(p.shapes.CUSTOM_GEOMETRY, {
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
