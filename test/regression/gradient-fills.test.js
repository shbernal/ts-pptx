import {
	defineRegressionSuite,
	build,
	readEntry,
	assert,
	assertIncludes,
	assertXmlOrder,
	firstXmlBlock,
} from '../helpers.js'

async function expectBuildError(buildFn, expectedMessage) {
	let err
	try {
		await build(buildFn)
	} catch (e) {
		err = e
	}
	assert(err, 'expected build to fail')
	const message = String(err?.message || err)
	assert(message.includes(expectedMessage), `expected error to include "${expectedMessage}"; got: ${message}`)
}

defineRegressionSuite('Gradient fills', [
	{
		name: 'shape fill emits native linear gradient before line properties',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.RECTANGLE, {
					x: 1,
					y: 1,
					w: 3,
					h: 1,
					fill: {
						type: 'gradient',
						gradient: {
							kind: 'linear',
							angle: 90,
							scaled: true,
							stops: [
								{ position: 100, color: 'accent1', transparency: 25 },
								{ position: 0, color: '#451DC7' },
							],
						},
					},
					line: { color: '111111', width: 1 },
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const shapeBlock = firstXmlBlock(xml, 'p:sp', 'shape')

			assertIncludes(shapeBlock, '<a:gradFill rotWithShape="1">', 'shape gradient fill')
			assertIncludes(shapeBlock, '<a:gs pos="0"><a:srgbClr val="451DC7"/></a:gs>', 'first gradient stop')
			assertIncludes(
				shapeBlock,
				'<a:gs pos="100000"><a:schemeClr val="accent1"><a:alpha val="75000"/></a:schemeClr></a:gs>',
				'theme gradient stop with transparency'
			)
			assertIncludes(shapeBlock, '<a:lin ang="5400000" scaled="1"/>', 'linear gradient settings')
			assertXmlOrder(shapeBlock, '<a:prstGeom', '<a:gradFill', 'shape properties')
			assertXmlOrder(shapeBlock, '<a:gradFill', '<a:ln', 'shape properties')
		},
	},
	{
		name: 'slide background emits native linear gradient before effect list',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.background = {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 90,
						scaled: true,
						rotateWithShape: true,
						stops: [
							{ position: 0, color: '451DC7' },
							{ position: 100, color: '0B003D', transparency: 10 },
						],
					},
				}
				s.addText('hi', { x: 1, y: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const backgroundBlock = firstXmlBlock(xml, 'p:bg', 'slide background')

			assertIncludes(backgroundBlock, '<a:gradFill rotWithShape="1">', 'background gradient fill')
			assertIncludes(
				backgroundBlock,
				'<a:gs pos="100000"><a:srgbClr val="0B003D"><a:alpha val="90000"/></a:srgbClr></a:gs>',
				'transparent gradient stop'
			)
			assertIncludes(backgroundBlock, '<a:lin ang="5400000" scaled="1"/>', 'background linear gradient settings')
			assertXmlOrder(backgroundBlock, '<a:gradFill', '<a:effectLst/>', 'background properties')
		},
	},
	{
		name: 'linear gradient angles normalize to legal positive fixed angles',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				;[0, 359, 360, -90].forEach((angle, idx) => {
					s.addShape(p.shapes.RECTANGLE, {
						x: 1,
						y: 1 + idx,
						w: 2,
						h: 0.5,
						fill: {
							type: 'gradient',
							gradient: {
								kind: 'linear',
								angle,
								stops: [
									{ position: 0, color: '000000' },
									{ position: 100, color: 'FFFFFF' },
								],
							},
						},
					})
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const angles = [...xml.matchAll(/<a:lin ang="(\d+)"\/>/g)].map((match) => match[1])
			const expected = ['0', '21540000', '0', '16200000']
			assert(
				JSON.stringify(angles) === JSON.stringify(expected),
				`expected normalized angles ${JSON.stringify(expected)}; got: ${JSON.stringify(angles)}`
			)
		},
	},
	{
		name: 'gradient fill rejects fewer than two stops',
		fn: async () => {
			await expectBuildError((p) => {
				const s = p.addSlide()
				s.background = {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						stops: [{ position: 0, color: '000000' }],
					},
				}
			}, 'at least two stops')
		},
	},
	{
		name: 'gradient fill rejects out-of-range stop positions',
		fn: async () => {
			await expectBuildError((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.RECTANGLE, {
					x: 1,
					y: 1,
					w: 2,
					h: 1,
					fill: {
						type: 'gradient',
						gradient: {
							kind: 'linear',
							stops: [
								{ position: -1, color: '000000' },
								{ position: 100, color: 'FFFFFF' },
							],
						},
					},
				})
			}, 'from 0 to 100')
		},
	},
	{
		name: 'gradient fill rejects non-finite angles',
		fn: async () => {
			await expectBuildError((p) => {
				const s = p.addSlide()
				s.background = {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: Number.POSITIVE_INFINITY,
						stops: [
							{ position: 0, color: '000000' },
							{ position: 100, color: 'FFFFFF' },
						],
					},
				}
			}, 'finite number')
		},
	},
])
