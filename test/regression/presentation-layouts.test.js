import {
	EMU_PER_INCH,
	EMU_PER_POINT,
	STANDARD_LAYOUTS,
	emuToInches,
	emuToPixels,
	emuToPoints,
	inchesToEmu,
	pixelsToEmu,
	pointsToEmu,
} from '../../dist/core.js'
import { defineRegressionSuite, build, readEntry, assert, assertEqual } from '../helpers.js'

const WIDE = STANDARD_LAYOUTS.LAYOUT_WIDE

async function assertPresentationSize(buildFn, expected, label) {
	const { zip } = await build(buildFn)
	const xml = await readEntry(zip, 'ppt/presentation.xml')
	const expectedTag = `<p:sldSz cx="${expected.widthEmu}" cy="${expected.heightEmu}"/>`
	assert(xml.includes(expectedTag), `${label}: expected presentation size ${expectedTag}, got ${xml}`)
}

defineRegressionSuite('Presentation layouts', 'legacy bug-22', [
	{
		name: 'public unit helpers expose PowerPoint EMU conversions',
		fn: () => {
			assertEqual(EMU_PER_INCH, 914400, 'EMU_PER_INCH')
			assertEqual(EMU_PER_POINT, 12700, 'EMU_PER_POINT')
			assertEqual(inchesToEmu(1), 914400, 'inchesToEmu')
			assertEqual(pointsToEmu(72), 914400, 'pointsToEmu')
			assertEqual(pixelsToEmu(1920, 144), 12192000, 'pixelsToEmu')
			assertEqual(emuToInches(914400), 1, 'emuToInches')
			assertEqual(emuToPoints(914400), 72, 'emuToPoints')
			assertEqual(emuToPixels(12192000, 144), 1920, 'emuToPixels')
		},
	},
	{
		name: 'standard wide layout constant matches PowerPoint widescreen',
		fn: () => {
			assertEqual(WIDE.layout, 'LAYOUT_WIDE', 'layout key')
			assertEqual(WIDE.name, 'custom', 'presentation layout name')
			assert(Math.abs(WIDE.widthIn - 13.333333333333334) < 0.00000000000001, 'expected 40/3 inch width')
			assertEqual(WIDE.heightIn, 7.5, 'heightIn')
			assertEqual(WIDE.widthEmu, 12192000, 'widthEmu')
			assertEqual(WIDE.heightEmu, 6858000, 'heightEmu')
			assertEqual(inchesToEmu(WIDE.widthIn), WIDE.widthEmu, 'widthIn converts to widthEmu')
			assertEqual(inchesToEmu(WIDE.heightIn), WIDE.heightEmu, 'heightIn converts to heightEmu')
		},
	},
	{
		name: 'built-in LAYOUT_WIDE writes exact PowerPoint widescreen EMUs',
		fn: async () => {
			await assertPresentationSize(
				(p) => {
					p.layout = 'LAYOUT_WIDE'
					p.addSlide()
				},
				WIDE,
				'LAYOUT_WIDE'
			)
		},
	},
	{
		name: 'custom layout from wide constants writes exact PowerPoint widescreen EMUs',
		fn: async () => {
			await assertPresentationSize(
				(p) => {
					p.defineLayout({ name: 'POWERPOINT_WIDESCREEN', width: WIDE.widthIn, height: WIDE.heightIn })
					p.layout = 'POWERPOINT_WIDESCREEN'
					p.addSlide()
				},
				WIDE,
				'custom wide layout'
			)
		},
	},
	{
		name: 'standard layout presets expose intuitive inch accessors (.width/.height)',
		fn: () => {
			const wide = STANDARD_LAYOUTS.LAYOUT_WIDE
			const std = STANDARD_LAYOUTS.LAYOUT_16x9
			assertEqual(wide.width, wide.widthIn, 'LAYOUT_WIDE .width aliases .widthIn')
			assertEqual(wide.height, wide.heightIn, 'LAYOUT_WIDE .height aliases .heightIn')
			assertEqual(std.width, 10, 'LAYOUT_16x9 .width is 10in')
			assertEqual(std.height, 5.625, 'LAYOUT_16x9 .height is 5.625in')
		},
	},
	{
		name: 'pptx.layout accepts a STANDARD_LAYOUTS preset object directly',
		fn: async () => {
			await assertPresentationSize(
				(p) => {
					p.layout = STANDARD_LAYOUTS.LAYOUT_WIDE
					p.addSlide()
				},
				WIDE,
				'preset object assignment'
			)
		},
	},
	{
		name: 'slide.width/slide.height return the active layout size in inches',
		fn: async () => {
			const { pres } = await build((p) => {
				p.layout = STANDARD_LAYOUTS.LAYOUT_16x9
				const slide = p.addSlide()
				assertEqual(slide.width, 10, 'slide.width inches')
				assertEqual(slide.height, 5.625, 'slide.height inches')
			})
			assert(pres, 'presentation built')
		},
	},
	{
		name: 'non-finite coordinates fail loud instead of emitting zero-size objects',
		fn: async () => {
			const layout = STANDARD_LAYOUTS.LAYOUT_16x9
			let threw = null
			try {
				// Reproduces the footgun: reading `.width`/`.height` off a value that lacks them
				// yields undefined -> NaN coordinate math.
				const bogus = undefined
				await build((p) => {
					p.layout = layout
					p.addSlide().addText('collapses', { x: 0.5, y: 0.5, w: bogus - 1, h: 1 })
				})
			} catch (err) {
				threw = err
			}
			assert(threw instanceof Error, 'expected a thrown Error for a NaN width')
			assert(/finite number/.test(threw.message), `expected a descriptive message, got: ${threw && threw.message}`)
		},
	},
])
