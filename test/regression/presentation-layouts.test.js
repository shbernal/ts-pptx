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
])
