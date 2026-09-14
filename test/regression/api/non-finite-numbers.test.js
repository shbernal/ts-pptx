// A number option that is not a number, or is out of range, on the write side. The rule the library
// follows elsewhere (`clampRangedInput`, docs/contributing/development.md "Warn or throw?") is that a finite
// value past a bound has a nearest legal neighbour, so it moves there with a warning, and `NaN` has
// none, so it throws naming the option. The paths below each did something else: threw without
// naming the option, dropped the value without a word, or wrote `NaN` into the part.

import { describe, test } from 'vitest'
import { TsPptx, assert, assertEqual, build, captureDiagnostics, readEntry } from '../../helpers.js'

const BOX = { x: 1, y: 1, w: 4, h: 1 }

/**
 * Build a deck and read slide 1 back with the diagnostics it raised, or the error the build threw.
 * @param {(pres: any) => void} author
 * @returns {Promise<{ xml?: string, rels?: string, codes?: string[], error?: any }>}
 */
async function outcome(author) {
	try {
		const { result, codes } = await captureDiagnostics(async () => {
			const { zip } = await build(author)
			return {
				xml: await readEntry(zip, 'ppt/slides/slide1.xml'),
				rels: await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels'),
			}
		})
		return { ...result, codes }
	} catch (error) {
		return { error }
	}
}

/**
 * The code of the error `fn` throws, or `undefined` when it returns.
 * @param {() => unknown} fn
 */
function thrownCode(fn) {
	try {
		fn()
	} catch (err) {
		return /** @type {any} */ (err).code
	}
	return undefined
}

describe('non-finite and out-of-range numbers on the write side', () => {
	test('every clamped text measure clamps Infinity with a warning and refuses NaN by name', async () => {
		/** @type {Array<[string, string, string, string]>} option, attribute at the bound, clamp code, NaN code */
		const cases = [
			['fontSize', ' sz="400000"', 'font/size-out-of-range', 'coord/non-finite'],
			['charSpacing', ' spc="400000"', 'text/char-spacing-out-of-range', 'coord/non-finite'],
			['lineSpacing', '<a:spcPts val="158400"', 'text/line-spacing-out-of-range', 'coord/non-finite'],
			['lineSpacingMultiple', '<a:spcPct val="13200000"', 'text/line-spacing-out-of-range', 'percent/non-finite'],
			['transparency', '<a:alpha val="0"', 'transparency/out-of-range', 'percent/non-finite'],
		]
		for (const [option, atBound, clampCode, nanCode] of cases) {
			const infinite = await outcome((p) => p.addSlide().addText('x', { ...BOX, [option]: Infinity }))
			assert(!infinite.error, `${option} Infinity is clamped, not refused: ${infinite.error?.message}`)
			assert(infinite.xml?.includes(atBound), `${option} Infinity is written at the bound ${atBound}`)
			assert(infinite.codes?.includes(clampCode), `${option} Infinity warns ${clampCode}: ${infinite.codes}`)

			const nan = await outcome((p) => p.addSlide().addText('x', { ...BOX, [option]: NaN }))
			assertEqual(nan.error?.code, nanCode, `${option} NaN is refused`)
			assert(String(nan.error?.message).includes(option), `the refusal names ${option}: ${nan.error?.message}`)
		}
	})

	test('an OLE image size that is not an EMU size in range is refused, and a fraction is rounded', async () => {
		const ole = (imgW, imgH) => (p) =>
			p.addSlide().addOleObject({ data: 'UEsDBA==', progId: 'Excel.Sheet.12', imgW, imgH, ...BOX })
		assertEqual((await outcome(ole(NaN, 100))).error?.code, 'ole/invalid-image-size', 'a NaN width')
		assertEqual((await outcome(ole(100, -5))).error?.code, 'ole/invalid-image-size', 'a negative height')
		const rounded = await outcome(ole(1.5, 2))
		assert(rounded.xml?.includes('imgW="2" imgH="2"'), `a fraction is rounded: ${rounded.xml}`)
	})

	test('a zoom transitionDur is refused when NaN, clamped when out of range, and rounded', async () => {
		/** @param {number} transitionDur */
		const zoom = (transitionDur) => (p) => {
			const host = p.addSlide()
			p.addSlide()
			host.addSlideZoom({ target: 2, ...BOX, transitionDur })
		}
		assertEqual((await outcome(zoom(NaN))).error?.code, 'zoom/invalid-transition-duration', 'NaN')
		for (const [given, written, warned] of [
			[-1, '0', true],
			[Infinity, '2147483647', true],
			[1.5, '2', false],
			[1000, '1000', false],
		]) {
			const result = await outcome(zoom(/** @type {number} */ (given)))
			assert(result.xml?.includes(`transitionDur="${written}"`), `${given} is written as ${written}`)
			assertEqual(result.codes?.includes('zoom/transition-duration-out-of-range'), warned, `${given} warns: ${warned}`)
		}

		// The refusal comes before the zoom registers its preview image or its slide link.
		const refused = await outcome((p) => {
			const host = p.addSlide()
			p.addSlide()
			try {
				host.addSlideZoom({ target: 2, ...BOX, transitionDur: NaN })
			} catch {
				// refused, as asserted above
			}
		})
		assert(
			!refused.rels?.includes('../media/') && !refused.rels?.includes('slide2.xml'),
			`nothing registered: ${refused.rels}`
		)
	})

	test('a media loopCount that plays nothing warns and plays once, and a real count is kept', async () => {
		const media = (loopCount) => (p) =>
			p.addSlide().addMedia({ type: 'audio', data: 'data:audio/mpeg;base64,SUQz', ...BOX, loopCount })
		for (const loopCount of [NaN, 0, -1, Infinity]) {
			const result = await outcome(media(loopCount))
			assert(result.codes?.includes('media/invalid-loop-count'), `${loopCount} warns: ${result.codes}`)
			assert(!result.xml?.includes('repeatCount'), `${loopCount} writes no repeat count`)
		}
		const twice = await outcome(media(2))
		assert(twice.xml?.includes('repeatCount="2000"'), 'a count of 2 plays twice')
		assert(!twice.codes?.includes('media/invalid-loop-count'), 'and says nothing')
	})

	test('measureText refuses a box it cannot lay out, and the height checks refuse one too', () => {
		const pres = new TsPptx()
		const words = 'one two three four five six seven eight nine ten eleven twelve'
		const measure = (opts) => pres.measureText(words, /** @type {any} */ ({ fontFace: 'Arial', fontSize: 12, ...opts }))
		assertEqual(
			thrownCode(() => measure({ wIn: NaN })),
			'coord/non-finite',
			'a NaN width'
		)
		assertEqual(
			thrownCode(() => measure({})),
			'coord/non-finite',
			'no width'
		)
		assertEqual(
			thrownCode(() => measure({ wIn: Infinity })),
			'coord/non-finite',
			'an infinite width'
		)
		assertEqual(
			thrownCode(() => measure({ wIn: 2, insetIn: NaN })),
			'coord/non-finite',
			'a NaN inset'
		)
		assertEqual(
			thrownCode(() => measure({ wIn: -1 })),
			'coord/not-positive',
			'a negative width'
		)
		assertEqual(
			thrownCode(() => measure({ wIn: 2, insetIn: 1 })),
			'coord/not-positive',
			'insets that leave no width'
		)

		const measured = measure({ wIn: 3 })
		assert(measured.measurable, 'a box with width still measures')
		assertEqual(
			thrownCode(() => measured.fitsBox(NaN)),
			'coord/non-finite',
			'fitsBox NaN'
		)
		assertEqual(
			thrownCode(() => measured.shrinkScaleFor(0)),
			'coord/not-positive',
			'shrinkScaleFor 0'
		)
		assertEqual(
			thrownCode(() => pres.overflowsBox(words, { fontFace: 'Arial', fontSize: 12, wIn: 3, hIn: NaN })),
			'coord/non-finite',
			'overflowsBox NaN'
		)
		assertEqual(typeof measured.fitsBox(1), 'boolean', 'a real height is answered')
	})
})
