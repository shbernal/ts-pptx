// Read-model coverage for the `ST_Percentage` UNION, which no corpus fixture exercises.
//
// The Transitional profile admits two lexical forms for a percentage: the fixed-point integer
// PowerPoint writes (`100%` → `100000`) and a decimal string carrying a literal `%` (`"62.5%"`).
// The STRICT profile has only the second. `read/oxml/dom.ts` has carried the union parser and a
// note saying exactly this for a while; four getters read the attribute as a bare number instead,
// so on a Strict-profile or non-PowerPoint deck they reported a value that is present as `null` —
// and the model's contract for those nulls says "absent".
//
// The input is a real deck the write API authored, with the attributes rewritten into the string
// form: the schema is the oracle for a lexical form, the corpus holds none of it, and a synthetic
// `p:sld` with no package behind it cannot resolve a theme.

import { describe, test } from 'vitest'
import JSZip from 'jszip'
import TsPptx from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const SLIDE_PATH = 'ppt/slides/slide1.xml'

/**
 * Author a one-shape deck carrying every percentage under test, optionally rewrite its slide
 * XML, and return the shape's text frame from the loaded result.
 * @param {(xml: string) => string} [edit] - rewrite applied to `ppt/slides/slide1.xml`
 */
async function frameFrom(edit) {
	const pres = new TsPptx()
	pres.addSlide().addText('probe', {
		x: 1,
		y: 1,
		w: 4,
		h: 1,
		bullet: { type: 'bullet', size: 75 },
		lineSpacingMultiple: 1.5,
		fit: { type: 'shrink', fontScale: 62.5, lnSpcReduction: 20 },
	})
	const zip = await JSZip.loadAsync(await pres.toBytes())
	if (edit) {
		const before = await zip.file(SLIDE_PATH).async('string')
		const after = edit(before)
		assert(after !== before, 'the rewrite must actually change the slide XML')
		zip.file(SLIDE_PATH, after)
	}
	const presentation = await Presentation.load(await zip.generateAsync({ type: 'nodebuffer' }))
	return presentation.slides[0].shapes[0].textFrame
}

/** Add a `baseline` attribute to the run properties, in whichever lexical form. */
const withBaseline = (value) => (xml) => xml.replace('<a:rPr lang="en-US"', `<a:rPr baseline="${value}" lang="en-US"`)

describe('ST_Percentage union — the string form reads as the value it states', () => {
	test('a:rPr/@baseline: both lexical forms give the same percentage', async () => {
		assertEqual((await frameFrom()).paragraphs[0].runs[0].baselinePct, null, 'no @baseline is still null')
		for (const [value, expected] of [
			['30000', 30],
			['30%', 30],
			['62.5%', 62.5],
			['-40%', -40],
		]) {
			const frame = await frameFrom(withBaseline(value))
			assertEqual(frame.paragraphs[0].runs[0].baselinePct, expected, `baseline ${value}`)
		}
	})

	test('a:spcPct/@val: both lexical forms give the same line spacing', async () => {
		/** The paragraph's line spacing, asserted to be the percent variant. */
		const percentSpacing = (frame, label) => {
			const spacing = frame.paragraphs[0].lineSpacing
			assert(spacing?.type === 'percent', `${label}: expected a percent spacing; got ${spacing?.type}`)
			return spacing.percent
		}
		assertEqual(percentSpacing(await frameFrom(), 'fixed-point'), 150, 'fixed-point 150000')
		for (const [value, expected] of [
			['150%', 150],
			['112.5%', 112.5],
		]) {
			const frame = await frameFrom((xml) => xml.replace('<a:spcPct val="150000"', `<a:spcPct val="${value}"`))
			assertEqual(percentSpacing(frame, value), expected, `spcPct ${value}`)
		}
	})

	test('a:normAutofit percentages: both lexical forms give the same scale', async () => {
		const fixed = await frameFrom()
		assertEqual(fixed.autofitFontScale, 62.5, 'fixed-point fontScale')
		assertEqual(fixed.autofitLineSpaceReduction, 20, 'fixed-point lnSpcReduction')
		const string = await frameFrom((xml) =>
			xml.replace(
				'<a:normAutofit fontScale="62500" lnSpcReduction="20000"/>',
				'<a:normAutofit fontScale="62.5%" lnSpcReduction="20%"/>'
			)
		)
		assertEqual(string.autofitFontScale, 62.5, 'string fontScale')
		assertEqual(string.autofitLineSpaceReduction, 20, 'string lnSpcReduction')
	})

	test('a value in neither form is still null', async () => {
		const frame = await frameFrom(withBaseline('lots'))
		assertEqual(frame.paragraphs[0].runs[0].baselinePct, null, 'unparseable @baseline')
	})

	test('a:buSzPct is NOT a union and stays a bare fixed-point read', async () => {
		// `ST_TextBulletSizePercent` has no string form, so this one was correct as written and
		// must not be converted along with its neighbours.
		const detail = (await frameFrom()).paragraphs[0].bulletDetail
		assert(detail?.kind === 'char', `expected a glyph bullet; got ${detail?.kind}`)
		assertEqual(detail.sizePct, 75, 'buSzPct in thousandths of a percent')
	})
})
