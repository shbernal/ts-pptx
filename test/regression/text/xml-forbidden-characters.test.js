// Characters XML 1.0 forbids never reach a written part.
//
// The escapers strip what the `Char` production excludes before escaping, because a part carrying
// one is malformed and PowerPoint offers to repair the deck. Their set covered the C0 controls and
// missed the two non-characters U+FFFE and U+FFFF: `addText` wrote both raw into `<a:t>`, and an
// object name carrying one drew no warning.
//
// xmldom accepts both characters even with `onErrorStopParsing`, so a parse cannot be the strict
// check here. The production itself can: the part's every code point is walked against it.

import { describe, expect, test } from 'vitest'
import { build, readEntry, setDiagnosticHandler } from '../../helpers.js'

const FFFE = String.fromCharCode(0xfffe)
const FFFF = String.fromCharCode(0xffff)

/**
 * The code points of `text` outside XML 1.0's `Char` production, as hex.
 * @param {string} text
 * @returns {string[]}
 */
function forbiddenCodePoints(text) {
	const found = []
	for (const char of text) {
		const cp = char.codePointAt(0) ?? 0
		const allowed =
			cp === 0x9 ||
			cp === 0xa ||
			cp === 0xd ||
			(cp >= 0x20 && cp <= 0xd7ff) ||
			(cp >= 0xe000 && cp <= 0xfffd) ||
			(cp >= 0x10000 && cp <= 0x10ffff)
		if (!allowed) found.push(cp.toString(16))
	}
	return found
}

describe('characters XML 1.0 forbids', () => {
	test('are stripped from run text, so the slide part stays within the Char production', async () => {
		const { zip } = await build((pres) => {
			pres.addSlide().addText(`a${FFFF}b${FFFE}c`, { x: 1, y: 1, w: 4, h: 1 })
		})
		const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
		expect(forbiddenCodePoints(xml)).toEqual([])
		expect(xml).toContain('<a:t>abc</a:t>')
	})

	test('in an object name are reported before they are stripped', async () => {
		/** @type {string[]} */
		const codes = []
		setDiagnosticHandler((diagnostic) => codes.push(diagnostic.code))
		try {
			await build((pres) => {
				pres.addSlide().addText('x', { x: 1, y: 1, w: 4, h: 1, objectName: `Box${FFFF}` })
			})
		} finally {
			setDiagnosticHandler(null)
		}
		expect(codes).toContain('object-name/control-characters')
	})
})
