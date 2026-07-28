import { describe, test } from 'vitest'
import { assert } from '../helpers.js'
import { parseCssPx, parseCssWidthBasis, pickColWidthBasis, readCellText } from '../../src/gen/table/html-dom.ts'

// Acceptance: HTML-table conversion must produce usable column widths and cell text on a DOM
// with no layout engine. `offsetWidth` is 0 for every cell there, which made the proportional
// calc a 0/0 divide and emitted a zero-width table; `innerText` is absent on some DOMs, which
// emptied every cell. These are the three DOM-independent decisions behind that, unit-tested
// directly (the pattern docs/project-target.md prescribes for this file).

describe('parseCssPx', () => {
	test('parses a px length', () => {
		assert(parseCssPx('120px') === 120, 'px length must parse to its magnitude')
	})

	test('preserves a fractional px length', () => {
		assert(parseCssPx('1.5px') === 1.5, 'the decimal point must not be stripped')
	})

	test('parses a leading-dot fraction', () => {
		assert(parseCssPx('.5px') === 0.5, '".5px" is a valid CSS length')
	})

	test('parses a bare number as px', () => {
		assert(parseCssPx('12') === 12, 'a unitless number reads as px')
	})

	test('zero parses to zero, not NaN', () => {
		assert(parseCssPx('0px') === 0, '0px is a real length')
	})

	test('an empty value is NaN', () => {
		assert(Number.isNaN(parseCssPx('')), 'absent value has no magnitude')
	})

	test('a keyword is NaN', () => {
		assert(Number.isNaN(parseCssPx('auto')), 'auto is not an absolute length')
	})

	test('a percentage is NaN', () => {
		assert(Number.isNaN(parseCssPx('30%')), 'a percentage is not an absolute length')
	})

	test('a non-px unit is NaN', () => {
		assert(Number.isNaN(parseCssPx('3em')), 'em cannot be resolved without a font context')
	})

	test('a nullish value is NaN rather than a throw', () => {
		assert(Number.isNaN(parseCssPx(undefined)), 'a missing property must not throw')
	})
})

describe('parseCssWidthBasis', () => {
	test('an all-px set parses to its magnitudes', () => {
		const basis = parseCssWidthBasis(['100px', '200px', '100px'])
		assert(basis.join() === '100,200,100', `expected 100,200,100, got ${basis.join()}`)
	})

	test('an all-percent set is a valid proportional basis', () => {
		const basis = parseCssWidthBasis(['25%', '25%', '50%'])
		assert(basis.join() === '25,25,50', `percentages are usable proportionally, got ${basis.join()}`)
	})

	test('a mixed-unit set is rejected', () => {
		const basis = parseCssWidthBasis(['100px', '50%'])
		assert(basis.length === 0, 'px and % cannot share a basis')
	})

	test('one unparseable entry rejects the whole set', () => {
		const basis = parseCssWidthBasis(['100px', 'auto', '100px'])
		assert(basis.length === 0, 'a partial basis would zero the unparseable column')
	})

	test('an empty-string entry rejects the whole set', () => {
		assert(parseCssWidthBasis(['100px', '']).length === 0, 'an unstated width is not zero')
	})

	test('a negative magnitude rejects the whole set', () => {
		assert(parseCssWidthBasis(['100px', '-20px']).length === 0, 'a negative column has no proportional meaning')
	})

	test('bare numbers are accepted alongside px', () => {
		const basis = parseCssWidthBasis(['100', '200px'])
		assert(basis.join() === '100,200', `a unitless number is px, got ${basis.join()}`)
	})

	test('an all-zero set parses (the caller decides it is unusable)', () => {
		const basis = parseCssWidthBasis(['0px', '0px'])
		assert(basis.join() === '0,0', 'zeros parse; pickColWidthBasis is what rejects them')
	})

	test('no columns yields no basis', () => {
		assert(parseCssWidthBasis([]).length === 0, 'empty in, empty out')
	})
})

describe('pickColWidthBasis', () => {
	test('measured widths win when the table was laid out', () => {
		const basis = pickColWidthBasis([120, 240], [50, 50])
		assert(basis.join() === '120,240', `offsetWidth is the best basis, got ${basis.join()}`)
	})

	test('CSS widths are used when nothing was laid out', () => {
		const basis = pickColWidthBasis([0, 0, 0], [100, 200, 100])
		assert(basis.join() === '100,200,100', `expected the CSS basis, got ${basis.join()}`)
	})

	test('equal split when neither basis carries width', () => {
		const basis = pickColWidthBasis([0, 0, 0], [])
		assert(basis.join() === '1,1,1', `expected an all-ones basis, got ${basis.join()}`)
	})

	test('equal split when the CSS basis is all zeros', () => {
		const basis = pickColWidthBasis([0, 0], [0, 0])
		assert(basis.join() === '1,1', `zero-width CSS is no basis at all, got ${basis.join()}`)
	})

	test('a CSS basis of the wrong length is ignored', () => {
		const basis = pickColWidthBasis([0, 0, 0], [100, 200])
		assert(basis.join() === '1,1,1', `a misaligned basis must not be applied, got ${basis.join()}`)
	})

	test('non-finite measured entries do not defeat the measured arm', () => {
		const basis = pickColWidthBasis([NaN, 240], [50, 50])
		assert(basis.length === 2 && basis[1] === 240, 'a partly-measured table still uses offsetWidth')
	})

	test('an all-NaN measured vector falls through to CSS', () => {
		const basis = pickColWidthBasis([NaN, NaN], [10, 30])
		assert(basis.join() === '10,30', `NaN sums to 0, so CSS must win, got ${basis.join()}`)
	})

	test('the measured array is copied, not aliased', () => {
		const measured = [120, 240]
		const basis = pickColWidthBasis(measured, [])
		basis[0] = 999
		assert(measured[0] === 120, 'the caller"s array must not be mutated through the basis')
	})

	test('no columns yields no basis', () => {
		assert(pickColWidthBasis([], []).length === 0, 'empty in, empty out')
	})
})

// Minimal structural stand-ins for DOM nodes — `readCellText` is typed against the shape it
// actually reads, so the fallback walk can be exercised with no DOM implementation at all.
const textNode = (data) => ({ nodeType: 3, nodeValue: data })
const element = (nodeName, childNodes = []) => ({ nodeType: 1, nodeName, childNodes })
const BR = element('BR')

describe('readCellText', () => {
	test('innerText is used verbatim when the DOM provides it', () => {
		const cell = { innerText: 'Line 1\nLine 2', childNodes: [textNode('ignored')] }
		assert(readCellText(cell) === 'Line 1\nLine 2', 'the rendered text is authoritative')
	})

	test('an empty innerText is still authoritative', () => {
		const cell = { innerText: '', childNodes: [textNode('hidden')] }
		assert(readCellText(cell) === '', 'an empty string is a value, not an absence')
	})

	test('the fallback concatenates text nodes', () => {
		const cell = element('TD', [textNode('Hello '), element('SPAN', [textNode('world')])])
		assert(readCellText(cell) === 'Hello world', `expected "Hello world", got ${JSON.stringify(readCellText(cell))}`)
	})

	test('the fallback maps <br> to a newline', () => {
		const cell = element('TD', [textNode('Line 1'), BR, textNode('Line 2')])
		assert(readCellText(cell) === 'Line 1\nLine 2', `got ${JSON.stringify(readCellText(cell))}`)
	})

	test('the fallback finds a <br> nested inside an element', () => {
		const cell = element('TD', [element('SPAN', [textNode('a'), BR, textNode('b')])])
		assert(readCellText(cell) === 'a\nb', `got ${JSON.stringify(readCellText(cell))}`)
	})

	test('a lowercase br tag name is matched too', () => {
		const cell = element('TD', [textNode('a'), element('br'), textNode('b')])
		assert(
			readCellText(cell) === 'a\nb',
			`tag-name case must not decide this, got ${JSON.stringify(readCellText(cell))}`
		)
	})

	test('the fallback collapses whitespace and trims per line', () => {
		const cell = element('TD', [textNode('\n\t  Hello   there  '), BR, textNode('  World \n')])
		assert(readCellText(cell) === 'Hello there\nWorld', `got ${JSON.stringify(readCellText(cell))}`)
	})

	test('the fallback ignores comment nodes', () => {
		const cell = element('TD', [textNode('a'), { nodeType: 8, nodeValue: ' note ' }, textNode('b')])
		assert(readCellText(cell) === 'ab', `comments are not text, got ${JSON.stringify(readCellText(cell))}`)
	})

	test('an empty cell yields an empty string', () => {
		assert(readCellText(element('TD')) === '', 'no children means no text')
	})

	test('a cell with no childNodes at all does not throw', () => {
		assert(readCellText({ nodeType: 1, nodeName: 'TD' }) === '', 'a childless shape must degrade, not throw')
	})
})
