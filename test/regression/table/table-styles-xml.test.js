import { describe, expect, test } from 'vitest'
import { makeXmlTableStyles } from '../../../src/gen/pres/table-styles.ts'
import { TableStyle } from '../../../src/enums.ts'

// `ppt/tableStyles.xml` is a fixed stub: it names a default style id and defines nothing.
//
// This file used to characterize the custom-style payload `defineTableStyle()` wrote here —
// region order, side order, the tcTxStyle-before-tcStyle sequence, styleName escaping. All of
// it went once rendering in PowerPoint desktop 16.0 showed that a definition in this part is
// never read: PowerPoint resolves `<a:tableStyleId>` against its own gallery, so a built-in
// GUID paints with nothing here, and a custom GUID paints nothing however complete the
// definition. The byte-identity harness still cannot see this part change — no showcase deck
// ever produced a non-stub one — so what remains is the guard that it stays a stub.

describe('makeXmlTableStyles', () => {
	test('emits a self-closing tblStyleLst naming the default style id', () => {
		expect(makeXmlTableStyles()).toContain(
			`<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="${TableStyle.MEDIUM_STYLE_2_ACCENT_1}"/>`
		)
	})

	test('defines no styles of its own', () => {
		// A change that starts writing definitions here again is reintroducing markup PowerPoint
		// will not read, which is exactly the defect the removal closed.
		const xml = makeXmlTableStyles()
		expect(xml).not.toContain('<a:tblStyle ')
		expect(xml).not.toContain('<a:tcTxStyle')
		expect(xml).not.toContain('<a:tcStyle')
		expect(xml).not.toContain('<a:tcBdr')
	})

	test('takes no arguments — there is nothing per-deck to vary', () => {
		expect(makeXmlTableStyles.length).toBe(0)
		expect(makeXmlTableStyles()).toBe(makeXmlTableStyles())
	})

	test('is a complete XML document', () => {
		expect(makeXmlTableStyles().startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')).toBe(true)
	})
})
