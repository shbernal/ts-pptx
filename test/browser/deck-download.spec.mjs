import { expect, test } from '@playwright/test'
import JSZip from 'jszip'
import { buildDeckInBrowser } from './helpers.mjs'

/**
 * The browser can build a deck at all.
 *
 * Everything asserted here is downstream of `src/runtime/browser.ts:writeFile` — the
 * object-URL `<a download>` path, which until this lane existed had never run anywhere
 * but a human's tab. It is also the path the closed `writeFile`-hangs report was about.
 *
 * Read with jszip on purpose, not with the `fflate` the library writes with: an
 * independent zip implementation is what makes "the download is a real OPC package"
 * an assertion rather than a tautology.
 */
test('the demo builds a downloadable .pptx package in the browser', async ({ page }) => {
	const { bytes, fileName } = await buildDeckInBrowser(page)

	expect(fileName).toBe('Meridian_Q3_Business_Review.pptx')
	expect(bytes.byteLength).toBeGreaterThan(10_000)

	const zip = await JSZip.loadAsync(bytes)
	const names = Object.keys(zip.files)

	// The three parts that make it an OPC package PowerPoint will open at all.
	expect(names).toContain('[Content_Types].xml')
	expect(names).toContain('_rels/.rels')
	expect(names).toContain('ppt/presentation.xml')

	// The showcase is eleven slides. A package that unzips but lost its slides would
	// otherwise pass every assertion above.
	const slides = names.filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
	expect(slides).toHaveLength(11)

	const presentation = await zip.file('ppt/presentation.xml').async('string')
	expect(presentation).toContain('<p:sldIdLst>')
	expect(presentation).toContain('http://schemas.openxmlformats.org/presentationml/2006/main')
})
