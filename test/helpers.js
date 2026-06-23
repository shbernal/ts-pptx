// Tests intentionally read the generated .pptx with jszip rather than the
// library's own src/zip.ts (fflate). The write path uses fflate, so reading
// back with a *different* zip implementation makes jszip an independent oracle:
// a round-trip bug in fflate can't mask itself by being used on both sides.
// Keep jszip as a devDep for this reason — do not "consolidate" onto src/zip.ts.
import JSZip from 'jszip'
import PptxGenJS from '../dist/node.js'
import { describe, test } from 'vitest'

async function build(buildFn) {
	const pres = new PptxGenJS()
	buildFn(pres)
	const buf = await pres.stream()
	const zip = await JSZip.loadAsync(buf)
	return { pres, zip, buf }
}

async function readEntry(zip, path) {
	const entry = zip.file(path)
	if (!entry) throw new Error('zip entry not found: ' + path)
	return entry.async('string')
}

function listEntries(zip) {
	return Object.keys(zip.files)
}

function defineRegressionSuite(suiteName, legacyIssueOrCases, maybeCases) {
	const cases = Array.isArray(legacyIssueOrCases) ? legacyIssueOrCases : maybeCases
	if (!Array.isArray(cases)) throw new Error('defineRegressionSuite requires an array of test cases')

	describe(suiteName, () => {
		for (const fixture of cases) {
			test(fixture.name, async () => {
				await fixture.fn()
			})
		}
	})
}

function assert(cond, msg) {
	if (!cond) throw new Error('assertion failed: ' + msg)
}

function assertEqual(actual, expected, msg) {
	if (actual !== expected)
		throw new Error(
			'assertion failed: ' + (msg || '') + ' expected ' + JSON.stringify(expected) + ' got ' + JSON.stringify(actual)
		)
}

function assertIncludes(haystack, needle, label) {
	assert(haystack.includes(needle), `expected ${label || 'value'} to include ${needle}; got: ${haystack}`)
}

function assertNotIncludes(haystack, needle, label) {
	assert(!haystack.includes(needle), `expected ${label || 'value'} not to include ${needle}; got: ${haystack}`)
}

function xmlBlocks(xml, tagName) {
	const escapedName = tagName.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')
	const re = new RegExp(`<${escapedName}\\b[\\s\\S]*?<\\/${escapedName}>`, 'g')
	return xml.match(re) || []
}

function firstXmlBlock(xml, tagName, label = tagName) {
	const block = xmlBlocks(xml, tagName)[0]
	assert(block, `expected ${label} block in XML; got: ${xml}`)
	return block
}

function xmlAttributes(tag) {
	const attrs = {}
	for (const match of tag.matchAll(/\s([\w:-]+)="([^"]*)"/g)) {
		attrs[match[1]] = match[2]
	}
	return attrs
}

function selfClosingTags(xml, tagName) {
	const escapedName = tagName.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')
	const re = new RegExp(`<${escapedName}\\b[^>]*/>`, 'g')
	return xml.match(re) || []
}

function xmlOpeningTags(xml, tagName) {
	const escapedName = tagName.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')
	const re = new RegExp(`<${escapedName}\\b[^>]*(?:/>|>)`, 'g')
	return xml.match(re) || []
}

function contentTypeDefaultExtensions(xml) {
	return selfClosingTags(xml, 'Default').map((tag) => xmlAttributes(tag).Extension)
}

function contentTypeOverrideParts(xml) {
	return selfClosingTags(xml, 'Override').map((tag) => xmlAttributes(tag).PartName)
}

function contentTypeForExtension(xml, extension) {
	const tag = selfClosingTags(xml, 'Default').find((t) => xmlAttributes(t).Extension === extension)
	return tag ? xmlAttributes(tag).ContentType : undefined
}

function assertContentTypeDefault(xml, extension) {
	const extensions = contentTypeDefaultExtensions(xml)
	assert(
		extensions.includes(extension),
		`expected Content_Types Default for ${extension}; got: ${extensions.join(', ')}`
	)
}

function assertNoContentTypeDefault(xml, extension) {
	const extensions = contentTypeDefaultExtensions(xml)
	assert(
		!extensions.includes(extension),
		`did not expect Content_Types Default for ${extension}; got: ${extensions.join(', ')}`
	)
}

function assertContentTypeOverride(xml, partName) {
	const parts = contentTypeOverrideParts(xml)
	assert(parts.includes(partName), `expected Content_Types Override for ${partName}; got: ${parts.join(', ')}`)
}

function assertXmlOrder(xml, before, after, label) {
	const beforeIndex = xml.indexOf(before)
	const afterIndex = xml.indexOf(after)
	assert(beforeIndex !== -1, `expected ${before} in ${label || 'XML'}; got: ${xml}`)
	assert(afterIndex !== -1, `expected ${after} in ${label || 'XML'}; got: ${xml}`)
	assert(
		beforeIndex < afterIndex,
		`expected ${before} before ${after} in ${label || 'XML'}; got order ${beforeIndex} then ${afterIndex}: ${xml}`
	)
}

function nonVisualDrawingProperties(xml) {
	const tags = xmlOpeningTags(xml, 'p:cNvPr')
	return tags.map((tag) => ({ tag, attrs: xmlAttributes(tag) }))
}

function findNonVisualDrawingProperty(xml, attrs) {
	return nonVisualDrawingProperties(xml).find(({ attrs: actual }) =>
		Object.entries(attrs).every(([name, value]) => actual[name] === value)
	)
}

function assertNonVisualDrawingProperty(xml, attrs, label) {
	const match = findNonVisualDrawingProperty(xml, attrs)
	assert(match, `expected ${label || 'p:cNvPr'} with ${JSON.stringify(attrs)}; got: ${xml}`)
	return match
}

export {
	PptxGenJS,
	build,
	readEntry,
	listEntries,
	defineRegressionSuite,
	assert,
	assertEqual,
	assertIncludes,
	assertNotIncludes,
	xmlBlocks,
	firstXmlBlock,
	xmlAttributes,
	selfClosingTags,
	contentTypeDefaultExtensions,
	contentTypeForExtension,
	contentTypeOverrideParts,
	assertContentTypeDefault,
	assertNoContentTypeDefault,
	assertContentTypeOverride,
	assertXmlOrder,
	nonVisualDrawingProperties,
	findNonVisualDrawingProperty,
	assertNonVisualDrawingProperty,
	xmlOpeningTags,
}
