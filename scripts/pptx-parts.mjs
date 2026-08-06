/**
 * Shared OPC package explode/normalize/diff machinery.
 *
 * Extracted from `scripts/byte-identity.mjs` when the browser lane needed the same
 * comparison. Two callers now assert "these two .pptx packages are the same bytes":
 *
 *   - `scripts/byte-identity.mjs` — same runtime, before vs after a refactor.
 *   - `test/browser/cross-runtime-bytes.spec.mjs` — same commit, Node vs a real browser.
 *
 * They must agree on what "the same" means, and in particular on the normalizer list:
 * a second, hand-rolled comparison would drift, and the way it drifts is silent — one
 * gate would start tolerating a difference the other still calls a regression, and
 * nobody would know which one was right.
 */

import fs from 'node:fs'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { ROOT } from './script-utils.mjs'

const SHOWCASES_ENTRY = path.join(ROOT, 'demos', 'showcases', 'lib', 'showcases.mjs')

/**
 * Emitted values that legitimately differ between two identical runs.
 * Deliberately narrow: normalizing ONLY these keeps a changed *fixed* GUID
 * (e.g. a built-in table-style id) visible as a real diff.
 */
export const NORMALIZERS = [
	// core.xml timestamps — the deck's and every embedded workbook's
	[
		/<dcterms:(created|modified) xsi:type="dcterms:W3CDTF">[^<]*<\/dcterms:\1>/g,
		'<dcterms:$1 xsi:type="dcterms:W3CDTF">NORMALIZED-TIMESTAMP</dcterms:$1>',
	],
	// presentation.xml section ids — random GUID per run
	[/(<p14:section[^>]*\bid=")\{[^}]*\}"/g, '$1{NORMALIZED-SECTION}"'],
	// chartN.xml uniqueId — random GUID per run
	[/(<c16:uniqueId[^>]*\bval=")\{[^}]*\}"/g, '$1{NORMALIZED-UNIQUEID}"'],
]

/** Apply every normalizer to one part's text. */
export function normalize(text) {
	return NORMALIZERS.reduce((out, [re, sub]) => out.replace(re, sub), text)
}

/**
 * The showcase registry, loaded by URL rather than by bare specifier.
 *
 * `demos/showcases` is a workspace package the root does not depend on, so a bare
 * import would not resolve from here. Loading it by file URL also keeps it out of
 * the typechecked module graph — the demo decks are plain untyped ESM that no
 * tsconfig includes.
 */
export async function loadShowcases() {
	const { SHOWCASES } = await import(pathToFileURL(SHOWCASES_ENTRY).href)
	if (!Array.isArray(SHOWCASES) || SHOWCASES.length === 0)
		throw new Error('no showcases registered in ' + path.relative(ROOT, SHOWCASES_ENTRY))
	return SHOWCASES
}

/** One showcase by slug, or a throw naming the ones that exist. */
export async function loadShowcase(slug) {
	const showcases = await loadShowcases()
	const found = showcases.find((showcase) => showcase.slug === slug)
	if (!found) throw new Error('no showcase with slug "' + slug + '"; have: ' + showcases.map((s) => s.slug).join(', '))
	return found
}

/** fflate's `unzipSync`, loaded the same way `byte-identity.mjs` has always loaded it. */
async function unzipSync() {
	const fflate = await import(pathToFileURL(path.join(ROOT, 'node_modules', 'fflate', 'esm', 'browser.js')).href)
	return fflate.unzipSync
}

/**
 * Explode one `.pptx` into `destDir`, recursing into embedded `.xlsx` parts.
 *
 * XML parts are written through `normalize()`; everything else is written verbatim.
 * Each embedded workbook is its own OPC zip, so it is recursed into rather than
 * diffed as opaque compressed bytes.
 */
export async function explodePackage(bytes, destDir) {
	const unzip = await unzipSync()
	const decoder = new TextDecoder('utf-8')

	const dump = (zipBytes, dir) => {
		const entries = unzip(zipBytes)
		for (const name of Object.keys(entries).sort()) {
			const partBytes = entries[name]
			if (/\.xlsx$/i.test(name)) {
				dump(partBytes, path.join(dir, name + '!'))
				continue
			}
			const dest = path.join(dir, name)
			fs.mkdirSync(path.dirname(dest), { recursive: true })
			if (/\.(xml|rels)$/i.test(name)) fs.writeFileSync(dest, normalize(decoder.decode(partBytes)), 'utf8')
			else fs.writeFileSync(dest, partBytes)
		}
	}

	fs.rmSync(destDir, { recursive: true, force: true })
	fs.mkdirSync(destDir, { recursive: true })
	dump(bytes, destDir)
	return destDir
}

/** Every part path under an exploded package directory, depth-first and sorted. */
export function listParts(dir) {
	const out = []
	const walk = (d, prefix) => {
		for (const entry of fs.readdirSync(d, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) {
			const rel = prefix ? prefix + '/' + entry.name : entry.name
			if (entry.isDirectory()) walk(path.join(d, entry.name), rel)
			else out.push(rel)
		}
	}
	walk(dir, '')
	return out
}

/** Compare two exploded packages. Returns a list of human-readable differences. */
export function diffParts(baseDir, curDir) {
	const base = new Set(listParts(baseDir))
	const cur = new Set(listParts(curDir))
	const diffs = []
	for (const part of base) if (!cur.has(part)) diffs.push('REMOVED  ' + part)
	for (const part of cur) if (!base.has(part)) diffs.push('ADDED    ' + part)
	for (const part of base) {
		if (!cur.has(part)) continue
		const a = fs.readFileSync(path.join(baseDir, part))
		const b = fs.readFileSync(path.join(curDir, part))
		if (!a.equals(b)) diffs.push('CHANGED  ' + part)
	}
	return diffs.sort()
}
