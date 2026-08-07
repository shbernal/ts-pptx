/**
 * Read-blindness census — which OOXML elements appear in real PowerPoint fixtures but are
 * never addressed by any read accessor.
 *
 * Why this exists. A *write*-API gap is visible loss: a converter sees the construct and can
 * report it. A *read*-API gap is invisible — nothing upstream knows the element was there, so
 * a round-trip diff comes back clean and certifies a deck it never inspected. No amount of
 * care downstream detects that class of defect; only this census does.
 *
 * Method. Two sets, subtracted:
 *   present  — every distinct element QName reachable from the fixture parts, walked through
 *              the documented raw hatch (`element_` / `part.dom`).
 *   consumed — element names mentioned anywhere in `src/read/`. The read path addresses
 *              elements three different ways, so both forms are harvested:
 *                `firstChild(sp, 'p:nvSpPr')`               -> qname literal
 *                `el.localName === 'par' && el.namespaceURI === OOXML_NS.p`
 *                `switch (name) { case 'lumMod': … }`       -> bare localName literal
 *              Parsed from source, so the set cannot drift from the code.
 *
 * The residue is reported at three severities, which are genuinely different problems:
 *   UNADDRESSABLE — the element's namespace appears NOWHERE in `src/read/`: not in the
 *                   shared `OOXML_NS` map, and not as a module-local URI constant either
 *                   (several accessors declare their own, e.g. `ASVG_NS`, `ADEC_NS`, and
 *                   reach the element by `localName` + `namespaceURI` comparison). Since
 *                   `firstChild`/`getElements` throw on an unknown prefix, an element here
 *                   cannot be reached by any means without a source change.
 *   UNREAD        — neither the qname nor the bare local name appears anywhere in
 *                   `src/read/`. High confidence that nothing reads it.
 *   WEAK          — the qname never appears, but the bare local name does somewhere. Either
 *                   a namespace-comparison accessor reads it, or the name merely collides
 *                   with an unrelated string. Needs manual triage; listed, not counted.
 *
 * `consumed` is a deliberate over-approximation: a name in a comment or an error string
 * counts as consumed. That biases the census toward *under*-reporting blindness, so a
 * reported element is a real finding while a silent one is not proof of coverage.
 *
 * Scope caveat. The default corpus is `test/read/fixtures/`, which is *construct-targeted*:
 * each deck exercises one feature. So it measures **coverage**, not **frequency** — it
 * cannot tell you whether `custGeom` or placeholder inheritance is worth the next week of
 * work. Point `--dir` at a corpus of real decks — any path, including one outside the
 * repo — to get a frequency-weighted read instead.
 *
 * Usage:
 *   node scripts/read-blindness-census.mjs            # slides only (the default surface)
 *   node scripts/read-blindness-census.mjs --all      # + layouts, masters, theme, notes
 *   node scripts/read-blindness-census.mjs --json     # machine-readable, for a test to assert on
 *   node scripts/read-blindness-census.mjs --fixture mixed.pptx
 *   node scripts/read-blindness-census.mjs --dir ~/decks     # frequency over real decks
 */
import fs from 'node:fs/promises'
import path from 'node:path'
import { ROOT } from './script-utils.mjs'
import { Presentation } from '../dist/read.js'

const DEFAULT_DIR = path.join('test', 'read', 'fixtures')
const READ_SRC = path.join(ROOT, 'src', 'read')
const ELEMENT_NODE = 1

/** Parse the read path's own namespace map so this script cannot drift from it. */
async function readNamespaceMap() {
	const source = await fs.readFile(path.join(READ_SRC, 'oxml', 'dom.ts'), 'utf8')
	const block = /export const OOXML_NS = Object\.freeze\(\{([\s\S]*?)\}\)/.exec(source)
	if (!block?.[1]) throw new Error('could not locate OOXML_NS in src/read/oxml/dom.ts')
	const uriToPrefix = new Map()
	for (const m of block[1].matchAll(/^\s*([A-Za-z][\w]*)\s*:\s*'([^']+)'/gm)) {
		uriToPrefix.set(m[2], m[1])
	}
	if (uriToPrefix.size === 0) throw new Error('OOXML_NS parsed empty')
	return uriToPrefix
}

/**
 * Element names mentioned in src/read/, in both addressing forms:
 * `qnames` — `'p:nvSpPr'` style literals, used by firstChild/getElements.
 * `locals` — bare `'par'` / `'lumMod'` literals, used by localName comparisons and switches.
 */
async function readConsumedNames(knownPrefixes) {
	const qnames = new Set()
	const locals = new Set()
	const uris = new Set()
	async function walk(dir) {
		for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
			const full = path.join(dir, entry.name)
			if (entry.isDirectory()) await walk(full)
			else if (entry.name.endsWith('.ts')) {
				const text = await fs.readFile(full, 'utf8')
				for (const m of text.matchAll(/['"`]([A-Za-z][\w]*):([A-Za-z][\w-]*)['"`]/g)) {
					if (knownPrefixes.has(m[1])) {
						qnames.add(`${m[1]}:${m[2]}`)
						locals.add(m[2])
					}
				}
				for (const m of text.matchAll(/['"`]([A-Za-z][\w-]*)['"`]/g)) locals.add(m[1])
				// Module-local namespace constants (ASVG_NS, ADEC_NS, …): an element in one of
				// these is reachable by localName + namespaceURI comparison, so it is not
				// structurally unaddressable even though its prefix is absent from OOXML_NS.
				for (const m of text.matchAll(/['"`](https?:\/\/[^'"`\s]+)['"`]/g)) uris.add(m[1])
			}
		}
	}
	await walk(READ_SRC)
	return { qnames, locals, uris }
}

function collect(node, uriToPrefix, tally) {
	const uri = node.namespaceURI || ''
	const prefix = uriToPrefix.get(uri)
	// An element outside the read path's namespace map cannot be addressed by qname at all.
	const key = prefix ? `${prefix}:${node.localName}` : `{${uri}}${node.localName}`
	tally.set(key, (tally.get(key) || 0) + 1)
	for (let child = node.firstChild; child; child = child.nextSibling) {
		if (child.nodeType === ELEMENT_NODE) collect(child, uriToPrefix, tally)
	}
}

/** The parts to walk for one deck. Slides are the default surface; --all adds the chrome. */
function surfaceOf(pres, includeChrome) {
	const parts = []
	for (const slide of pres.slides) {
		parts.push({ label: 'slide', element: slide.element_ })
		const notes = typeof slide.notesSlide === 'function' ? slide.notesSlide() : slide.notesSlide
		if (includeChrome && notes?.element_) parts.push({ label: 'notesSlide', element: notes.element_ })
	}
	if (includeChrome) {
		for (const layout of pres.layouts()) {
			const el = pres.opc.part(layout.partName)?.dom?.documentElement
			if (el) parts.push({ label: 'slideLayout', element: el })
		}
		for (const [partName, part] of pres.opc.parts) {
			if (/slideMaster\d+\.xml$|theme\d+\.xml$/.test(partName)) {
				const el = part.dom?.documentElement
				if (el) parts.push({ label: path.basename(partName), element: el })
			}
		}
	}
	return parts
}

async function main() {
	const argv = process.argv.slice(2)
	const includeChrome = argv.includes('--all')
	const asJson = argv.includes('--json')
	const only = argv[argv.indexOf('--fixture') + 1]
	const dirArg = argv.indexOf('--dir')
	const corpusDir = path.resolve(ROOT, (dirArg === -1 ? null : argv[dirArg + 1]) ?? DEFAULT_DIR)

	const uriToPrefix = await readNamespaceMap()
	const knownPrefixes = new Set(uriToPrefix.values())
	const { qnames: consumed, locals: consumedLocals, uris: consumedUris } = await readConsumedNames(knownPrefixes)

	const fixtures = (await fs.readdir(corpusDir))
		.filter((f) => f.endsWith('.pptx'))
		.filter((f) => !only || argv.indexOf('--fixture') === -1 || f === only)
		.sort()

	const totals = new Map() // qname -> { count, fixtures:Set }
	const failures = []
	for (const name of fixtures) {
		let pres
		try {
			pres = await Presentation.load(await fs.readFile(path.join(corpusDir, name)))
		} catch (err) {
			failures.push(`${name}: ${err instanceof Error ? err.message : String(err)}`)
			continue
		}
		const tally = new Map()
		for (const { element } of surfaceOf(pres, includeChrome)) collect(element, uriToPrefix, tally)
		for (const [qname, count] of tally) {
			const entry = totals.get(qname) || { count: 0, fixtures: new Set() }
			entry.count += count
			entry.fixtures.add(name)
			totals.set(qname, entry)
		}
	}

	const unaddressable = []
	const unread = []
	const weak = []
	for (const [qname, entry] of totals) {
		if (consumed.has(qname)) continue
		const row = { qname, occurrences: entry.count, fixtures: [...entry.fixtures].sort() }
		const braced = /^\{([^}]*)\}(.+)$/.exec(qname)
		if (braced && !consumedUris.has(braced[1])) {
			// Namespace absent from src/read/ entirely — no accessor can name it.
			unaddressable.push(row)
			continue
		}
		// A bare-localName mention means a namespace-comparison accessor may well read it.
		const local = braced ? braced[2] : qname.slice(qname.indexOf(':') + 1)
		;(consumedLocals.has(local) ? weak : unread).push(row)
	}
	const byFrequency = (a, b) => b.occurrences - a.occurrences || a.qname.localeCompare(b.qname)
	unaddressable.sort(byFrequency)
	unread.sort(byFrequency)
	weak.sort(byFrequency)

	const summary = {
		corpus: path.relative(ROOT, corpusDir) || '.',
		fixtures: fixtures.length,
		surface: includeChrome ? 'slides + layouts + masters + theme + notes' : 'slides',
		distinctPresent: totals.size,
		consumedQNames: consumed.size,
		consumedLocals: consumedLocals.size,
		unaddressable,
		unread,
		weak,
		unreadable: failures,
	}

	if (asJson) {
		console.log(JSON.stringify(summary, null, 2))
		return
	}

	console.log(`read-blindness census — ${fixtures.length} decks from ${summary.corpus}, surface: ${summary.surface}`)
	if (!fixtures.length) {
		console.log('  (no .pptx files found — nothing to census)')
		return
	}
	console.log(
		`  ${totals.size} distinct element qnames present; ` +
			`${consumed.size} qname literals + ${consumedLocals.size} bare-name literals in src/read/`
	)
	if (failures.length) {
		console.log(`\n  ${failures.length} fixture(s) failed to load:`)
		for (const f of failures) console.log(`    ${f}`)
	}

	const section = (title, rows, note) => {
		console.log(`\n${title} — ${rows.length}`)
		if (note) console.log(`  ${note}`)
		if (!rows.length) {
			console.log('  (none)')
			return
		}
		for (const r of rows) {
			const where = r.fixtures.length <= 3 ? r.fixtures.join(', ') : `${r.fixtures.length} fixtures`
			console.log(`  ${String(r.occurrences).padStart(5)}×  ${r.qname.padEnd(34)} ${where}`)
		}
	}
	section(
		'UNADDRESSABLE (namespace appears nowhere in src/read/)',
		unaddressable,
		'Not in OOXML_NS and not a module-local URI constant — unreachable by any accessor without a source change.'
	)
	section('UNREAD (neither qname nor bare local name appears in src/read/)', unread)
	section(
		'WEAK (qname absent, but the bare local name appears somewhere)',
		weak,
		'Either a localName/namespaceURI comparison reads it, or the name collides with an unrelated string. Triage by hand.'
	)
	console.log(
		`\nNote: "consumed" counts any name literal in src/read/, including ones in comments and error ` +
			`strings, so this census UNDER-reports blindness. A listed element is a real gap; an absent one is not proof.`
	)
}

await main()
