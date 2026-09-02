#!/usr/bin/env node
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
 * Usage: `pnpm run read:census -- --help` (the flag list lives in USAGE below, so there is
 * one copy of it to keep true).
 */
import fs from 'node:fs/promises'
import path from 'node:path'
import { FIXTURES_DIR, ROOT, corpusDecks, parseCliOrExit, resolveCorpusDir } from './script-utils.mjs'
import { Presentation } from '../dist/read.js'

/** The parsed-XML element type the read model exposes — xmldom's, not the DOM's. */
/** @typedef {import('@xmldom/xmldom').Element} XmlElement */

const READ_SRC = path.join(ROOT, 'src', 'read')
const ELEMENT_NODE = 1

/**
 * Parse the library's own namespace map so this script cannot drift from it.
 *
 * The registry moved from `src/read/oxml/dom.ts` (which now re-exports it) to
 * `src/ooxml/namespaces.ts`, so that both halves of the library could reach it. Nothing runs
 * this census in a gate, so the throw below went unseen: read it as the reason the source path
 * is named once, here.
 */
const NAMESPACES_SRC = path.join(ROOT, 'src', 'ooxml', 'namespaces.ts')

async function readNamespaceMap() {
	const source = await fs.readFile(NAMESPACES_SRC, 'utf8')
	const block = /export const OOXML_NS = Object\.freeze\(\{([\s\S]*?)\}\)/.exec(source)
	if (!block?.[1]) throw new Error(`could not locate OOXML_NS in ${path.relative(ROOT, NAMESPACES_SRC)}`)
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
 * @param {Set<string>} knownPrefixes
 * @returns {Promise<{qnames: Set<string>, locals: Set<string>, uris: Set<string>}>}
 */
async function readConsumedNames(knownPrefixes) {
	/** @type {Set<string>} */
	const qnames = new Set()
	/** @type {Set<string>} */
	const locals = new Set()
	/** @type {Set<string>} */
	const uris = new Set()
	/**
	 * @param {string} dir
	 * @returns {Promise<void>}
	 */
	async function walk(dir) {
		for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
			const full = path.join(dir, entry.name)
			if (entry.isDirectory()) await walk(full)
			else if (entry.name.endsWith('.ts')) {
				const text = await fs.readFile(full, 'utf8')
				for (const m of text.matchAll(/['"`]([A-Za-z][\w]*):([A-Za-z][\w-]*)['"`]/g)) {
					const [, prefix = '', local = ''] = m
					if (knownPrefixes.has(prefix)) {
						qnames.add(`${prefix}:${local}`)
						locals.add(local)
					}
				}
				for (const m of text.matchAll(/['"`]([A-Za-z][\w-]*)['"`]/g)) if (m[1]) locals.add(m[1])
				// Module-local namespace constants (ASVG_NS, ADEC_NS, …): an element in one of
				// these is reachable by localName + namespaceURI comparison, so it is not
				// structurally unaddressable even though its prefix is absent from OOXML_NS.
				for (const m of text.matchAll(/['"`](https?:\/\/[^'"`\s]+)['"`]/g)) if (m[1]) uris.add(m[1])
			}
		}
	}
	await walk(READ_SRC)
	return { qnames, locals, uris }
}

/**
 * Tally every element in the subtree by the qname an accessor would have to name.
 * @param {XmlElement} node
 * @param {Map<string, string>} uriToPrefix
 * @param {Map<string, number>} tally
 * @returns {void}
 */
function collect(node, uriToPrefix, tally) {
	const uri = node.namespaceURI || ''
	const prefix = uriToPrefix.get(uri)
	// An element outside the read path's namespace map cannot be addressed by qname at all.
	const key = prefix ? `${prefix}:${node.localName}` : `{${uri}}${node.localName}`
	tally.set(key, (tally.get(key) || 0) + 1)
	for (let child = node.firstChild; child; child = child.nextSibling) {
		if (child.nodeType === ELEMENT_NODE) collect(/** @type {XmlElement} */ (child), uriToPrefix, tally)
	}
}

/**
 * The parts to walk for one deck. Slides are the default surface; --all adds the chrome.
 * @param {import('../dist/read.js').Presentation} pres
 * @param {boolean} includeChrome
 * @returns {{label: string, element: XmlElement}[]}
 */
function surfaceOf(pres, includeChrome) {
	/** @type {{label: string, element: XmlElement}[]} */
	const parts = []
	for (const slide of pres.slides) {
		parts.push({ label: 'slide', element: slide.element_ })
		// `notesSlide` is a getter, not a method: the `typeof … === 'function'` call form this
		// used to guard for evaluates the getter and then tests its *result*, so that arm was
		// unreachable and the getter ran twice on the way to the arm that is taken.
		const notes = slide.notesSlide
		// Through the notes slide's OPC part, the way layouts and masters are reached below:
		// `NotesSlide` exposes no `element_`, so the property this read for was always
		// `undefined` and `--all` never censused a notes slide at all.
		const notesEl = notes?.part.dom?.documentElement
		if (includeChrome && notesEl) parts.push({ label: 'notesSlide', element: notesEl })
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

const USAGE = `Read-blindness census — which OOXML the read model never looks at.

  pnpm run read:census
  pnpm run read:census -- --all
  pnpm run read:census -- --fixture table.pptx --json

Options:
  --all              include chrome elements, not just the content surface
  --fixture <name>   restrict the census to one .pptx by file name
  --dir <path>       corpus directory (default ${path.relative(ROOT, FIXTURES_DIR)})
  --json             machine-readable report on stdout
  -h, --help         show this message`

async function main() {
	// Previously `argv[argv.indexOf('--fixture') + 1]`, which resolves to argv[0] when the
	// flag is absent — the filter below had to carry a second `indexOf` check to undo it.
	const { values } = parseCliOrExit(process.argv.slice(2), {
		usage: USAGE,
		options: {
			all: { type: 'boolean', default: false },
			json: { type: 'boolean', default: false },
			fixture: { type: 'string' },
			dir: { type: 'string' },
		},
	})
	const includeChrome = values.all
	const asJson = values.json
	const only = values.fixture
	const corpusDir = resolveCorpusDir(values.dir)

	const uriToPrefix = await readNamespaceMap()
	const knownPrefixes = new Set(uriToPrefix.values())
	const { qnames: consumed, locals: consumedLocals, uris: consumedUris } = await readConsumedNames(knownPrefixes)

	const fixtures = await corpusDecks({ dir: corpusDir, only })

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

	/**
	 * One element name the corpus contains, with where it was seen.
	 * @typedef {{qname: string, occurrences: number, fixtures: string[]}} Row
	 */
	/** @type {Row[]} */
	const unaddressable = []
	/** @type {Row[]} */
	const unread = []
	/** @type {Row[]} */
	const weak = []
	for (const [qname, entry] of totals) {
		if (consumed.has(qname)) continue
		const row = { qname, occurrences: entry.count, fixtures: [...entry.fixtures].sort() }
		const braced = /^\{([^}]*)\}(.+)$/.exec(qname)
		if (braced && !consumedUris.has(braced[1] ?? '')) {
			// Namespace absent from src/read/ entirely — no accessor can name it.
			unaddressable.push(row)
			continue
		}
		// A bare-localName mention means a namespace-comparison accessor may well read it.
		const local = braced ? braced[2] : qname.slice(qname.indexOf(':') + 1)
		;(consumedLocals.has(local) ? weak : unread).push(row)
	}
	/**
	 * @param {Row} a
	 * @param {Row} b
	 */
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

	/**
	 * @param {string} title
	 * @param {Row[]} rows
	 * @param {string} [note]
	 */
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
