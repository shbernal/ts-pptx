#!/usr/bin/env node
/**
 * OOXML literal gate — a schema URI or content type outside `src/ooxml/` has to say why it is there.
 *
 * `src/ooxml/` exists because a private copy of a format fact on each side of the library gives no
 * compile-time signal when the two copies diverge (`docs/architecture.md`, Boundaries). Moving the
 * copies there does not keep them there: nothing stops the next emitter from spelling a URI inline
 * again, and a writer's copy and a reader's copy drifting apart is a silent round-trip failure.
 *
 * So this fails on any `http://schemas.` or `application/vnd.` literal in `src/` outside
 * `src/ooxml/`, unless an allowlist entry names that literal in that file and gives the reason it
 * stays. The documented reasons:
 *
 * - A fact exactly one module reads or writes stays declared beside that module, as the header of
 *   `src/ooxml/rel-types.ts` explains. A writer and a reader of the same part are two modules.
 * - XML captured verbatim from Office needs no entry at all: a declaration marked `@raw-xml-asset`
 *   is exempt here for the reason the raw-XML ratchet exempts it.
 *
 * The scan is `raw-xml-ratchet.mjs`'s TypeScript AST walk run with a different pattern, so a doc
 * comment is not a finding and neither is an argument to a message sink (`warn`, `note`, an error
 * constructor): prose that names a content type is not a second copy of it.
 *
 * An entry whose literal is gone from its file fails too, so the allowlist cannot outlive the
 * code it describes.
 *
 *   node scripts/ooxml-literal-gate.mjs          # check (exit 1 on an unlisted, stale or unexplained entry)
 *   node scripts/ooxml-literal-gate.mjs --list   # every literal outside src/ooxml/, with line numbers
 */

import fs from 'node:fs'
import path from 'node:path'
import { ROOT, isMain, parseCli, runCli } from './script-utils.mjs'
import { scanSource, tsFilesUnder } from './raw-xml-ratchet.mjs'

const SRC = path.join(ROOT, 'src')
const ALLOWLIST = path.join(ROOT, 'scripts', 'ooxml-literal-allowlist.json')
/** Where these literals belong. */
const HOME_DIR = 'src/ooxml/'

/** A schema URI or a vendor content type, to the end of its token. */
export const OOXML_LITERAL = /(?:http:\/\/schemas\.|application\/vnd\.)[^\s'"`<>]*/g

/**
 * @typedef {{ file: string, line: number, literal: string }} Finding
 * @typedef {{ file: string, literal: string, reason: string }} AllowlistEntry
 */

/**
 * Every OOXML literal in `src/` outside `src/ooxml/`.
 * @returns {Finding[]}
 */
export function collectLiterals() {
	/** @type {Finding[]} */
	const found = []
	for (const file of tsFilesUnder(SRC)) {
		const rel = path.relative(ROOT, file).split(path.sep).join('/')
		if (rel.startsWith(HOME_DIR)) continue
		for (const { line, text } of scanSource(fs.readFileSync(file, 'utf8'), file, OOXML_LITERAL)) {
			found.push({ file: rel, line, literal: text })
		}
	}
	return found
}

/**
 * Compare what the scan found with the allowlist.
 *
 * Keyed by file and literal together, so a literal allowlisted in one file is still a finding in
 * another, and a new URI in an allowlisted file is still a finding there.
 * @param {Finding[]} found
 * @param {AllowlistEntry[]} allowlist
 * @returns {{ unlisted: Finding[], stale: AllowlistEntry[], reasonless: AllowlistEntry[] }}
 */
export function compareToAllowlist(found, allowlist) {
	/** @param {{ file: string, literal: string }} entry */
	const key = (entry) => `${entry.file}\0${entry.literal}`
	const allowed = new Set(allowlist.map(key))
	const present = new Set(found.map(key))
	return {
		unlisted: found.filter((finding) => !allowed.has(key(finding))),
		stale: allowlist.filter((entry) => !present.has(key(entry))),
		reasonless: allowlist.filter((entry) => typeof entry.reason !== 'string' || entry.reason.trim() === ''),
	}
}

// ---------------------------------------------------------------- CLI

const USAGE = `OOXML literal gate — a schema URI or content type outside src/ooxml/ has to say why it is there.

  node scripts/ooxml-literal-gate.mjs          check (exit 1 on an unlisted, stale or unexplained entry)
  node scripts/ooxml-literal-gate.mjs --list   every literal outside src/ooxml/, with line numbers

Options:
  --list      print every literal found rather than the allowlist comparison
  -h, --help  show this message`

/** @param {string[]} argv @returns {number} process exit code */
export function main(argv) {
	const { values } = parseCli(argv, {
		usage: USAGE,
		options: { list: { type: 'boolean', default: false } },
	})

	const found = collectLiterals()
	if (values.list) {
		for (const { file, line, literal } of found) console.log(`${file}:${line}  ${literal}`)
		return 0
	}

	/** @type {AllowlistEntry[]} */
	const allowlist = JSON.parse(fs.readFileSync(ALLOWLIST, 'utf8'))
	const relAllowlist = path.relative(ROOT, ALLOWLIST).split(path.sep).join('/')
	const { unlisted, stale, reasonless } = compareToAllowlist(found, allowlist)

	if (unlisted.length || stale.length || reasonless.length) {
		console.error('ooxml literal gate FAILED\n')
		if (unlisted.length) {
			console.error(`Literals outside ${HOME_DIR} with no allowlist entry:`)
			for (const { file, line, literal } of unlisted) console.error(`  ${file}:${line}  ${literal}`)
			console.error(
				`\nImport it from ${HOME_DIR}rel-types.ts or ${HOME_DIR}namespaces.ts, adding it there if it is new.`
			)
			console.error(`If exactly one module reads or writes it, add { file, literal, reason } to ${relAllowlist}.\n`)
		}
		if (stale.length) {
			console.error(`Allowlist entries whose literal is no longer in that file — drop them from ${relAllowlist}:`)
			for (const { file, literal } of stale) console.error(`  ${file}  ${literal}`)
			console.error('')
		}
		if (reasonless.length) {
			console.error('Allowlist entries with no reason:')
			for (const { file, literal } of reasonless) console.error(`  ${file}  ${literal}`)
		}
		return 1
	}

	console.log(`ooxml literal gate: ok (${found.length} allowlisted literal(s) outside ${HOME_DIR})`)
	return 0
}

if (isMain(import.meta.url)) await runCli(() => main(process.argv.slice(2)))
