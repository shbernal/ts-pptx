// Read-model coverage for placeholder text-property INHERITANCE
// (src/read/api/theme-context.ts `resolveInheritedRun{Size,FontFace,Bold}`,
// surfaced by Run.resolved{SizePt,FontFace,Bold} in src/read/api/text.ts).
//
// placeholder-inherit.pptx is a minimal, locally PowerPoint-authored deck: one
// slide with a title placeholder and a body placeholder, each holding plain text
// typed with NO run-level formatting. PowerPoint writes those runs as bare
// `<a:rPr lang="en-US"/>` — no `@sz`, no `a:latin`, no `@b` — so the effective
// size / typeface / bold must be resolved by walking the placeholder → master
// `p:txStyles` chain and (for the face) resolving the `+mj-lt`/`+mn-lt` theme
// token through the theme `fontScheme`. The theme is baked into the committed
// bytes, so the expected values are machine-stable (this deck ships Aptos).
//
// This is the construct the existing read fixtures did not exercise: they carry
// explicit run formatting, so the inherited-value paths (notably the font-face
// resolution and the "chain defines no bold" fallthrough) stayed uncovered.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

async function open(name) {
	return Presentation.load(await readFile(path.join(__dirname, 'fixtures', `${name}.pptx`)))
}

/** First run of the first paragraph of a shape's text frame. */
function firstRun(shape) {
	const frame = shape.textFrame
	assert(frame, `expected a text frame on shape "${shape?.name}"`)
	const para = frame.paragraphs[0]
	assert(para, `expected a paragraph on shape "${shape?.name}"`)
	const run = para.runs[0]
	assert(run, `expected a run on shape "${shape?.name}"`)
	return run
}

describe('read: placeholder text-property inheritance', () => {
	test('a bare title run inherits size + theme (major) face from the master txStyles', async () => {
		const slide = (await open('placeholder-inherit')).slides[0]
		const title = slide.shapes.find((s) => /^Title/.test(s.name ?? ''))
		assert(title, 'expected a title placeholder shape')
		const run = firstRun(title)

		// The run itself sets nothing — proving the resolved values below come from inheritance.
		assertEqual(run.fontName, null, 'title run sets no own typeface')
		assertEqual(run.fontSizePt, null, 'title run sets no own size')
		assertEqual(run.bold, null, 'title run sets no own bold')

		// titleStyle/lvl1 defRPr: sz=4400, latin=+mj-lt → theme major font (Aptos Display).
		assertEqual(run.resolvedSizePt, 44, 'title inherits 44pt from titleStyle')
		assertEqual(run.resolvedFontFace, 'Aptos Display', 'title inherits the +mj-lt theme major font')
		// titleStyle defRPr sets no @b, and nothing above it does → the bold chain resolves to null.
		assertEqual(run.resolvedBold, null, 'title inherits no explicit bold from the chain')
	})

	test('a bare body run inherits size + theme (minor) face from the master txStyles', async () => {
		const slide = (await open('placeholder-inherit')).slides[0]
		// The body placeholder is the non-title text placeholder.
		const body = slide.shapes.find((s) => !/^Title/.test(s.name ?? '') && s.textFrame)
		assert(body, 'expected a body placeholder shape')
		const run = firstRun(body)

		assertEqual(run.fontName, null, 'body run sets no own typeface')
		assertEqual(run.fontSizePt, null, 'body run sets no own size')

		// bodyStyle/lvl1 defRPr: sz=2800, latin=+mn-lt → theme minor font (Aptos).
		assertEqual(run.resolvedSizePt, 28, 'body inherits 28pt from bodyStyle')
		assertEqual(run.resolvedFontFace, 'Aptos', 'body inherits the +mn-lt theme minor font')
		assertEqual(run.resolvedBold, null, 'body inherits no explicit bold from the chain')
	})
})
