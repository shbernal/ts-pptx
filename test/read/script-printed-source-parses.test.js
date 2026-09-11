// A printed script parses, whatever the deck's names and text hold.
//
// Deck text reaches the printed source two ways. Inside a string literal `printString` escapes
// it. Inside a comment (the fidelity banner, the `//` line above each slide and each layout
// handle) there is no literal around it, so a name that ends the comment early turns the rest of
// it into code: a `*/` closes the banner, and U+2028, which `JSON.stringify` leaves raw, is a
// line terminator that closes a `//` line. A lone surrogate has no UTF-8 encoding at all.
//
// So the deck below gives a slide, a shape and two layouts a name holding both, puts a lone
// surrogate in a run, and parses what each tier prints.

import { describe, test } from 'vitest'
import JSZip from 'jszip'
import ts from 'typescript-6'
import { Presentation } from '../../dist/read.js'
import { printScript, printStandaloneScript, readModelToIr } from '../../dist/script.js'
import { assert, assertEqual } from '../helpers.js'
import { authorRead } from './authored.js'

/** A name that ends a block comment and, separately, a line comment. */
const HOSTILE = `Band */ one${String.fromCharCode(0x2028)}two`
/** Run text the IR walk swaps for {@link LONE}; the writer cannot store a lone surrogate itself. */
const RUN = 'RUN-MARKER'
const LONE = 'half \uD800 pair'

/**
 * A deck whose slide, shape and both layouts are named {@link HOSTILE}. The layouts share the
 * name, so the template-anchored tier binds by position and prints the name beside the handle.
 */
async function hostileDeck() {
	const { buf } = await authorRead((pres) => {
		pres.defineSlideMaster({ title: 'LayoutA' })
		pres.defineSlideMaster({ title: 'LayoutB' })
		pres.addSlide({ masterTitle: 'LayoutA' }).addText(RUN, { x: 1, y: 1, w: 4, h: 1, objectName: HOSTILE })
	})
	const zip = await JSZip.loadAsync(buf)
	let layoutRenames = 0
	let slideRenames = 0
	for (const name of Object.keys(zip.files)) {
		const isLayout = /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name)
		const isSlide = /^ppt\/slides\/slide\d+\.xml$/.test(name)
		if (!isLayout && !isSlide) continue
		let xml = await zip.file(name).async('string')
		if (isLayout) {
			xml = xml.replace(/<p:cSld name="Layout[AB]">/, () => {
				layoutRenames++
				return `<p:cSld name="${HOSTILE}">`
			})
		} else {
			xml = xml.replace(/<p:cSld\b[^>]*>/, () => {
				slideRenames++
				return `<p:cSld name="${HOSTILE}">`
			})
		}
		zip.file(name, xml)
	}
	assertEqual(layoutRenames, 2, 'both layouts take the same name')
	assertEqual(slideRenames, 1, 'the slide is named')
	return Presentation.load(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
}

/** The deck's IR with every {@link RUN} string replaced by {@link LONE}. */
async function hostileIr() {
	const ir = readModelToIr(await hostileDeck())
	let planted = 0
	/** @param {any} value */
	const plant = (value) => {
		if (!value || typeof value !== 'object') return
		for (const key of Object.keys(value)) {
			if (value[key] === RUN) {
				value[key] = LONE
				planted++
			} else plant(value[key])
		}
	}
	plant(ir.slides)
	assert(planted > 0, 'the run text is in the IR to replace')
	return ir
}

/** The printed module's parse errors, as message text. */
function parseErrors(code) {
	const source = ts.createSourceFile('script.ts', code, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS)
	// `parseDiagnostics` is not in the public typings, but it is exactly the syntax-only set wanted:
	// a program's diagnostics would also type-check against a package this test does not install.
	const diagnostics = /** @type {{ parseDiagnostics: readonly import('typescript-6').Diagnostic[] }} */ (
		/** @type {unknown} */ (source)
	).parseDiagnostics
	return diagnostics.map((diagnostic) => ts.flattenDiagnosticMessageText(diagnostic.messageText, '\n'))
}

/** No unpaired UTF-16 surrogate anywhere in `text`. */
const LONE_SURROGATE = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/

describe('a printed script parses whatever the deck names hold', () => {
	test('the template-anchored tier', async () => {
		const { code } = printScript(await hostileIr())
		// The three comment sites this tier puts a name in, so a parse that passes is not one that
		// simply never printed them.
		assert(code.includes('slide.name'), 'the banner quotes the slide name')
		assert(/deck\.layouts\(\)\[\d+\] \/\/ /.test(code), 'the ambiguous layout is bound by position, with its name')
		assert(code.includes('// Slide 1 — layout '), 'the slide comment names its layout')
		assertEqual(parseErrors(code).join('; '), '', 'the module parses')
		assert(!LONE_SURROGATE.test(code), 'no lone surrogate is written raw')
		assert(code.includes('\\ud800'), 'it is escaped instead')
	})

	test('the standalone tier', async () => {
		const { code } = printStandaloneScript(await hostileIr())
		assert(code.includes('slide.name'), 'the banner quotes the slide name')
		assert(code.includes('// Slide 1 — master '), 'the slide comment names its master')
		assertEqual(parseErrors(code).join('; '), '', 'the module parses')
		assert(!LONE_SURROGATE.test(code), 'no lone surrogate is written raw')
		assert(code.includes('\\ud800'), 'it is escaped instead')
	})
})
