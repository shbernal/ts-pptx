// Fixture-driven edges for src/read/api/chrome.ts that an authored deck cannot show.
//
// chrome-read.test.js measures the write→read round-trip of the chrome model, but
// the writer authors a deliberately plain master: no `p:cSld/p:bg` of its own, and
// every placeholder carrying explicit geometry. Real PowerPoint decks are the other
// way round — the master owns the background and layout placeholders routinely omit
// `a:xfrm` to inherit the master's. So `SlideMaster.background` and the
// inherited-geometry reads are asserted here against genuine PowerPoint output.
//
// Fixtures used: mixed.pptx (a master with an explicit `p:bgPr` solid fill),
// theme-colors.pptx (the Ion theme — a master `p:bgRef`, plus layouts that define no
// background of their own).

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { authorRead } from './authored.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

async function open(name, ext = 'pptx') {
	return Presentation.load(await readFile(path.join(__dirname, 'fixtures', `${name}.${ext}`)))
}

describe('SlideMaster.background — the master tier of the chain', () => {
	test("an explicit p:bgPr on the master reads as the master's own solid fill", async () => {
		const bg = (await open('mixed')).masters()[0].background
		assert(bg?.type === 'solid', "mixed.pptx's master authors a solid p:bgPr")
		assertEqual(bg.source, 'master', 'the source records the master tier, not slide/layout')
		assertEqual(bg.color?.effectiveHex, 'FFFFFF', 'the fill colour resolves')
	})

	test('a p:bgRef on the master keeps its idx and resolves through the theme', async () => {
		const bg = (await open('theme-colors')).masters()[0].background
		assert(bg?.type === 'themeRef', 'the Ion master authors a theme-indexed background')
		assertEqual(bg.source, 'master', 'sourced from the master')
		assertEqual(bg.idx, 1003, 'the raw bgRef index')
		assertEqual(bg.color?.effectiveHex, '1E5155', "the bgRef's own schemeClr resolves through the Ion palette")
	})

	test('a master defining no background of its own reads null, not a none-fill', async () => {
		// The writer authors the background onto the layout, leaving the master's
		// `p:cSld` without a `p:bg` — so the getter must distinguish "defines none"
		// (null) from "explicitly transparent" (`type: 'none'`).
		const { presentation } = await authorRead((pres) => {
			pres.defineSlideMaster({ title: 'BRANDED', background: { color: 'F1F2F3' } })
			pres.addSlide({ masterTitle: 'BRANDED' })
		})
		const master = presentation.masters()[0]
		assertEqual(master.background, null, 'no p:bg on the master → null')
		// …and the layout that does define one still reads it, so the null above is a
		// real absence rather than a broken lookup.
		const branded = master.layouts.find((l) => l.name === 'BRANDED')
		assertEqual(branded?.background?.type, 'solid', 'the layout tier still resolves')
	})
})

describe('SlideLayout — inherited tiers', () => {
	test('a layout defining no background of its own reads null', async () => {
		const layout = (await open('theme-colors')).masters()[0].layouts[0]
		assertEqual(layout.name, 'Title Slide', 'the first Ion layout')
		assertEqual(layout.background, null, 'PowerPoint layouts usually inherit the master background')
	})

	test('a layout resolves its theme through its master', async () => {
		const layout = (await open('theme-colors')).masters()[0].layouts[0]
		const theme = layout.theme
		assert(theme, 'layout.theme walks layout → master → theme')
		assertEqual(theme.name, 'Ion', 'the fixture theme')
		assertEqual(theme.partName, layout.master?.theme?.partName, 'the shorthand reaches the same part as the long walk')
	})

	test("the import-only layout @type reads PowerPoint's value", async () => {
		// chrome-read.test.js pins this as null on an authored deck (the writer emits
		// none); the imported side is where a real value shows up.
		const layout = (await open('theme-colors')).masters()[0].layouts[0]
		assertEqual(layout.type, 'title', 'an imported layout carries its @type')
	})
})

describe('Placeholder geometry — inherited rather than own', () => {
	test('a layout placeholder with no own a:xfrm reads null geometry on all four axes', async () => {
		const layouts = (await open('theme-colors')).masters()[0].layouts
		const inherited = layouts.flatMap((l) => l.placeholders).find((ph) => ph.type === 'dt')
		assert(inherited, 'the Ion layouts carry a date placeholder')
		// It still identifies itself — only the geometry is absent, because PowerPoint
		// lets it inherit the master placeholder's box.
		assertEqual(inherited.idx, '10', 'the p:ph idx still reads')
		assert(inherited.name.length > 0, 'the shape name still reads')
		assert(typeof inherited.id === 'number', 'the drawing id still reads')
		assertEqual(inherited.left, null, 'no own a:xfrm → null left')
		assertEqual(inherited.top, null, 'null top')
		assertEqual(inherited.width, null, 'null width')
		assertEqual(inherited.height, null, 'null height')
	})
})
