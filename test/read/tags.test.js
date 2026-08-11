// Read the programmatic tags (`p:custDataLst/p:tags` → `ppt/tags/tagN.xml`) off a
// genuine PowerPoint-authored fixture. tags.pptx was authored via desktop
// PowerPoint COM (`Presentation.Tags.Add` + `Slide.Tags.Add`); the oracle is the
// exact name/val pairs passed to `.Add`. Slide 2 is left tag-free to pin the
// empty-owner → [] case. There is no writer for tags, so a fixture is the only
// source (unlike the document-properties round-trip).

import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'

import { assert, assertEqual } from '../helpers.js'
import { openFixture } from './corpus.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

/** Render a tag list as "name=val,name=val" for order-preserving equality. */
function flatten(tags) {
	return tags.map((t) => `${t.name}=${t.val}`).join(',')
}

describe('Presentation.tags / Slide.tags', () => {
	test('reads the deck-level tags in authored order', async () => {
		const pres = await openFixture('tags')
		assertEqual(flatten(pres.tags), 'REVIEWER=Ada Lovelace,STAGE=draft', 'deck tags')
	})

	test('reads a slide’s own tags, resolved through its rel', async () => {
		const pres = await openFixture('tags')
		const [slide1] = pres.slides
		assertEqual(flatten(slide1.tags), 'REGION=EMEA,PRIORITY=high', 'slide 1 tags')
	})

	test('a tag-free slide reads as []', async () => {
		const pres = await openFixture('tags')
		const slide2 = pres.slides[1]
		assert(Array.isArray(slide2.tags), 'tags is an array')
		assertEqual(slide2.tags.length, 0, 'slide 2 has no tags')
	})

	test('each tag is a { name, val } string pair', async () => {
		const pres = await openFixture('tags')
		for (const t of [...pres.tags, ...pres.slides[0].tags]) {
			assert(typeof t.name === 'string' && typeof t.val === 'string', `tag is name/val strings: ${JSON.stringify(t)}`)
		}
	})

	test('a deck with no tag parts → [] at every level', async () => {
		// empty.pptx carries no custDataLst on the presentation or its slide.
		const pres = await openFixture('empty')
		assertEqual(pres.tags.length, 0, 'no deck tags')
		assertEqual(pres.slides[0].tags.length, 0, 'no slide tags')
	})
})
