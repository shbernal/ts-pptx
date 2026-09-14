// Importing again after a slide is removed.
//
// `removeSlide` frees partnames, and `reservePartNameLike` gives them to the next import. The deck's
// import memos hold destination partnames, so an entry that outlived its part answered for whatever
// came next: the copy registry called a deleted slide or image already copied, and the rescale memo
// called a new slide already scaled.

import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual, captureDiagnostics } from '../helpers.js'
import { openFixture, readFixture } from './corpus.js'
import { assertNoDanglingRels } from './opc.js'

/** The left edge of every picture on `slide`, in EMU. */
function pictureLefts(slide) {
	return slide.shapes.filter((shape) => shape.shapeType === 'picture').map((shape) => shape.left)
}

describe('an import after removeSlide', () => {
	test('re-importing a page onto the name its removal freed copies it again', async () => {
		const source = await openFixture('image')
		const deck = await openFixture('image')
		const first = deck.importSlide(source, 0)
		const shapeCount = first.shapes.length
		deck.removeSlide(deck.slides.length - 1)

		const again = deck.importSlide(source, 0)
		assertEqual(again.partName, first.partName, 'the import is given the freed name')
		assertEqual(again.shapes.length, shapeCount, 'and carries the page')
		assertNoDanglingRels((await Presentation.load(await deck.save())).opc)
	})

	test('a later import does not link to media the removal pruned', async () => {
		const source = await openFixture('image')
		const deck = await openFixture('autofit-cjk-wrap')
		const media = () => [...deck.opc.parts.keys()].filter((partName) => partName.startsWith('/ppt/media/'))
		assertEqual(media().length, 0, 'the destination starts with no media')

		deck.importSlide(source, 0)
		assert(media().length > 0, 'the import brings media')
		deck.removeSlide(deck.slides.length - 1)
		assertEqual(media().length, 0, 'the removal prunes it')

		deck.importSlide(source, 1)
		assert(media().length > 0, 'the next import copies its media again')
		assertNoDanglingRels(deck.opc)
		assertNoDanglingRels((await Presentation.load(await deck.save())).opc)
	})

	test('a rescaled import given a freed name is rescaled', async () => {
		// Rescaling warns, and the warnings are not what this measures.
		const { result } = await captureDiagnostics(async () => {
			const deck = await openFixture('mixed')
			// Two loads of one file are two sources with a registry each, so only the rescale memo
			// could skip the second.
			const first = deck.importSlide(await Presentation.load(await readFixture('image')), 0, { rescale: 'fit' })
			const firstLefts = pictureLefts(first)
			deck.removeSlide(deck.slides.length - 1)
			const second = deck.importSlide(await Presentation.load(await readFixture('image')), 0, { rescale: 'fit' })
			return { deck, first: first.partName, firstLefts, second: second.partName, secondLefts: pictureLefts(second) }
		})
		assertEqual(result.second, result.first, 'the import is given the freed name')
		assert(result.firstLefts.length > 0, 'the page has a picture to measure')
		assertEqual(result.secondLefts.join(), result.firstLefts.join(), 'and is scaled as the first was')
		assertNoDanglingRels(result.deck.opc)
	})
})
