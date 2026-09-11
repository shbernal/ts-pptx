// An import's dry run is its copy, run as a plan: the traversal that copies a page, with every
// write sent to a `CopyPlan` instead of the deck. What the plan promises is that it agrees with the
// copy — refusing a source the copy would fail on halfway, and accepting one the copy completes.
// The most direct evidence of that is the parts: across the corpus, the parts a plan lists are
// exactly the parts the import then adds, name for name and in the same order.
//
// Imports `src/` rather than `dist/`, because the plan is internal and the question is about it.

import { describe, test } from 'vitest'
import { Presentation, setDiagnosticHandler } from '../../src/read.ts'
import { CopyPlan } from '../../src/read/api/ops/part-copy.ts'
import { copyBatch, planSlideImport } from '../../src/read/api/presentation-imports.ts'
import { assertEqual } from '../helpers.js'
import { fixtureNames, readFixture } from './corpus.js'

// Rescaling a fixture onto the destination's canvas warns, and warnings are not what this measures.
setDiagnosticHandler(() => {})

const MODES = /** @type {const} */ ([
	{ theme: 'copy', importNotes: true, rescale: 'fit' },
	{ theme: 'preserve', importNotes: true, carryMasterGraphics: true, rescale: 'fit' },
	{ theme: 'restyle', importNotes: true, carryMasterGraphics: true, rescale: 'fit' },
])

/** Two destinations: another deck, and a template of the source itself, where reuse applies. */
async function destinations(bytes) {
	return [await Presentation.load(await readFixture('mixed')), await Presentation.fromTemplate(bytes)]
}

/** Run `apply` and return the partnames it added to `deck`, in the order it added them. */
function addedBy(deck, apply) {
	const before = new Set(deck.opc.parts.keys())
	apply()
	return [...deck.opc.parts.keys()].filter((name) => !before.has(name))
}

/** The partnames a plan lists, as one comparable string. */
function listed(plan) {
	return JSON.stringify(plan.parts.map((part) => part.partName))
}

describe('an import plan lists exactly the parts the import adds', () => {
	test.for(fixtureNames)('importSlide from %s, in every theme mode', async (name) => {
		const bytes = await readFixture(name)
		for (const options of MODES) {
			for (const dest of await destinations(bytes)) {
				const source = await Presentation.load(bytes)
				// Every page, then the first again: a repeat is where the copy registry and the page's
				// ownership scope decide what is copied and what is shared.
				const indexes = source.slides.length > 0 ? [...source.slides.keys(), 0] : []
				for (const index of indexes) {
					const partCount = dest.opc.parts.size
					const plan = planSlideImport(dest, source, source.slides[index], options)
					assertEqual(dest.opc.parts.size, partCount, `${name}: planning added no part`)
					const added = addedBy(dest, () => dest.importSlide(source, index, options))
					assertEqual(listed(plan), JSON.stringify(added), `${name}, ${options.theme}, page ${index}`)
				}
			}
		}
	})

	test.for(fixtureNames)('importSlides from %s, every page and the first twice', async (name) => {
		const bytes = await readFixture(name)
		for (const dest of await destinations(bytes)) {
			const source = await Presentation.load(bytes)
			if (source.slides.length === 0) continue
			const base = dest.slides.length
			const requests = [...source.slides.keys(), 0].map((sourceIndex, j) => ({
				source,
				sourceIndex,
				outputIndex: base + j,
				importNotes: j % 2 === 0,
				rescale: /** @type {const} */ ('fit'),
			}))
			const plan = new CopyPlan(dest, 'importSlides')
			copyBatch(
				dest,
				requests.map((request) => ({ ...request, sourceSlide: source.slides[request.sourceIndex] })),
				plan
			)
			const added = addedBy(dest, () => dest.importSlides(requests))
			assertEqual(listed(plan), JSON.stringify(added), `${name}, batch`)
		}
	})
})
