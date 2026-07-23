// Self-test for the shared write→read fidelity harness (test/read/authored.js).
//
// Proves the author→read→validate plumbing end-to-end against the real dist
// entries, so the read-side-expansion batches build a feature on top of a
// known-good harness instead of debugging the harness and the feature at once.
// Assertions here mirror constructs already proven elsewhere (a rect autoShape
// with a solid fill, per style-accessors.test.js) so the self-test can only fail
// if the harness plumbing itself breaks.

import { describe, test } from 'vitest'
import { authorRead, firstShape, schemaErrors, validatorInstalled } from './authored.js'
import { assert, assertEqual } from '../helpers.js'

describe('write→read harness (authored.js)', () => {
	test('authorRead: a written rect round-trips into the deep read model', async () => {
		const { presentation, buf } = await authorRead((pres) => {
			pres.addSlide().addShape(pres.ShapeType.rect, { x: 1, y: 1, w: 3, h: 1, fill: { color: 'CCCCCC' } })
		})
		assert(buf.length > 0, 'stream() produced bytes')
		assertEqual(presentation.slides.length, 1, 'one authored slide is read back')

		const rect = firstShape(presentation, (s) => s.shapeType === 'autoShape' && s.presetGeometry === 'rect')
		assert(rect, 'the authored rect is located in the read model')
		assertEqual(rect.fillColor, 'CCCCCC', 'the authored solid fill reads back')
	})

	test('firstShape returns null when nothing matches', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addText('hi', { x: 1, y: 1, w: 3, h: 1 })
		})
		assertEqual(
			firstShape(presentation, (s) => s.shapeType === 'graphicFrame'),
			null,
			'a text-only deck has no graphicFrame'
		)
	})

	test.skipIf(!validatorInstalled)('schemaErrors: authored bytes are schema-valid', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addShape(pres.ShapeType.rect, { x: 1, y: 1, w: 3, h: 1, fill: { color: 'CCCCCC' } })
		})
		const errors = await schemaErrors(buf)
		assertEqual(errors.length, 0, `authored deck validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
