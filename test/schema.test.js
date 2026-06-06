'use strict'

// Schema-validation fixtures. Each case builds a representative `.pptx`
// and asserts the OpenXmlValidator (via OOXMLValidatorCLI) reports no
// errors.
//
// Fixtures are intentionally small and orthogonal — they exercise one
// API surface each — so when an error appears we can localise it.
//
// Run with: npm run test:schema

const { build, assert } = require('./helpers')
const { validateBuf } = require('./validator')

async function expectNoSchemaErrors (buf, label) {
	const errors = await validateBuf(buf)
	if (errors.length === 0) return
	const summary = errors
		.slice(0, 5)
		.map(e => `  - [${e.ErrorType}] ${e.Description} (path: ${(e.Path && e.Path.PartUri) || '?'})`)
		.join('\n')
	const more = errors.length > 5 ? `\n  ...(${errors.length - 5} more)` : ''
	assert(
		false,
		`${label}: ${errors.length} schema error(s):\n${summary}${more}`
	)
}

module.exports = [
	{
		name: 'empty deck (one slide, no content)',
		fn: async () => {
			const { buf } = await build(p => { p.addSlide() })
			await expectNoSchemaErrors(buf, 'empty-deck')
		}
	},
	{
		name: 'single text box',
		fn: async () => {
			const { buf } = await build(p => {
				p.addSlide().addText('hello', { x: 1, y: 1, w: 4, h: 0.5 })
			})
			await expectNoSchemaErrors(buf, 'single-text')
		}
	},
	{
		name: 'single rectangle shape',
		fn: async () => {
			const { buf } = await build(p => {
				const s = p.addSlide()
				s.addShape(p.shapes.RECTANGLE, { x: 1, y: 1, w: 2, h: 1, fill: { color: 'FF0000' } })
			})
			await expectNoSchemaErrors(buf, 'single-shape')
		}
	},
	{
		name: 'shape with shadow',
		fn: async () => {
			const { buf } = await build(p => {
				const s = p.addSlide()
				s.addShape(p.shapes.RECTANGLE, {
					x: 1, y: 1, w: 4, h: 1,
					fill: { color: '00B0B9' },
					shadow: { type: 'outer', blur: 6, offset: 2, color: '000000', opacity: 0.15 }
				})
			})
			await expectNoSchemaErrors(buf, 'shape-with-shadow')
		}
	},
	{
		name: 'solid-color slide background',
		fn: async () => {
			const { buf } = await build(p => {
				const s = p.addSlide()
				s.background = { color: '0088CC' }
				s.addText('hi', { x: 1, y: 1 })
			})
			await expectNoSchemaErrors(buf, 'solid-bg')
		}
	},
	{
		name: 'bullet text',
		fn: async () => {
			const { buf } = await build(p => {
				p.addSlide().addText('item', { x: 1, y: 1, w: 4, h: 0.5, bullet: true })
			})
			await expectNoSchemaErrors(buf, 'bullet-text')
		}
	},
	{
		name: 'simple table',
		fn: async () => {
			const { buf } = await build(p => {
				p.addSlide().addTable(
					[
						[ { text: 'A1' }, { text: 'B1' } ],
						[ { text: 'A2' }, { text: 'B2' } ]
					],
					{ x: 1, y: 1, w: 4 }
				)
			})
			await expectNoSchemaErrors(buf, 'simple-table')
		}
	},
	{
		name: 'embedded PNG image',
		fn: async () => {
			const b64 =
				'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { buf } = await build(p => {
				p.addSlide().addImage({ data: 'image/png;base64,' + b64, x: 1, y: 1, w: 1, h: 1 })
			})
			await expectNoSchemaErrors(buf, 'embedded-png')
		}
	}
]
