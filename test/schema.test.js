// Schema-validation fixtures. Each case builds a representative `.pptx`
// and asserts the OpenXmlValidator (via OOXMLValidatorCLI) reports no
// errors.
//
// Fixtures are intentionally small and orthogonal — they exercise one
// API surface each — so when an error appears we can localise it.
//
// Run with: pnpm run test:schema

import { build, assert } from './helpers.js'
import { validateBuf } from './validator.js'

async function expectNoSchemaErrors(buf, label) {
	const errors = await validateBuf(buf)
	if (errors.length === 0) return
	const summary = errors
		.slice(0, 5)
		.map((e) => `  - [${e.ErrorType}] ${e.Description} (path: ${(e.Path && e.Path.PartUri) || '?'})`)
		.join('\n')
	const more = errors.length > 5 ? `\n  ...(${errors.length - 5} more)` : ''
	assert(false, `${label}: ${errors.length} schema error(s):\n${summary}${more}`)
}

export default [
	{
		name: 'empty deck (one slide, no content)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide()
			})
			await expectNoSchemaErrors(buf, 'empty-deck')
		},
	},
	{
		name: 'single text box',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText('hello', { x: 1, y: 1, w: 4, h: 0.5 })
			})
			await expectNoSchemaErrors(buf, 'single-text')
		},
	},
	{
		name: 'company metadata with XML entities',
		fn: async () => {
			const { buf } = await build((p) => {
				p.company = 'A & B <C>'
				p.addSlide().addText('hello', { x: 1, y: 1, w: 4, h: 0.5 })
			})
			await expectNoSchemaErrors(buf, 'company-metadata-xml-entities')
		},
	},
	{
		// Asserts the body-property serialization stays schema-valid. Note: this
		// proves the XML is well-formed, not that PowerPoint/LibreOffice renders
		// a particular layout (see UPSTREAMING_CANDIDATES.md "Text Box Behavior").
		name: 'text box with margins',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText('hello', { x: 1, y: 1, w: 4, h: 1, margin: [10, 5, 10, 5] })
			})
			await expectNoSchemaErrors(buf, 'text-margins')
		},
	},
	{
		// Asserts numCol/spcCol body-property serialization stays schema-valid
		// (upstream-issue-1320). numCol is bounded 1-16 by ECMA-376
		// ST_TextColumnCount; spcCol is EMU. Rendering layout is not asserted here.
		name: 'text box with multiple columns',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText('column flow text', { x: 1, y: 1, w: 6, h: 2, columns: 2, columnSpacing: 12 })
			})
			await expectNoSchemaErrors(buf, 'text-columns')
		},
	},
	{
		name: 'text box with vertical alignment',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addText('top', { x: 1, y: 1, w: 4, h: 1, valign: 'top' })
				s.addText('middle', { x: 1, y: 2, w: 4, h: 1, valign: 'middle' })
				s.addText('bottom', { x: 1, y: 3, w: 4, h: 1, valign: 'bottom' })
			})
			await expectNoSchemaErrors(buf, 'text-valign')
		},
	},
	{
		name: 'text box with mixed bold/color runs',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText(
					[
						{ text: 'bold red ', options: { bold: true, color: 'FF0000' } },
						{ text: 'plain ', options: {} },
						{ text: 'blue', options: { color: '0000FF' } },
					],
					{ x: 1, y: 1, w: 4, h: 1 }
				)
			})
			await expectNoSchemaErrors(buf, 'text-mixed-runs')
		},
	},
	{
		// Serialization-contract fixture: breakLine: false on a CRLF-containing run must
		// produce valid OOXML (upstream-issue-1138). The rendering result is layout-dependent
		// and not asserted here — see UPSTREAMING_CANDIDATES.md "Text Box Behavior".
		name: 'rich text with breakLine: false on CRLF-containing run',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText(
					[
						{ text: 'first\nsecond', options: { breakLine: false } },
						{ text: ' tail', options: {} },
					],
					{ x: 1, y: 1, w: 4, h: 1 }
				)
			})
			await expectNoSchemaErrors(buf, 'rich-text-breakline-false')
		},
	},
	{
		name: 'single rectangle shape',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.RECTANGLE, { x: 1, y: 1, w: 2, h: 1, fill: { color: 'FF0000' } })
			})
			await expectNoSchemaErrors(buf, 'single-shape')
		},
	},
	{
		// Object lock flags (upstream-issue-438): spLocks on a shape, picLocks on an
		// image, graphicFrameLocks on a table. Asserts each locking element + its
		// element-type-specific attributes serialize to schema-valid OOXML.
		name: 'object locks on shape, image, and table',
		fn: async () => {
			const b64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.RECTANGLE, {
					x: 1,
					y: 1,
					w: 2,
					h: 1,
					fill: { color: 'FF0000' },
					objectLock: { noMove: true, noResize: true, noRot: true, noChangeShapeType: true, noTextEdit: true },
				})
				s.addImage({
					data: 'image/png;base64,' + b64,
					x: 4,
					y: 1,
					w: 1,
					h: 1,
					objectLock: { noChangeAspect: false, noCrop: true, noMove: true },
				})
				s.addTable([[{ text: 'locked' }]], {
					x: 1,
					y: 3,
					w: 4,
					objectLock: { noGrp: true, noSelect: true, noDrilldown: true },
				})
			})
			await expectNoSchemaErrors(buf, 'object-locks')
		},
	},
	{
		name: 'shape line with round cap',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.LINE, {
					x: 1,
					y: 1,
					w: 4,
					h: 0,
					line: { color: '0070C0', width: 3, cap: 'round', dashType: 'dash' },
				})
			})
			await expectNoSchemaErrors(buf, 'shape-line-round-cap')
		},
	},
	{
		name: 'shape with shadow',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.RECTANGLE, {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					fill: { color: '00B0B9' },
					shadow: { type: 'outer', blur: 6, offset: 2, color: '000000', opacity: 0.15 },
				})
			})
			await expectNoSchemaErrors(buf, 'shape-with-shadow')
		},
	},
	{
		name: 'shape with inner shadow',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.RECTANGLE, {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					fill: { color: '00B0B9' },
					shadow: { type: 'inner', blur: 6, offset: 2, color: '000000', opacity: 0.15 },
				})
			})
			await expectNoSchemaErrors(buf, 'shape-with-inner-shadow')
		},
	},
	{
		name: 'text run shadow in table cell and combined with glow',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				// table cell text has no shape spPr; shadow must emit at the run level (inside <a:rPr>)
				s.addTable(
					[
						[
							{
								text: 'Shadowed cell',
								options: { shadow: { type: 'outer', blur: 4, offset: 3, angle: 45, color: '404040', opacity: 0.6 } },
							},
						],
					],
					{ x: 1, y: 1, w: 4, h: 1 }
				)
				// glow + shadow together must share a single <a:effectLst> (only one allowed per CT_TextCharacterProperties)
				s.addText('Glow and shadow', {
					x: 1,
					y: 3,
					w: 4,
					h: 1,
					glow: { size: 6, color: 'FFFF00', opacity: 0.5 },
					shadow: { type: 'outer', blur: 5, offset: 2, color: '000000', opacity: 0.5 },
				})
			})
			await expectNoSchemaErrors(buf, 'text-run-shadow')
		},
	},
	{
		// RGBA (8-char) effect colors must not emit two <a:alpha> children when the
		// effect also carries an explicit `opacity`. Cell text skips correctShadowOptions,
		// so the RGBA byte reaches createColorElement directly — the caller's opacity wins.
		name: 'RGBA effect color with explicit opacity (shadow + glow)',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addTable([[{ text: 'A', options: { shadow: { type: 'outer', color: '404040CC', opacity: 0.6 } } }]], {
					x: 1,
					y: 1,
					w: 3,
					h: 1,
				})
				s.addText('B', { x: 1, y: 3, w: 3, h: 1, glow: { size: 6, color: 'FFFF0080', opacity: 0.5 } })
			})
			await expectNoSchemaErrors(buf, 'rgba-effect-color-opacity')
		},
	},
	{
		name: 'shape with native linear gradient fill',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.RECTANGLE, {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					fill: {
						type: 'gradient',
						gradient: {
							kind: 'linear',
							angle: 90,
							scaled: true,
							stops: [
								{ position: 0, color: '451DC7' },
								{ position: 100, color: '0B003D', transparency: 10 },
							],
						},
					},
				})
			})
			await expectNoSchemaErrors(buf, 'shape-native-linear-gradient')
		},
	},
	{
		name: 'shape with pattern fill',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.RECTANGLE, {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					fill: {
						type: 'pattern',
						pattern: { preset: 'diagCross', fgColor: '003366', bgColor: 'FFFFFF' },
					},
				})
			})
			await expectNoSchemaErrors(buf, 'shape-pattern-fill')
		},
	},
	{
		name: 'solid-color slide background',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.background = { color: '0088CC' }
				s.addText('hi', { x: 1, y: 1 })
			})
			await expectNoSchemaErrors(buf, 'solid-bg')
		},
	},
	{
		name: 'native linear gradient slide background',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.background = {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 90,
						scaled: true,
						stops: [
							{ position: 0, color: '451DC7' },
							{ position: 100, color: '0B003D' },
						],
					},
				}
				s.addText('hi', { x: 1, y: 1 })
			})
			await expectNoSchemaErrors(buf, 'native-linear-gradient-bg')
		},
	},
	{
		name: 'bullet text',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText('item', { x: 1, y: 1, w: 4, h: 0.5, bullet: true })
			})
			await expectNoSchemaErrors(buf, 'bullet-text')
		},
	},
	{
		name: 'simple table',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addTable(
					[
						[{ text: 'A1' }, { text: 'B1' }],
						[{ text: 'A2' }, { text: 'B2' }],
					],
					{ x: 1, y: 1, w: 4 }
				)
			})
			await expectNoSchemaErrors(buf, 'simple-table')
		},
	},
	{
		name: 'table with hasHeader',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addTable(
					[
						[{ text: 'Col A' }, { text: 'Col B' }],
						[{ text: 'A1' }, { text: 'B1' }],
					],
					{ x: 1, y: 1, w: 4, hasHeader: true }
				)
			})
			await expectNoSchemaErrors(buf, 'table-has-header')
		},
	},
	{
		name: 'table with built-in style and all style flags',
		fn: async () => {
			const { TABLE_STYLE } = await import('../dist/index.js')
			const { buf } = await build((p) => {
				p.addSlide().addTable(
					[
						[{ text: 'Col A' }, { text: 'Col B' }],
						[{ text: 'A1' }, { text: 'B1' }],
						[{ text: 'A2' }, { text: 'B2' }],
						[{ text: 'Total' }, { text: '42' }],
					],
					{
						x: 1,
						y: 1,
						w: 4,
						tableStyle: TABLE_STYLE.MEDIUM_STYLE_2_ACCENT_1,
						hasHeader: true,
						hasFooter: true,
						hasBandedRows: true,
						hasBandedColumns: false,
						hasFirstColumn: false,
						hasLastColumn: false,
					}
				)
			})
			await expectNoSchemaErrors(buf, 'table-built-in-style-all-flags')
		},
	},
	{
		name: 'custom table style exercising every region',
		fn: async () => {
			const { buf } = await build((p) => {
				const style = p.defineTableStyle({
					name: 'Brand & <Banded>',
					wholeTbl: { border: { type: 'solid', color: 'D9D9D9', pt: 0.5 } },
					firstRow: { fill: '1A2B3C', color: 'FFFFFF', bold: true },
					lastRow: { fill: 'CCCCCC', bold: true, italic: true },
					firstCol: { color: '1A2B3C', bold: true },
					lastCol: { color: '1A2B3C' },
					band1H: { fill: 'EAF1F8' },
					band2H: { fill: 'FFFFFF' },
					band1V: { fill: 'F4F7FB' },
					band2V: { fill: 'FFFFFF' },
				})
				p.addSlide().addTable(
					[
						[{ text: 'Col A' }, { text: 'Col B' }, { text: 'Col C' }],
						[{ text: 'A1' }, { text: 'B1' }, { text: 'C1' }],
						[{ text: 'A2' }, { text: 'B2' }, { text: 'C2' }],
						[{ text: 'Total' }, { text: '42' }, { text: '99' }],
					],
					{
						x: 1,
						y: 1,
						w: 6,
						tableStyle: style,
						hasHeader: true,
						hasFooter: true,
						hasBandedRows: true,
						hasBandedColumns: true,
						hasFirstColumn: true,
						hasLastColumn: true,
					}
				)
			})
			await expectNoSchemaErrors(buf, 'custom-table-style-all-regions')
		},
	},
	{
		name: 'custom table style with TRBL border array',
		fn: async () => {
			const { buf } = await build((p) => {
				const style = p.defineTableStyle({
					name: 'Outline Only',
					firstRow: {
						fill: '004400',
						color: 'FFFFFF',
						border: [
							{ type: 'solid', color: '000000', pt: 2 },
							{ type: 'none' },
							{ type: 'dash', color: '888888', pt: 1 },
							{ type: 'none' },
						],
					},
				})
				p.addSlide().addTable(
					[
						[{ text: 'H1' }, { text: 'H2' }],
						[{ text: 'a' }, { text: 'b' }],
					],
					{ x: 1, y: 1, w: 4, tableStyle: style, hasHeader: true }
				)
			})
			await expectNoSchemaErrors(buf, 'custom-table-style-trbl-border')
		},
	},
	{
		name: 'table cell border with line caps',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addTable(
					[
						[
							{
								text: 'capped',
								options: {
									border: [
										{ type: 'solid', color: '000000', pt: 2, cap: 'round' },
										{ type: 'solid', color: '000000', pt: 2, cap: 'square' },
										{ type: 'none', cap: 'round' },
										{ type: 'solid', color: '000000', pt: 2 },
									],
								},
							},
						],
					],
					{ x: 1, y: 1, w: 4 }
				)
			})
			await expectNoSchemaErrors(buf, 'table-cell-border-line-caps')
		},
	},
	{
		name: 'table with merged cells carrying borders and fill (colspan + rowspan)',
		fn: async () => {
			const { buf } = await build((p) => {
				const red = [
					{ type: 'solid', color: 'FF0000', pt: 2 },
					{ type: 'solid', color: 'FF0000', pt: 2 },
					{ type: 'solid', color: 'FF0000', pt: 2 },
					{ type: 'solid', color: 'FF0000', pt: 2 },
				]
				const blue = [
					{ type: 'solid', color: '0000FF', pt: 2 },
					{ type: 'solid', color: '0000FF', pt: 2 },
					{ type: 'solid', color: '0000FF', pt: 2 },
					{ type: 'solid', color: '0000FF', pt: 2 },
				]
				p.addSlide().addTable(
					[
						[
							{ text: 'tall', options: { rowspan: 2, border: blue, fill: { color: 'E0E0FF' } } },
							{ text: 'wide', options: { colspan: 2, border: red, fill: { color: 'FFE0E0' } } },
						],
						[{ text: 'b1' }, { text: 'b2' }],
					],
					{ x: 1, y: 1, w: 8 }
				)
			})
			await expectNoSchemaErrors(buf, 'table-merged-cell-borders')
		},
	},
	{
		name: 'embedded PNG image',
		fn: async () => {
			const b64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { buf } = await build((p) => {
				p.addSlide().addImage({ data: 'image/png;base64,' + b64, x: 1, y: 1, w: 1, h: 1 })
			})
			await expectNoSchemaErrors(buf, 'embedded-png')
		},
	},
	{
		name: 'image clipped to a freeform custGeom path',
		fn: async () => {
			const b64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { buf } = await build((p) => {
				p.addSlide().addImage({
					data: 'image/png;base64,' + b64,
					x: 1,
					y: 1,
					w: 2,
					h: 2,
					points: [{ x: 1, y: 0 }, { x: 2, y: 2 }, { x: 0, y: 2 }, { close: true }],
				})
			})
			await expectNoSchemaErrors(buf, 'image-custgeom')
		},
	},
	{
		name: 'image with duotone recolor',
		fn: async () => {
			const b64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { buf } = await build((p) => {
				p.addSlide().addImage({
					data: 'image/png;base64,' + b64,
					x: 1,
					y: 1,
					w: 2,
					h: 2,
					duotone: { shadow: '250F6B', highlight: 'FFFFFF' },
				})
			})
			await expectNoSchemaErrors(buf, 'image-duotone')
		},
	},
	{
		name: 'image with border line (and shadow) emits a:ln before a:effectLst',
		fn: async () => {
			const b64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { buf } = await build((p) => {
				const s = p.addSlide()
				// solid border + shadow: a:ln must precede a:effectLst per CT_ShapeProperties order
				s.addImage({
					data: 'image/png;base64,' + b64,
					x: 1,
					y: 1,
					w: 2,
					h: 2,
					line: { color: '0088CC', width: 2 },
					shadow: { type: 'outer', color: '000000', opacity: 0.5, blur: 8, offset: 4, angle: 270 },
				})
				// dashed border
				s.addImage({
					data: 'image/png;base64,' + b64,
					x: 4,
					y: 1,
					w: 2,
					h: 2,
					line: { color: '666666', width: 1, dashType: 'dash' },
				})
			})
			await expectNoSchemaErrors(buf, 'image-border-line')
		},
	},
	{
		name: 'image cover/contain sizing emits schema-valid srcRect (incl. negative contain inset)',
		fn: async () => {
			// 1x1 PNG (natural square): cover crops, contain pads with a negative srcRect inset —
			// both must stay schema-valid (CT_RelativeRect permits negative ST_Percentage).
			const b64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addImage({
					data: 'image/png;base64,' + b64,
					x: 0.5,
					y: 0.5,
					w: 4,
					h: 3,
					sizing: { type: 'cover', w: 4, h: 3 },
				})
				s.addImage({
					data: 'image/png;base64,' + b64,
					x: 5,
					y: 0.5,
					w: 4,
					h: 3,
					sizing: { type: 'contain', w: 4, h: 3 },
				})
			})
			await expectNoSchemaErrors(buf, 'image-cover-contain')
		},
	},
	{
		// "Image embedded in a shape": a freeform custGeom clip (spPr) composed with a source
		// crop (srcRect in blipFill) on one picture — the placeholder-equivalent form. Also an
		// arcTo-based half-disc clip. Both must stay schema-valid (CT_Picture child order:
		// blipFill before spPr) with the explicit <a:fillRect/> inside <a:stretch>.
		name: 'image clipped to custGeom AND source-cropped (points + sizing), incl. arcTo',
		fn: async () => {
			const b64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { buf } = await build((p) => {
				const s = p.addSlide()
				// triangular clip + cover crop
				s.addImage({
					data: 'image/png;base64,' + b64,
					x: 1,
					y: 1,
					w: 2,
					h: 3,
					points: [{ x: 1, y: 0 }, { x: 2, y: 0 }, { x: 2, y: 3 }, { x: 0, y: 3 }, { close: true }],
					sizing: { type: 'cover', w: 2, h: 3 },
				})
				// half-disc ("D") clip expressed with an arcTo for the curved edge
				s.addImage({
					data: 'image/png;base64,' + b64,
					x: 5,
					y: 1,
					w: 2,
					h: 3,
					points: [
						{ x: 0.64, y: 0 },
						{ x: 2, y: 0 },
						{ x: 2, y: 3 },
						{ x: 0.64, y: 3 },
						{ x: 0, y: 1.5, curve: { type: 'arc', hR: 1.5, wR: 0.64, stAng: 90, swAng: 180 } },
						{ close: true },
					],
					sizing: { type: 'cover', w: 2, h: 3 },
				})
			})
			await expectNoSchemaErrors(buf, 'image-custgeom-plus-sizing')
		},
	},
	{
		name: 'text caps: all-caps and small-caps run properties',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addText(
					[
						{ text: 'ALL CAPS ', options: { caps: 'all' } },
						{ text: 'Small Caps ', options: { caps: 'small' } },
						{ text: 'Normal', options: { caps: 'none' } },
					],
					{ x: 1, y: 1, w: 6, h: 0.5 }
				)
			})
			await expectNoSchemaErrors(buf, 'text-caps')
		},
	},
	{
		name: 'scatter chart with valAxisCrossesAt zero',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addChart(
					p.charts.SCATTER,
					[
						{ name: 'X-Axis', values: [0, 1, 2] },
						{ name: 'Y-Value 1', values: [1, 4, 9], labels: ['A', 'B', 'C'] },
					],
					{ x: 1, y: 1, w: 6, h: 3, valAxisCrossesAt: 0 }
				)
			})
			await expectNoSchemaErrors(buf, 'scatter-val-axis-crosses-at-zero')
		},
	},
	{
		name: 'scatter chart with independent axis format codes',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addChart(
					p.charts.SCATTER,
					[
						{ name: 'X-Axis', values: [0, 1, 2] },
						{ name: 'Y-Value 1', values: [1, 4, 9], labels: ['A', 'B', 'C'] },
					],
					{ x: 1, y: 1, w: 6, h: 3, catAxisLabelFormatCode: '0.0', valAxisLabelFormatCode: '#,##0' }
				)
			})
			await expectNoSchemaErrors(buf, 'scatter-independent-axis-format-codes')
		},
	},
	{
		// upstream #1309: value number format must reach the series numCache (and stay schema-valid)
		// so PowerPoint/Google Slides honor it, not just LibreOffice via the dLbls mask.
		name: 'charts with dataLabelFormatCode in the value numCache (bar, pie, scatter)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(p.charts.BAR, [{ name: 'S1', labels: ['A', 'B', 'C'], values: [0.1, 0.2, 0.3] }], {
					x: 0.5,
					y: 0.5,
					w: 4,
					h: 3,
					showValue: true,
					dataLabelFormatCode: '0%',
				})
				p.addSlide().addChart(p.charts.PIE, [{ name: 'S1', labels: ['A', 'B', 'C'], values: [0.5, 0.3, 0.2] }], {
					x: 0.5,
					y: 0.5,
					w: 4,
					h: 3,
					showPercent: true,
					dataLabelFormatCode: '0%',
				})
				p.addSlide().addChart(
					p.charts.SCATTER,
					[
						{ name: 'X-Axis', values: [0, 1, 2] },
						{ name: 'Y-Value 1', values: [0.1, 0.4, 0.9], labels: ['A', 'B', 'C'] },
					],
					{ x: 0.5, y: 0.5, w: 4, h: 3, showValue: true, dataLabelFormatCode: '0.0%' }
				)
			})
			await expectNoSchemaErrors(buf, 'chart-value-format-code-numcache')
		},
	},
	{
		// Upstream #744: bubble/bubble3D charts can show each bubble's size as a data label.
		// The `showBubbleSize` option flips the previously hard-coded <c:showBubbleSize val="0"/>;
		// lock in that the enabled flag stays schema-valid in CT_DLbls.
		name: 'bubble charts show bubble-size data labels (upstream #744)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.BUBBLE,
					[
						{ name: 'X-Axis', values: [1, 2, 3, 4] },
						{ name: 'Y-Values 1', values: [13, 20, 21, 25], sizes: [10, 5, 20, 15] },
					],
					{ x: 0.5, y: 0.5, w: 6, h: 3, showBubbleSize: true }
				)
				p.addSlide().addChart(
					p.charts.BUBBLE3D,
					[
						{ name: 'X-Axis', values: [1, 2, 3, 4] },
						{ name: 'Y-Values 1', values: [13, 20, 21, 25], sizes: [10, 5, 20, 15] },
					],
					{ x: 0.5, y: 0.5, w: 6, h: 3, showBubbleSize: true }
				)
			})
			await expectNoSchemaErrors(buf, 'chart-bubble-size-data-label')
		},
	},
	{
		// Upstream #1420: chart text fonts (title, legend, axis labels, data labels) emit the
		// `<a:latin>/<a:ea>/<a:cs>` typeface trio so East Asian text honors the requested font.
		// Lock in that the ea/cs additions stay schema-valid (correct CT_TextCharacterProperties order).
		name: 'chart text fonts emit schema-valid latin/ea/cs typeface trio (upstream #1420)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(p.charts.BAR, [{ name: '系列', labels: ['甲', '乙', '丙'], values: [1, 2, 3] }], {
					x: 0.5,
					y: 0.5,
					w: 6,
					h: 3,
					showTitle: true,
					title: '图表标题',
					titleFontFace: 'Microsoft YaHei',
					showLegend: true,
					legendFontFace: 'SimSun',
					showValue: true,
					dataLabelFontFace: 'NSimSun',
					catAxisLabelFontFace: 'KaiTi',
					valAxisLabelFontFace: 'FangSong',
				})
			})
			await expectNoSchemaErrors(buf, 'chart-east-asian-font-trio')
		},
	},
	{
		name: 'bar chart with valAxisCrossBetween midCat',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addChart(p.charts.BAR, [{ name: 'Series 1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
					x: 1,
					y: 1,
					w: 6,
					h: 3,
					valAxisCrossBetween: 'midCat',
				})
			})
			await expectNoSchemaErrors(buf, 'bar-chart-cross-between-midcat')
		},
	},
	{
		name: 'stacked bar chart with series lines (upstream #1329)',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addChart(
					p.charts.BAR,
					[
						{ name: 'Series 1', labels: ['A', 'B', 'C'], values: [1, 2, 3] },
						{ name: 'Series 2', labels: ['A', 'B', 'C'], values: [2, 1, 2] },
					],
					{
						x: 1,
						y: 1,
						w: 6,
						h: 3,
						barGrouping: 'stacked',
						barSeriesLine: { color: '777777', size: 1, style: 'dash' },
					}
				)
			})
			await expectNoSchemaErrors(buf, 'bar-chart-series-lines-styled')
		},
	},
	{
		name: 'stacked bar chart with automatic series lines (upstream #1329)',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addChart(
					p.charts.BAR,
					[
						{ name: 'Series 1', labels: ['A', 'B', 'C'], values: [1, 2, 3] },
						{ name: 'Series 2', labels: ['A', 'B', 'C'], values: [2, 1, 2] },
					],
					{ x: 1, y: 1, w: 6, h: 3, barGrouping: 'stacked', barSeriesLine: true }
				)
			})
			await expectNoSchemaErrors(buf, 'bar-chart-series-lines-auto')
		},
	},
	{
		name: 'chart with non-finite (NaN) values emits a valid sparse numCache (upstream #1357)',
		fn: async () => {
			const warnings = []
			const origWarn = console.warn
			console.warn = (...args) => warnings.push(args.join(' '))
			let buf
			try {
				;({ buf } = await build((p) => {
					const s = p.addSlide()
					s.addChart(
						p.charts.BAR,
						[
							{ name: 'S1', labels: ['A', 'B', 'C', 'D'], values: [5, NaN, 3, NaN] },
							{ name: 'S2', labels: ['A', 'B', 'C', 'D'], values: [2, 4, NaN, 1] },
						],
						{ x: 1, y: 1, w: 6, h: 3, barGrouping: 'stacked' }
					)
				}))
			} finally {
				console.warn = origWarn
			}
			// NaN data points must be dropped (not emitted as invalid <c:v>NaN</c:v>) and warned about.
			assert(
				warnings.some((w) => w.includes('not a finite number')),
				'expected a warning for non-finite chart values'
			)
			await expectNoSchemaErrors(buf, 'bar-chart-nonfinite-values')
		},
	},
	{
		name: 'line chart marker size out of range is clamped to valid ST_MarkerSize (upstream #1233)',
		fn: async () => {
			const warnings = []
			const origWarn = console.warn
			console.warn = (...args) => warnings.push(args.join(' '))
			let buf
			try {
				;({ buf } = await build((p) => {
					// 1 (below min 2), 100 (above max 72), and 5.5 (non-integer) all violate
					// ST_MarkerSize (integer 2-72) and would trigger PowerPoint repair if emitted as-is.
					const s = p.addSlide()
					s.addChart(p.charts.LINE, [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
						x: 1,
						y: 1,
						w: 6,
						h: 3,
						lineDataSymbolSize: 1,
					})
					const s2 = p.addSlide()
					s2.addChart(p.charts.LINE, [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
						x: 1,
						y: 1,
						w: 6,
						h: 3,
						lineDataSymbolSize: 100,
					})
				}))
			} finally {
				console.warn = origWarn
			}
			assert(
				warnings.some((w) => w.includes('valid marker size range')),
				'expected a warning for out-of-range lineDataSymbolSize'
			)
			await expectNoSchemaErrors(buf, 'line-chart-marker-size-clamped')
		},
	},
	{
		name: 'out-of-range chart gap/overlap/holeSize/firstSliceAng are clamped to valid ranges (upstream #1233)',
		fn: async () => {
			const warnings = []
			const origWarn = console.warn
			console.warn = (...args) => warnings.push(args.join(' '))
			let buf
			try {
				;({ buf } = await build((p) => {
					// gapWidth 600 (>500), overlap 200 (>100) violate ST_GapAmount / ST_Overlap.
					p.addSlide().addChart(
						p.charts.BAR,
						[
							{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] },
							{ name: 'S2', labels: ['A', 'B', 'C'], values: [2, 3, 1] },
						],
						{ x: 1, y: 1, w: 6, h: 3, barGapWidthPct: 600, barOverlapPct: 200 }
					)
					// holeSize 200 (>90) violates ST_HoleSize; firstSliceAng 400 (>360) violates ST_FirstSliceAng.
					p.addSlide().addChart(p.charts.DOUGHNUT, [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
						x: 1,
						y: 1,
						w: 6,
						h: 3,
						holeSize: 200,
						firstSliceAng: 400,
					})
				}))
			} finally {
				console.warn = origWarn
			}
			assert(
				warnings.some((w) => w.includes('barOverlapPct')) && warnings.some((w) => w.includes('holeSize')),
				'expected warnings for out-of-range chart options'
			)
			await expectNoSchemaErrors(buf, 'chart-bounded-attrs-clamped')
		},
	},
	{
		name: 'out-of-range text fontSize/charSpacing/lineSpacing are clamped to valid ranges',
		fn: async () => {
			const warnings = []
			const origWarn = console.warn
			console.warn = (...args) => warnings.push(args.join(' '))
			let buf
			try {
				;({ buf } = await build((p) => {
					// fontSize 5000pt -> sz 500000 (>400000), charSpacing 5000pt -> spc 500000 (>400000),
					// lineSpacing 2000pt -> spcPts 200000 (>158400): all violate their ST_Text* ranges.
					p.addSlide().addText('Too big', {
						x: 1,
						y: 1,
						w: 6,
						h: 2,
						fontSize: 5000,
						charSpacing: 5000,
						lineSpacing: 2000,
					})
					// Negatives violate the lower bounds (sz >= 100, spcPts >= 0).
					p.addSlide().addText('Negative', { x: 1, y: 1, w: 6, h: 2, fontSize: -10, charSpacing: -5000 })
					// Same surfaces inside a table cell (shares the run-property emission path).
					p.addSlide().addTable([[{ text: 'Cell', options: { fontSize: 5000 } }]], { x: 1, y: 1, w: 4 })
				}))
			} finally {
				console.warn = origWarn
			}
			assert(
				warnings.some((w) => w.includes('fontSize')) && warnings.some((w) => w.includes('charSpacing')),
				'expected warnings for out-of-range text options'
			)
			await expectNoSchemaErrors(buf, 'text-bounded-attrs-clamped')
		},
	},
	{
		name: 'out-of-range shape transparency/line-width are clamped to valid ranges',
		fn: async () => {
			const warnings = []
			const origWarn = console.warn
			console.warn = (...args) => warnings.push(args.join(' '))
			let buf
			try {
				;({ buf } = await build((p) => {
					const s = p.addSlide()
					// transparency 150 / -20 push <a:alpha> outside ST_PositiveFixedPercentage (0..100000).
					s.addShape(p.shapes.RECTANGLE, { x: 1, y: 1, w: 2, h: 1, fill: { color: 'FF0000', transparency: 150 } })
					s.addShape(p.shapes.RECTANGLE, { x: 4, y: 1, w: 2, h: 1, fill: { color: '00FF00', transparency: -20 } })
					// line width 2000pt -> w 25.4M EMU (>20116800), and a negative width, both violate ST_LineWidth.
					s.addShape(p.shapes.RECTANGLE, { x: 1, y: 3, w: 2, h: 1, line: { color: '0000FF', width: 2000 } })
					s.addShape(p.shapes.RECTANGLE, { x: 4, y: 3, w: 2, h: 1, line: { color: '0000FF', width: -5 } })
					// glow opacity 5 (valid 0-1) pushes <a:alpha> above 100000.
					s.addText('Glow', { x: 1, y: 5, w: 4, h: 1, glow: { size: 10, color: 'FF0000', opacity: 5 } })
				}))
			} finally {
				console.warn = origWarn
			}
			assert(
				warnings.some((w) => w.includes('transparency')) && warnings.some((w) => w.includes('line width')),
				'expected warnings for out-of-range shape options'
			)
			await expectNoSchemaErrors(buf, 'shape-bounded-attrs-clamped')
		},
	},
	{
		name: 'chart title with y-only manual layout (auto horizontal centering)',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addChart(p.charts.BAR, [{ name: 'Series 1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
					x: 1,
					y: 1,
					w: 6,
					h: 3,
					showTitle: true,
					title: 'Centered title, nudged down',
					titlePos: { y: 0.3 },
				})
			})
			await expectNoSchemaErrors(buf, 'chart-title-y-only-manual-layout')
		},
	},
	{
		name: 'chart title with italic and underline styling (upstream #1188)',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addChart(p.charts.BAR, [{ name: 'Series 1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
					x: 1,
					y: 1,
					w: 6,
					h: 3,
					showTitle: true,
					title: 'Italic underlined title',
					titleBold: true,
					titleItalic: true,
					titleUnderline: true,
				})
			})
			await expectNoSchemaErrors(buf, 'chart-title-italic-underline')
		},
	},
	{
		name: 'bullet color (buClr separate from text color)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText('item', {
					x: 1,
					y: 1,
					w: 4,
					h: 0.5,
					bullet: { color: 'FF0000', characterCode: '2022' },
					color: '000000',
				})
			})
			await expectNoSchemaErrors(buf, 'bullet-color')
		},
	},
	{
		name: 'bullet glyph font and size (buFont + buSzPct)',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addText('wingding', {
					x: 1,
					y: 1,
					w: 4,
					h: 0.5,
					bullet: { characterCode: 'F0E0', fontFace: 'Wingdings', size: 80 },
				})
				s.addText('numbered', { x: 1, y: 2, w: 4, h: 0.5, bullet: { type: 'number', fontFace: 'Arial', size: 150 } })
			})
			await expectNoSchemaErrors(buf, 'bullet-font-size')
		},
	},
	{
		name: 'shrink-text fit with explicit fontScale/lnSpcReduction',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText('shrink me', {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					fit: { type: 'shrink', fontScale: 85, lnSpcReduction: 20 },
				})
			})
			await expectNoSchemaErrors(buf, 'text-fit-shrink-normautofit')
		},
	},
	{
		name: 'line chart with transparent marker fill',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(p.charts.LINE, [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
					x: 1,
					y: 1,
					w: 6,
					h: 3,
					chartColors: ['transparent'],
				})
			})
			await expectNoSchemaErrors(buf, 'line-chart-transparent-marker')
		},
	},
	{
		name: 'line chart with null values defaults to gap',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(p.charts.LINE, [{ name: 'S1', labels: ['A', 'B', 'C', 'D'], values: [1, null, 3, 4] }], {
					x: 1,
					y: 1,
					w: 6,
					h: 3,
				})
			})
			await expectNoSchemaErrors(buf, 'line-chart-null-values-gap')
		},
	},
	{
		name: 'line chart with per-series lineDashValues',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.LINE,
					[
						{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] },
						{ name: 'S2', labels: ['A', 'B', 'C'], values: [4, 3, 2] },
						{ name: 'S3', labels: ['A', 'B', 'C'], values: [2, 4, 1] },
					],
					{ x: 1, y: 1, w: 6, h: 3, lineDashValues: ['solid', 'dash', 'lgDashDot'] }
				)
			})
			await expectNoSchemaErrors(buf, 'line-chart-per-series-dash')
		},
	},
	{
		name: 'image hyperlink with query-string ampersand produces valid XML',
		fn: async () => {
			const b64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { buf } = await build((p) => {
				p.addSlide().addImage({
					data: 'image/png;base64,' + b64,
					x: 1,
					y: 1,
					w: 2,
					h: 2,
					hyperlink: { url: 'https://example.com/page?a=1&b=2&c=3' },
				})
			})
			await expectNoSchemaErrors(buf, 'image-hyperlink-query-string')
		},
	},
	{
		name: 'combo chart with per-subchart legend suppression',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					[
						{
							type: p.charts.BAR,
							data: [{ name: 'Revenue', labels: ['Q1', 'Q2', 'Q3'], values: [10, 20, 30] }],
							options: {},
						},
						{
							type: p.charts.LINE,
							data: [{ name: 'Target', labels: ['Q1', 'Q2', 'Q3'], values: [15, 15, 15] }],
							options: { showLegend: false },
						},
					],
					{ x: 1, y: 1, w: 6, h: 3, showLegend: true }
				)
			})
			await expectNoSchemaErrors(buf, 'combo-chart-subchart-legend-suppress')
		},
	},
	{
		// #1355: a scatter subchart in a combo chart needs its category (X) axis
		// emitted as a <c:valAx>, not a <c:catAx>. Emitting a catAx made
		// PowerPoint flag the file for repair. Scatter rides the secondary axes.
		name: 'combo chart with bar and scatter on secondary axes',
		fn: async () => {
			const labels = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					[
						{
							type: p.charts.BAR,
							data: [{ name: 'Bottom', labels, values: [17, 26, 53, 10, 4] }],
							options: { barDir: 'bar', barGrouping: 'clustered' },
						},
						{
							type: p.charts.SCATTER,
							data: [
								{ name: 'X-Axis', labels, values: [1, 2, 3, 4, 5] },
								{ name: 'Y', labels, values: [25, 35, 55, 10, 5] },
							],
							options: { secondaryValAxis: true, secondaryCatAxis: true },
						},
					],
					{
						x: 1,
						y: 1,
						w: 6,
						h: 3,
						showLegend: false,
						valAxes: [{ valAxisTitle: 'Primary' }, { valAxisTitle: 'Secondary' }],
						catAxes: [{ catAxisTitle: 'Primary Cat' }, { catAxisHidden: true }],
					}
				)
			})
			await expectNoSchemaErrors(buf, 'combo-bar-scatter-secondary-axes')
		},
	},
	{
		name: 'chart with per-series color and data-label overrides',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.BAR,
					[
						{ name: 'Alpha', labels: ['Q1', 'Q2', 'Q3'], values: [10, 20, 30] },
						{ name: 'Beta', labels: ['Q1', 'Q2', 'Q3'], values: [15, 25, 5] },
						{ name: 'Gamma', labels: ['Q1', 'Q2', 'Q3'], values: [5, 10, 20] },
					],
					{
						x: 1,
						y: 1,
						w: 6,
						h: 3,
						showValue: true,
						dataLabelColor: '000000',
						dataLabelFontSize: 10,
						seriesOptions: [
							{ color: 'FF0000', dataLabelColor: 'FFFFFF', dataLabelFontBold: true },
							{ color: '00AA00', dataLabelFontSize: 14, dataLabelFontItalic: true },
							{ lineSize: 0 },
						],
					}
				)
			})
			await expectNoSchemaErrors(buf, 'chart-series-options')
		},
	},
	{
		name: 'line chart with per-series lineSize overrides',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.LINE,
					[
						{ name: 'Thick', labels: ['Jan', 'Feb', 'Mar'], values: [1, 2, 3] },
						{ name: 'Normal', labels: ['Jan', 'Feb', 'Mar'], values: [3, 2, 1] },
						{ name: 'Hidden', labels: ['Jan', 'Feb', 'Mar'], values: [2, 2, 2] },
					],
					{
						x: 1,
						y: 1,
						w: 6,
						h: 3,
						lineSize: 2,
						seriesOptions: [{ lineSize: 6 }, {}, { lineSize: 0 }],
					}
				)
			})
			await expectNoSchemaErrors(buf, 'chart-series-linesize-overrides')
		},
	},
	{
		name: 'firstSlideNum sets presentation starting slide number',
		fn: async () => {
			const { buf } = await build((p) => {
				p.firstSlideNum = 5
				const slide = p.addSlide()
				slide.addText('', { x: 0, y: 0, w: 1, h: 1, slideNumber: { x: 0.5, y: 0.5 } })
			})
			await expectNoSchemaErrors(buf, 'first-slide-num')
		},
	},
	{
		name: 'round2SameRect and round2DiagRect with rectRadius emit adj1/adj2',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.shapes.ROUND_2_SAME_RECTANGLE, {
					x: 0.5,
					y: 0.5,
					w: 3,
					h: 2,
					rectRadius: 0.1,
					fill: { color: '4472C4' },
				})
				s.addShape(p.shapes.ROUND_2_DIAG_RECTANGLE, {
					x: 4,
					y: 0.5,
					w: 3,
					h: 2,
					rectRadius: 0.15,
					fill: { color: 'ED7D31' },
				})
			})
			await expectNoSchemaErrors(buf, 'round2-rect-adj1-adj2')
		},
	},
	{
		name: 'shapeAdjust emits avLst guides for preset shapes (single + array)',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				// Single guide: chevron point depth
				s.addShape(p.shapes.CHEVRON, {
					x: 0.5,
					y: 0.5,
					w: 3,
					h: 1,
					shapeAdjust: { name: 'adj', value: 0.25 },
					fill: { color: '4472C4' },
				})
				// Array form on a rounded-rectangle adjust handle
				s.addShape(p.shapes.ROUNDED_RECTANGLE, {
					x: 0.5,
					y: 2,
					w: 3,
					h: 1,
					shapeAdjust: [{ name: 'adj', value: 0.5 }],
					fill: { color: 'ED7D31' },
				})
			})
			await expectNoSchemaErrors(buf, 'shape-adjust-avlst-guides')
		},
	},
	{
		name: 'bar chart with multi-level category labels (multiLvlStrRef)',
		fn: async () => {
			const LABELS = [
				['Gear', 'Berg', 'Motr', 'Swch', 'Plug', 'Cord', 'Pump', 'Leak', 'Seal'],
				['Mech', '', '', 'Elec', '', '', 'Hydr', '', ''],
			]
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.BAR,
					[
						{ name: 'West', labels: LABELS, values: [11, 8, 3, 0, 11, 3, 0, 0, 0] },
						{ name: 'Ctrl', labels: LABELS, values: [0, 11, 6, 19, 12, 5, 0, 0, 0] },
						{ name: 'East', labels: LABELS, values: [0, 3, 2, 0, 0, 0, 4, 3, 1] },
					],
					{ x: 1, y: 1, w: 6, h: 4 }
				)
			})
			await expectNoSchemaErrors(buf, 'bar-chart-multilevel-categories')
		},
	},
	{
		name: 'custom document properties (string, integer, float, boolean, date)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.setCustomProperty('Author', 'Jane Smith')
				p.setCustomProperty('Version', 3)
				p.setCustomProperty('Score', 1.5)
				p.setCustomProperty('Published', true)
				p.setCustomProperty('CreatedAt', new Date('2026-01-01T00:00:00Z'))
				p.addSlide()
			})
			await expectNoSchemaErrors(buf, 'custom-document-properties')
		},
	},
	{
		name: 'slide master roundRect object and roundRect placeholder',
		fn: async () => {
			const { buf } = await build((p) => {
				p.defineSlideMaster({
					title: 'ROUNDRECT_MASTER',
					objects: [
						{
							roundRect: {
								x: 0.5,
								y: 0.5,
								w: 2,
								h: 1,
								rectRadius: 0.1,
								fill: { color: 'E8F0FE' },
							},
						},
						{
							placeholder: {
								options: {
									name: 'title',
									type: 'title',
									x: 0.5,
									y: 2,
									w: 9,
									h: 1.5,
									shape: 'roundRect',
									rectRadius: 0.15,
								},
								text: '',
							},
						},
					],
				})
				p.addSlide({ masterName: 'ROUNDRECT_MASTER' })
			})
			await expectNoSchemaErrors(buf, 'slide-master-roundrect')
		},
	},
	{
		name: 'bar chart with per-point customLabels',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.BAR,
					[
						{ name: 'Series 1', labels: ['A', 'B', 'C'], values: [10, 20, 30], customLabels: ['Low', '', 'High'] },
						{ name: 'Series 2', labels: ['A', 'B', 'C'], values: [15, 5, 25], customLabels: ['', 'Min', ''] },
					],
					{ x: 1, y: 1, w: 6, h: 3, showValue: true }
				)
			})
			await expectNoSchemaErrors(buf, 'bar-chart-custom-labels')
		},
	},
	{
		name: 'pie chart with per-point customLabels',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.PIE,
					[
						{
							name: 'Status',
							labels: ['Red', 'Amber', 'Green'],
							values: [10, 30, 60],
							customLabels: ['At Risk', 'Watch', 'On Track'],
						},
					],
					{ x: 1, y: 1, w: 4, h: 3, showValue: true }
				)
			})
			await expectNoSchemaErrors(buf, 'pie-chart-custom-labels')
		},
	},
	{
		name: 'bar chart with per-point pointStyles (border + fill)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.BAR,
					[
						{
							name: 'Status',
							labels: ['A', 'B', 'C', 'D'],
							values: [10, 20, 38, 2],
							pointStyles: [
								{ border: { pt: 2, color: 'FF0000' } },
								{},
								{ fill: '00B050', border: { type: 'dash', color: '404040' } },
								{ border: { type: 'none' } },
							],
						},
					],
					{ x: 1, y: 1, w: 6, h: 3 }
				)
			})
			await expectNoSchemaErrors(buf, 'bar-chart-point-styles')
		},
	},
	{
		name: 'pie chart with per-point pointStyles (border + fill)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.PIE,
					[
						{
							name: 'Status',
							labels: ['Red', 'Amber', 'Green'],
							values: [10, 30, 60],
							pointStyles: [{ border: { pt: 3, color: 'C00000' } }, {}, { fill: '70AD47' }],
						},
					],
					{ x: 1, y: 1, w: 4, h: 3 }
				)
			})
			await expectNoSchemaErrors(buf, 'pie-chart-point-styles')
		},
	},
	{
		name: 'pie/doughnut charts with configurable leader-line color and size',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.PIE,
					[{ name: 'Status', labels: ['Red', 'Amber', 'Green'], values: [10, 30, 60] }],
					{
						x: 1,
						y: 1,
						w: 4,
						h: 3,
						showPercent: true,
						showLeaderLines: true,
						leaderLineColor: 'FF0000',
						leaderLineSize: 1.5,
					}
				)
				p.addSlide().addChart(p.charts.DOUGHNUT, [{ name: 'Status', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
					x: 1,
					y: 1,
					w: 4,
					h: 3,
					showLeaderLines: true,
					leaderLineColor: '0070C0',
				})
			})
			await expectNoSchemaErrors(buf, 'pie-chart-leader-line-color')
		},
	},
	{
		name: 'line chart with per-point pointStyles (border)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.LINE,
					[
						{
							name: 'Series 1',
							labels: ['A', 'B', 'C', 'D'],
							values: [4, 8, 6, 10],
							pointStyles: [
								{},
								{ border: { pt: 2, color: 'FF0000' } },
								{},
								{ border: { type: 'dash', color: '0070C0' } },
							],
						},
					],
					{ x: 1, y: 1, w: 6, h: 3 }
				)
			})
			await expectNoSchemaErrors(buf, 'line-chart-point-styles')
		},
	},
	{
		name: 'area chart with per-point pointStyles (border + fill)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.AREA,
					[
						{
							name: 'Series 1',
							labels: ['A', 'B', 'C'],
							values: [5, 9, 7],
							pointStyles: [{ fill: 'FFC000' }, {}, { border: { pt: 1, color: '404040' } }],
						},
					],
					{ x: 1, y: 1, w: 6, h: 3 }
				)
			})
			await expectNoSchemaErrors(buf, 'area-chart-point-styles')
		},
	},
	{
		name: 'scatter chart with per-point pointStyles (border)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.charts.SCATTER,
					[
						{ name: 'X-Axis', values: [1, 2, 3, 4] },
						{
							name: 'Y-Values',
							values: [3, 6, 2, 8],
							pointStyles: [{ border: { pt: 2, color: 'FF0000' } }, {}, { fill: '00B050' }, {}],
						},
					],
					{ x: 1, y: 1, w: 6, h: 3 }
				)
			})
			await expectNoSchemaErrors(buf, 'scatter-chart-point-styles')
		},
	},
	{
		// upstream #1243: theme color scheme overrides must stay schema-valid, including dk1/lt1
		// switching from <a:sysClr> to <a:srgbClr> when overridden.
		name: 'theme color scheme overrides (incl. dk1/lt1 as srgbClr)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.theme = {
					colorScheme: {
						dk1: '101010',
						lt1: 'FAFAFA',
						dk2: '1F3864',
						lt2: 'D9D9D9',
						accent1: 'C00000',
						accent2: '00B050',
						accent3: '0070C0',
						accent4: '7030A0',
						accent5: 'FFC000',
						accent6: '00B0F0',
						hlink: '0563C1',
						folHlink: '954F72',
					},
				}
				p.addSlide().addText('themed', { x: 1, y: 1, w: 4, h: 0.5, color: p.SchemeColor.accent1 })
			})
			await expectNoSchemaErrors(buf, 'theme-color-scheme')
		},
	},
	{
		// upstream #1059: connectors emit <p:cxnSp> with connector preset geometries and must stay
		// schema-valid, including flipped boxes and arrowheads/dashes on the <a:ln>.
		name: 'connectors (straight/elbow/curved, flipped, arrowheads)',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addConnector({
					type: 'straight',
					x1: 1,
					y1: 1,
					x2: 4,
					y2: 3,
					color: 'FF0000',
					width: 2,
					endArrowType: 'triangle',
				})
				s.addConnector({ type: 'elbow', x1: 6, y1: 4, x2: 2, y2: 1, dashType: 'dash', beginArrowType: 'oval' })
				s.addConnector({ type: 'curved', x1: 1, y1: 5, x2: 5, y2: 6 })
				// Bend control: adjustable jogs emit <a:gd name="adjN"> guides that must stay schema-valid.
				s.addConnector({ type: 'elbow', x1: 7, y1: 1, x2: 9, y2: 3, adj: 25 })
				s.addConnector({ type: 'elbow', x1: 6, y1: 5, x2: 9, y2: 6, bends: 2, adj: [30, 70] })
				s.addConnector({ type: 'curved', x1: 7, y1: 4, x2: 9, y2: 5, bends: 3, adj: [10, 50, 90] })
			})
			await expectNoSchemaErrors(buf, 'connectors')
		},
	},
]
