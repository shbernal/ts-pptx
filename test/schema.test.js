// Schema-validation fixtures. Each case builds a representative `.pptx`
// and asserts the OpenXmlValidator (via OOXMLValidatorCLI) reports no
// errors.
//
// Fixtures are intentionally small and orthogonal — they exercise one
// API surface each — so when an error appears we can localise it.
//
// Run with: pnpm run test:schema

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import PptxGenJS from '../dist/node.js'
import { latexToOmml } from '../dist/math.js'
import { build, assert, readEntry, assertIncludes, firstXmlBlock, listEntries } from './helpers.js'
import { validateBuf } from './validator.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const fontsDir = path.join(__dirname, 'read', 'fixtures', 'fonts')

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
		// upstream-issue-1298: a standalone title text box (no matching layout placeholder)
		// emits <p:ph type="title"/> so PowerPoint sees an accessible slide title. Assert the
		// resulting package is schema-valid.
		name: 'standalone title placeholder text box',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText('Accessible Title', { x: 0.5, y: 0.3, w: 9, h: 1, fontSize: 32, placeholder: 'title' })
			})
			await expectNoSchemaErrors(buf, 'standalone-title-placeholder')
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
		// a particular layout (see docs/backlog.yml dn-doc-render-caveats).
		name: 'text box with margins',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addText('hello', { x: 1, y: 1, w: 4, h: 1, margin: [0.1, 0.05, 0.1, 0.05] })
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
		// textDirection is the documented public option; it must reach <a:bodyPr vert="…">
		// (ST_TextVerticalType). Previously only the undocumented `vert` alias was honored
		// for text boxes, so textDirection was silently dropped (dn-text-direction-serialization).
		name: 'text box with textDirection emits bodyPr vert (dn-text-direction-serialization)',
		fn: async () => {
			const { buf, zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('rotated', { x: 1, y: 1, w: 4, h: 2, textDirection: 'vert270' })
			})
			const slideXml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertIncludes(slideXml, 'vert="vert270"', 'textDirection vert270')
			await expectNoSchemaErrors(buf, 'text-direction')
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
		// and not asserted here — see docs/backlog.yml dn-doc-render-caveats.
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
				s.addShape(p.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, fill: { color: 'FF0000' } })
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
				s.addShape(p.ShapeType.rect, {
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
				s.addShape(p.ShapeType.line, {
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
				s.addShape(p.ShapeType.rect, {
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
				s.addShape(p.ShapeType.rect, {
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
				s.addShape(p.ShapeType.rect, {
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
		name: 'shape with gradient line stroke',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				// Gradient stroke via `line.gradient` (no fill) — `<a:gradFill>` inside `<a:ln>`.
				s.addShape(p.ShapeType.line, {
					x: 1,
					y: 1,
					w: 4,
					h: 0,
					line: {
						width: 3,
						gradient: {
							kind: 'linear',
							angle: 0,
							stops: [
								{ position: 0, color: 'accent3' },
								{ position: 100, color: 'accent4' },
							],
						},
					},
				})
				// Gradient border around a filled rectangle (stroke + fill coexisting).
				s.addShape(p.ShapeType.rect, {
					x: 1,
					y: 2,
					w: 4,
					h: 1,
					fill: { color: 'FFFFFF' },
					line: {
						width: 2,
						type: 'gradient',
						gradient: {
							kind: 'linear',
							angle: 45,
							stops: [
								{ position: 0, color: 'FF0000' },
								{ position: 100, color: '0000FF', transparency: 20 },
							],
						},
					},
				})
			})
			await expectNoSchemaErrors(buf, 'shape-gradient-line-stroke')
		},
	},
	{
		name: 'shape with pattern fill',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.ShapeType.rect, {
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
		name: 'shape with image (blip) fill',
		fn: async () => {
			const pngData =
				'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addShape(p.ShapeType.rect, {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					fill: { type: 'image', image: { data: pngData } },
				})
			})
			await expectNoSchemaErrors(buf, 'shape-image-fill')
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
		name: 'table with fitColumns shrink-to-fit (upstream-issue-1451)',
		fn: async () => {
			const { buf, zip } = await build((p) => {
				p.addSlide().addTable(
					[
						[{ text: 'A' }, { text: 'B' }, { text: 'C' }, { text: 'D' }],
						[{ text: '1' }, { text: '2' }, { text: '3' }, { text: '4' }],
					],
					{ x: 0.5, y: 1, colW: [4, 4, 4, 4], fitColumns: 'shrink' }
				)
			})
			const slideXml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const cols = [...slideXml.matchAll(/<a:gridCol w="(\d+)"\/>/g)].map((m) => Number(m[1]))
			const sum = cols.reduce((a, b) => a + b, 0)
			// 16in of columns scaled to the 9in usable width (10in slide - 0.5 x - 0.5 margin).
			assert(Math.abs(sum - 9 * 914400) <= cols.length, `expected gridCol sum ~9in; got ${sum / 914400}`)
			await expectNoSchemaErrors(buf, 'table-fit-columns')
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
		name: 'table with headerRow inline styling (upstream-issue-1256)',
		fn: async () => {
			const { buf, zip } = await build((p) => {
				p.addSlide().addTable(
					[
						[{ text: 'Col A' }, { text: 'Col B' }],
						[{ text: 'A1' }, { text: 'B1' }],
					],
					{ x: 1, y: 1, w: 4, headerRow: { fill: { color: '1A2B3C' }, color: 'FFFFFF', bold: true } }
				)
			})
			const slideXml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertIncludes(slideXml, 'firstRow="1"', 'headerRow implies hasHeader')
			await expectNoSchemaErrors(buf, 'table-header-row')
		},
	},
	{
		name: 'table with rtl emits rtl="1" on tblPr (upstream #1291)',
		fn: async () => {
			const { buf, zip } = await build((p) => {
				p.addSlide().addTable(
					[
						[{ text: 'Col A' }, { text: 'Col B' }],
						[{ text: 'A1' }, { text: 'B1' }],
					],
					{ x: 1, y: 1, w: 4, rtl: true }
				)
			})
			const slideXml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertIncludes(slideXml, '<a:tblPr rtl="1"', 'rtl table')
			await expectNoSchemaErrors(buf, 'table-rtl')
		},
	},
	{
		name: 'table with built-in style and all style flags',
		fn: async () => {
			const { TableStyle } = await import('../dist/index.js')
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
						tableStyle: TableStyle.MEDIUM_STYLE_2_ACCENT_1,
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
					wholeTbl: { border: { type: 'solid', color: 'D9D9D9', width: 0.5 } },
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
							{ type: 'solid', color: '000000', width: 2 },
							{ type: 'none' },
							{ type: 'dash', color: '888888', width: 1 },
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
										{ type: 'solid', color: '000000', width: 2, cap: 'round' },
										{ type: 'solid', color: '000000', width: 2, cap: 'square' },
										{ type: 'none', cap: 'round' },
										{ type: 'solid', color: '000000', width: 2 },
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
					{ type: 'solid', color: 'FF0000', width: 2 },
					{ type: 'solid', color: 'FF0000', width: 2 },
					{ type: 'solid', color: 'FF0000', width: 2 },
					{ type: 'solid', color: 'FF0000', width: 2 },
				]
				const blue = [
					{ type: 'solid', color: '0000FF', width: 2 },
					{ type: 'solid', color: '0000FF', width: 2 },
					{ type: 'solid', color: '0000FF', width: 2 },
					{ type: 'solid', color: '0000FF', width: 2 },
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
		// Author-side embedded fonts (Feature B): pptx.embedFont() emits raw .fntdata parts,
		// an `application/x-fontdata` Default, presentation font rels, and a p:embeddedFontLst
		// at CT_Presentation index 7. Validate the whole package against the oracle structure
		// (verbatim list from embedded-fonts.oracle.json) and the OpenXmlValidator.
		name: 'author-side embedded fonts (regular + bold)',
		fn: async () => {
			const reg = await readFile(path.join(fontsDir, 'Silkscreen-Regular.ttf'))
			const bold = await readFile(path.join(fontsDir, 'Silkscreen-Bold.ttf'))
			const p = new PptxGenJS()
			await p.embedFont({ data: new Uint8Array(reg), typeface: 'Silkscreen' })
			await p.embedFont({ data: new Uint8Array(bold), typeface: 'Silkscreen', style: 'bold' })
			p.addSlide().addText('Silkscreen', { x: 1, y: 1, w: 8, h: 1, fontFace: 'Silkscreen', fontSize: 24 })
			const buf = await p.stream()

			const zip = await JSZip.loadAsync(buf)
			const names = listEntries(zip)
			assert(
				names.includes('ppt/fonts/font1.fntdata') && names.includes('ppt/fonts/font2.fntdata'),
				'two .fntdata parts present'
			)
			const ct = await readEntry(zip, '[Content_Types].xml')
			assertIncludes(ct, '<Default Extension="fntdata" ContentType="application/x-fontdata"/>', 'fntdata Default')
			const pres = await readEntry(zip, 'ppt/presentation.xml')
			assertIncludes(pres, 'embedTrueTypeFonts="1"', 'embedTrueTypeFonts on')
			assertIncludes(pres, 'saveSubsetFonts="0"', 'saveSubsetFonts off (whole faces)')
			// Matches embedded-fonts.oracle.json embeddedFontLstXml (modulo the panose/pitchFamily/
			// charset PowerPoint inferred — the authoring API declares only typeface in v1).
			assertIncludes(
				pres,
				'<p:embeddedFontLst><p:embeddedFont><p:font typeface="Silkscreen"/><p:regular r:id="rId8"/><p:bold r:id="rId9"/></p:embeddedFont></p:embeddedFontLst>',
				'embeddedFontLst entry'
			)
			await expectNoSchemaErrors(buf, 'embedded-fonts')
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
		// upstream-issue-1258: an image targeting a slide-master/layout picture placeholder must
		// inherit the placeholder's position/size when no explicit w/h are supplied, instead of
		// collapsing to the image's natural (here 1px) size. Asserts the package is schema-valid
		// and that the slide picture's <a:ext> matches the placeholder geometry (4x3in in EMU).
		name: 'image inherits geometry from a master picture placeholder',
		fn: async () => {
			const b64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { buf, zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'PIC_MASTER',
					objects: [{ placeholder: { options: { name: 'picph', type: 'pic', x: 1, y: 1, w: 4, h: 3 }, text: '' } }],
				})
				const slide = p.addSlide({ masterTitle: 'PIC_MASTER' })
				// No w/h supplied: geometry must come from the placeholder, not the 1px natural size.
				slide.addImage({ placeholder: 'picph', data: 'image/png;base64,' + b64 })
			})
			await expectNoSchemaErrors(buf, 'image-master-placeholder-geometry')
			const slideXml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const picBlock = firstXmlBlock(slideXml, 'p:pic', 'slide picture')
			// 4in x 3in @ 914400 EMU/in
			assertIncludes(picBlock, 'cx="3657600" cy="2743200"', 'inherited placeholder ext')
		},
	},
	{
		// upstream-pr-1247 / upstream-issue-1208: a master/layout placeholder authored with a
		// vertical anchor (valign) and/or text insets (margin) must emit those in its <a:bodyPr>,
		// not silently fall back to the default. Before the fix, genXmlBodyProperties applied
		// _bodyProp only to ordinary text objects, so placeholders lost their margin/valign and a
		// slide inserted from the layout did not inherit them. Oracle: layout-placeholder-bodypr.pptx
		// (PowerPoint-authored) — title bottom-anchored 18/9pt insets, body middle-anchored 24/15/12/6pt.
		name: 'master/layout placeholder carries bodyPr insets + anchor',
		fn: async () => {
			const { buf, zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'BODYPR_MASTER',
					objects: [
						// margin is [Top, Right, Bottom, Left] (inches) → tIns/rIns/bIns/lIns; valign → anchor.
						{
							placeholder: {
								options: {
									name: 'title-ph',
									type: 'title',
									x: 0.5,
									y: 0.3,
									w: 9,
									h: 1.2,
									valign: 'bottom',
									margin: [0.1, 0.25, 0.1, 0.25],
								},
								text: '',
							},
						},
						{
							placeholder: {
								options: {
									name: 'body-ph',
									type: 'body',
									idx: 1,
									x: 0.5,
									y: 1.8,
									w: 9,
									h: 4,
									valign: 'middle',
									margin: [0.2, 0.15, 0.05, 0.3],
								},
								text: '',
							},
						},
					],
				})
				p.addSlide({ masterTitle: 'BODYPR_MASTER' })
			})
			await expectNoSchemaErrors(buf, 'master-placeholder-bodypr')
			// The defineSlideMaster placeholders are emitted on the master's layout part; find it.
			const layoutNames = listEntries(zip).filter((n) => /ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(n))
			const layoutXmls = await Promise.all(layoutNames.map((n) => readEntry(zip, n)))
			const layoutXml = layoutXmls.find((xml) => xml.includes('anchor="b"') && xml.includes('lIns="228600"'))
			assert(layoutXml, `found the BODYPR_MASTER layout part among ${layoutNames.join(', ')}`)
			// Title placeholder: bottom anchor, 0.25in L/R + 0.1in T/B insets (EMU @ 914400/in).
			assertIncludes(
				layoutXml,
				'lIns="228600" tIns="91440" rIns="228600" bIns="91440" rtlCol="0" anchor="b"',
				'title placeholder bodyPr'
			)
			// Body placeholder: center anchor, asymmetric 0.3/0.2/0.15/0.05in insets.
			assertIncludes(
				layoutXml,
				'lIns="274320" tIns="182880" rIns="137160" bIns="45720" rtlCol="0" anchor="ctr"',
				'body placeholder bodyPr'
			)
		},
	},
	{
		// upstream-pr-1151: a table can bind to a layout/master content placeholder via the new
		// `placeholder` table option. The table's <p:graphicFrame> then emits the placeholder's
		// <p:ph> on its <p:nvPr> (before <p:extLst>) and inherits the placeholder geometry for any
		// omitted x/y/w/h. Oracle: table-placeholder.pptx (PowerPoint binds AddTable into a content
		// placeholder, emitting <p:ph idx="1"/> on the graphicFrame nvPr).
		name: 'table bound to a layout placeholder emits p:ph on the graphicFrame',
		fn: async () => {
			const { buf, zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'TBL_MASTER',
					objects: [
						{
							placeholder: { options: { name: 'content', type: 'body', idx: 1, x: 0.5, y: 1.5, w: 9, h: 4 }, text: '' },
						},
					],
				})
				const slide = p.addSlide({ masterTitle: 'TBL_MASTER' })
				// No x/y/w/h: geometry must come from the placeholder.
				slide.addTable(
					[
						['A1', 'B1'],
						['A2', 'B2'],
					],
					{ placeholder: 'content' }
				)
			})
			await expectNoSchemaErrors(buf, 'table-placeholder')
			const slideXml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const frame = firstXmlBlock(slideXml, 'p:graphicFrame', 'table graphicFrame')
			const nvPr = firstXmlBlock(frame, 'p:nvPr', 'graphicFrame nvPr')
			// The graphicFrame fills the placeholder: it carries a <p:ph> binding (idx + body type).
			assertIncludes(nvPr, '<p:ph', 'graphicFrame placeholder binding')
			assertIncludes(nvPr, 'type="body"', 'placeholder body type')
			// The <p:ph> precedes <p:extLst> per CT_ApplicationNonVisualDrawingProps document order.
			assert(nvPr.indexOf('<p:ph') < nvPr.indexOf('<p:extLst>'), 'p:ph precedes p:extLst in nvPr')
			// Geometry inherited from the placeholder (9in x 4in @ 914400 EMU/in).
			const xfrm = firstXmlBlock(frame, 'p:xfrm', 'table xfrm')
			assertIncludes(xfrm, 'cx="8229600" cy="3657600"', 'inherited placeholder ext')
		},
	},
	{
		// upstream-issue-446: the notes print layout slide-image placeholder. The notesMaster
		// sldImg placeholder must carry its geometry (off/ext + 1pt black border) and the
		// notesSlide must carry a bare <p:ph type="sldImg"/> that inherits it, so the slide image
		// renders in notes print view. Oracle: notes-slide-image.pptx (PowerPoint-authored) — the
		// current writer output is byte-identical to it; this fixture locks that against regression.
		name: 'notes sldImg placeholder geometry (notesMaster) + bare placeholder (notesSlide)',
		fn: async () => {
			const { buf, zip } = await build((p) => {
				const slide = p.addSlide()
				slide.addText('Body', { x: 1, y: 1, w: 4, h: 1 })
				slide.addNotes('Speaker notes here')
			})
			await expectNoSchemaErrors(buf, 'notes-sldimg-placeholder')
			// notesMaster: the sldImg placeholder carries the print-layout geometry + black border.
			const masterXml = await readEntry(zip, 'ppt/notesMasters/notesMaster1.xml')
			assertIncludes(masterXml, '<p:ph type="sldImg" idx="2"/>', 'notesMaster sldImg placeholder')
			assertIncludes(
				masterXml,
				'<a:off x="685800" y="1143000"/><a:ext cx="5486400" cy="3086100"/>',
				'notesMaster sldImg geometry'
			)
			assertIncludes(
				masterXml,
				'<a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln>',
				'notesMaster sldImg 1pt black border'
			)
			// notesSlide: a bare sldImg placeholder (empty spPr) that inherits the master geometry.
			const slideXml = await readEntry(zip, 'ppt/notesSlides/notesSlide1.xml')
			assertIncludes(
				slideXml,
				'<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/>',
				'notesSlide bare sldImg placeholder'
			)
		},
	},
	{
		// upstream-pr-727: a bar/column chart with per-point fill colours AND per-point custom
		// data-label text, kept consistent with the embedded workbook value cache. The series-level
		// `pointStyles[].fill` (per-point c:dPt) and `customLabels[]` (per-point rich c:dLbl) APIs
		// cover this together. Oracle: bar-chart-data-labels.pptx (CT_BarSer with 4 recoloured bars
		// FF0000/00B050/0070C0/FFC000 + custom labels Low/Mid/High/Peak over numCache 10/25/18/30).
		name: 'bar chart per-point colours + custom data labels (consistent with value cache)',
		fn: async () => {
			const { buf, zip } = await build((p) => {
				p.addSlide().addChart(
					p.ChartType.bar,
					[
						{
							name: 'Revenue',
							labels: [['Q1', 'Q2', 'Q3', 'Q4']],
							values: [10, 25, 18, 30],
							pointStyles: [{ fill: 'FF0000' }, { fill: '00B050' }, { fill: '0070C0' }, { fill: 'FFC000' }],
							customLabels: ['Low', 'Mid', 'High', 'Peak'],
						},
					],
					{ barDir: 'col', showValue: true }
				)
			})
			await expectNoSchemaErrors(buf, 'bar-per-point-labels-colors')
			const chartXml = await readEntry(zip, 'ppt/charts/chart1.xml')
			const ser = firstXmlBlock(chartXml, 'c:ser', 'bar series')
			// Per-point fills: one <c:dPt> per recoloured bar.
			for (const hex of ['FF0000', '00B050', '0070C0', 'FFC000']) {
				assertIncludes(ser, `<a:srgbClr val="${hex}"/>`, `dPt fill ${hex}`)
			}
			// Per-point custom label text in rich <c:dLbl> runs.
			for (const text of ['Low', 'Mid', 'High', 'Peak']) {
				assertIncludes(ser, `<a:t>${text}</a:t>`, `dLbl text ${text}`)
			}
			// The value cache still holds the real numbers (labels override display, not data).
			const val = firstXmlBlock(ser, 'c:val', 'series values')
			for (const num of ['10', '25', '18', '30']) {
				assertIncludes(val, `<c:v>${num}</c:v>`, `numCache value ${num}`)
			}
			// CT_BarSer document order: dPt* → dLbls → cat → val.
			assert(ser.indexOf('<c:dPt>') < ser.indexOf('<c:dLbls>'), 'c:dPt precedes c:dLbls')
			assert(ser.indexOf('<c:dLbls>') < ser.indexOf('<c:cat>'), 'c:dLbls precedes c:cat')
		},
	},
	{
		// upstream-issue-1456: a native, editable PowerPoint equation (OMML) in a text box. A text
		// item's `math` raw-OMML property emits a display-math paragraph (<a14:m><m:oMathPara><m:oMath>)
		// and the whole shape is wrapped in <mc:AlternateContent><mc:Choice Requires="a14"> so non-a14
		// consumers + validators treat the a14 subtree as a known extension. Oracle: math-omml.pptx
		// (PowerPoint-authored x^2+1=y), which validates clean with the same envelope.
		name: 'native math equation (OMML) text run',
		fn: async () => {
			const omml =
				'<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup><m:r><m:t>+1=y</m:t></m:r>'
			const { buf, zip } = await build((p) => {
				p.addSlide().addText([{ math: omml }], { x: 1, y: 2, w: 8, h: 1 })
			})
			await expectNoSchemaErrors(buf, 'native-math-omml')
			const slideXml = await readEntry(zip, 'ppt/slides/slide1.xml')
			// The equation shape is wrapped in the a14 markup-compatibility envelope.
			assertIncludes(
				slideXml,
				'<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">',
				'a14 mc:Choice envelope'
			)
			// The math paragraph: a14:m → m:oMathPara → m:oMath carrying the supplied OMML.
			const ac = firstXmlBlock(slideXml, 'mc:AlternateContent', 'math AlternateContent')
			assertIncludes(ac, '<a14:m', 'a14:m equation marker')
			assertIncludes(
				ac,
				'<m:oMathPara><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr><m:oMath><m:sSup>',
				'oMathPara/oMath wrapping the OMML'
			)
			assertIncludes(ac, '<m:t>+1=y</m:t>', 'the supplied OMML run is present')
			// The m namespace is declared so the m: prefix resolves.
			assertIncludes(
				ac,
				'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
				'math namespace declared'
			)
		},
	},
	{
		// upstream-issue-1456 (LaTeX leg): the `@shbernal/pptxgenjs/math` subpath converts LaTeX to
		// OMML via latexToOmml() (LaTeX --temml--> MathML --mathml2omml--> OMML). Feeding that OMML to
		// the `math:` option must produce the same schema-valid a14 display-math envelope as raw OMML.
		// A representative corpus (fraction, radical, sum/int limits, matrix, cases, greek, accents,
		// fences) goes one-per-slide; the deck must validate clean.
		name: 'native math equation (LaTeX via /math) text runs',
		fn: async () => {
			const corpus = [
				'x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}',
				'\\sqrt{1+\\sqrt{1+x}}',
				'\\sum_{i=1}^{n} i = \\frac{n(n+1)}{2}',
				'\\int_0^\\infty e^{-x}\\,dx = 1',
				'\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}',
				'f(x) = \\begin{cases} 1 & x>0 \\\\ 0 & x\\le 0 \\end{cases}',
				'\\alpha + \\beta = \\gamma',
				'\\hat{a} + \\bar{b}',
				'\\left( \\frac{a}{b} \\right)',
			]
			const { buf, zip } = await build((p) => {
				for (const latex of corpus) {
					p.addSlide().addText([{ math: latexToOmml(latex) }], { x: 1, y: 2, w: 8, h: 1 })
				}
			})
			await expectNoSchemaErrors(buf, 'native-math-latex')
			// Structural check: the LaTeX path lands in the same envelope as the raw-OMML oracle
			// (math-omml.pptx) — mc:Choice Requires="a14" → a14:m → m:oMathPara → m:oMath.
			const slideXml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertIncludes(
				slideXml,
				'<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">',
				'a14 mc:Choice envelope'
			)
			const ac = firstXmlBlock(slideXml, 'mc:AlternateContent', 'math AlternateContent')
			assertIncludes(ac, '<a14:m', 'a14:m equation marker')
			assertIncludes(
				ac,
				'<m:oMathPara><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr><m:oMath>',
				'oMathPara/oMath wrapping the converted OMML'
			)
			assertIncludes(ac, '<m:f>', 'the fraction converted to an m:f')
		},
	},
	{
		// upstream-issue-1360: defineSlideMaster({ textStyles }) configures the shared slide
		// master's per-level <p:txStyles>. Assert the configured master is schema-valid and that
		// the body level overrides (bullet char, font size, color) landed in slideMaster1.xml.
		name: 'configurable master text styles (txStyles)',
		fn: async () => {
			const { buf, zip } = await build((p) => {
				p.defineSlideMaster({
					title: 'TXSTYLE_MASTER',
					textStyles: {
						title: { fontSize: 40, color: '1F3864', bold: true },
						body: [
							{ fontSize: 24, color: 'C00000', bold: true, bullet: { characterCode: '25AA', fontFace: 'Arial' } },
							{ fontSize: 20, align: 'right', bullet: false },
							{ bullet: { type: 'number', numberType: 'arabicPeriod' } },
						],
					},
				})
				p.addSlide({ masterTitle: 'TXSTYLE_MASTER' }).addText('Body', { x: 1, y: 1, w: 6, h: 1 })
			})
			await expectNoSchemaErrors(buf, 'master-txstyles')
			const masterXml = await readEntry(zip, 'ppt/slideMasters/slideMaster1.xml')
			const txStyles = firstXmlBlock(masterXml, 'p:txStyles', 'master txStyles')
			const titleStyle = firstXmlBlock(txStyles, 'p:titleStyle', 'titleStyle')
			const bodyStyle = firstXmlBlock(txStyles, 'p:bodyStyle', 'bodyStyle')
			assertIncludes(titleStyle, 'sz="4000"', 'title fontSize 40pt')
			assertIncludes(titleStyle, 'b="1"', 'title bold')
			assertIncludes(titleStyle, '<a:srgbClr val="1F3864"/>', 'title color')
			assertIncludes(bodyStyle, 'sz="2400"', 'body lvl1 fontSize 24pt')
			assertIncludes(bodyStyle, '<a:srgbClr val="C00000"/>', 'body lvl1 color')
			assertIncludes(bodyStyle, '<a:buChar char="&#x25AA;"/>', 'body lvl1 custom bullet char')
			assertIncludes(bodyStyle, '<a:lvl2pPr marL="742950" indent="-285750" algn="r"', 'body lvl2 right align')
			assertIncludes(bodyStyle, '<a:buNone/>', 'body lvl2 bullet suppressed')
			assertIncludes(bodyStyle, '<a:buAutoNum type="arabicPeriod"/>', 'body lvl3 auto-number bullet')
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
		name: 'image crop emits explicit srcRect (percentage edge insets)',
		fn: async () => {
			// `crop` maps a sub-region of the source verbatim into the box. Two pictures reference the
			// same composite raster, each keeping a different quadrant — the composite-icon use case.
			const b64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addImage({
					data: 'image/png;base64,' + b64,
					x: 0.5,
					y: 0.5,
					w: 2,
					h: 2,
					crop: { l: 0, t: 0, r: 50, b: 50 }, // top-left quadrant
				})
				s.addImage({
					data: 'image/png;base64,' + b64,
					x: 3,
					y: 0.5,
					w: 2,
					h: 2,
					crop: { l: 50, t: 50 }, // bottom-right quadrant (omitted edges default to 0)
				})
			})
			await expectNoSchemaErrors(buf, 'image-crop-srcrect')
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
					p.ChartType.scatter,
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
					p.ChartType.scatter,
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
				p.addSlide().addChart(p.ChartType.bar, [{ name: 'S1', labels: ['A', 'B', 'C'], values: [0.1, 0.2, 0.3] }], {
					x: 0.5,
					y: 0.5,
					w: 4,
					h: 3,
					showValue: true,
					dataLabelFormatCode: '0%',
				})
				p.addSlide().addChart(p.ChartType.pie, [{ name: 'S1', labels: ['A', 'B', 'C'], values: [0.5, 0.3, 0.2] }], {
					x: 0.5,
					y: 0.5,
					w: 4,
					h: 3,
					showPercent: true,
					dataLabelFormatCode: '0%',
				})
				p.addSlide().addChart(
					p.ChartType.scatter,
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
		// chart-metadata-extlst: custom chart-level metadata rides in the schema-valid extension
		// list on the chart space (CT_ChartSpace/c:extLst, the LAST child) under a stable PptxGenJS
		// vendor GUID, NOT as an invalid c:meta sibling PowerPoint would strip/repair. Each entry is
		// a foreign-namespace <pgm:item key="" value=""/> inside the lax-processed CT_Extension
		// wildcard. Lock in: schema-valid, extLst is last, payload escaped, and invalid entries drop.
		name: 'chart metadata emitted via schema-valid chartSpace extLst',
		fn: async () => {
			const { buf, zip } = await build((p) => {
				p.addSlide().addChart(p.ChartType.bar, [{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }], {
					x: 0.5,
					y: 0.5,
					w: 4,
					h: 3,
					metadata: { sourceId: 'q3-revenue', 'note&tag': 'a<b>"c"' },
				})
			})
			await expectNoSchemaErrors(buf, 'chart-metadata-extlst')
			// _chartCounter is a process-global, so the chart part name is not necessarily chart1.xml.
			const chartPath = listEntries(zip).find((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f))
			const chartXml = await readEntry(zip, chartPath)
			const extLst = firstXmlBlock(chartXml, 'c:extLst', 'chartSpace extLst')
			assertIncludes(extLst, '<c:ext uri="{094A432E-1F6C-499B-95B8-B57DC9536949}">', 'vendor ext uri')
			assertIncludes(extLst, '<pgm:metadata xmlns:pgm="http://pptxgenjs.com/schema/chart/metadata">', 'metadata ns')
			assertIncludes(extLst, '<pgm:item key="sourceId" value="q3-revenue"/>', 'plain entry')
			// Keys and values are XML-escaped (no raw &, <, >, ").
			assertIncludes(extLst, '<pgm:item key="note&amp;tag" value="a&lt;b&gt;&quot;c&quot;"/>', 'escaped entry')
			// extLst is the LAST child of CT_ChartSpace (after externalData).
			assert(chartXml.indexOf('<c:externalData') < chartXml.indexOf('<c:extLst>'), 'externalData precedes extLst')
			assert(chartXml.indexOf('<c:extLst>') < chartXml.indexOf('</c:chartSpace>'), 'extLst before chartSpace close')
		},
	},
	{
		// chart-metadata-extlst: a chart with no metadata (and one with only-invalid entries) emits
		// no extLst at all — the extension is purely opt-in and never produces an empty element.
		name: 'chart without metadata emits no chartSpace extLst',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(p.ChartType.bar, [{ name: 'S1', labels: ['A'], values: [1] }], {
					x: 0.5,
					y: 0.5,
					w: 4,
					h: 3,
				})
			})
			const chartPath = listEntries(zip).find((f) => /^ppt\/charts\/chart\d+\.xml$/.test(f))
			const chartXml = await readEntry(zip, chartPath)
			assert(!chartXml.includes('<c:extLst>'), 'no extLst when metadata absent')
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
					p.ChartType.bubble,
					[
						{ name: 'X-Axis', values: [1, 2, 3, 4] },
						{ name: 'Y-Values 1', values: [13, 20, 21, 25], sizes: [10, 5, 20, 15] },
					],
					{ x: 0.5, y: 0.5, w: 6, h: 3, showBubbleSize: true }
				)
				p.addSlide().addChart(
					p.ChartType.bubble3d,
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
				p.addSlide().addChart(p.ChartType.bar, [{ name: '系列', labels: ['甲', '乙', '丙'], values: [1, 2, 3] }], {
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
		// Upstream #1021: legendLayout emits a <c:manualLayout> inside <c:legend> so the
		// legend can be positioned and sized manually. Schema order inside CT_ManualLayout
		// is xMode, yMode, x, y, w, h; <c:layout> sits between legendEntry and overlay.
		name: 'chart legend with manual layout emits schema-valid manualLayout (upstream #1021)',
		fn: async () => {
			const { buf, zip } = await build((p) => {
				p.addSlide().addChart(p.ChartType.bar, [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
					x: 0.5,
					y: 0.5,
					w: 6,
					h: 3,
					showLegend: true,
					legendPos: 'r',
					legendLayout: { x: 0.7, y: 0.3, w: 0.25, h: 0.4 },
				})
			})
			// _chartCounter is global across the test run, so the chart file number is
			// not deterministic here; locate the single chart part by pattern.
			const chartPath = Object.keys(zip.files).find((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))
			const chartXml = await readEntry(zip, chartPath)
			assertIncludes(
				chartXml,
				'<c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.7"/><c:y val="0.3"/><c:w val="0.25"/><c:h val="0.4"/></c:manualLayout></c:layout>',
				'legend manual layout'
			)
			await expectNoSchemaErrors(buf, 'chart-legend-manual-layout')
		},
	},
	{
		name: 'bar chart with valAxisCrossBetween midCat',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addChart(p.ChartType.bar, [{ name: 'Series 1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
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
					p.ChartType.bar,
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
					p.ChartType.bar,
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
						p.ChartType.bar,
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
					s.addChart(p.ChartType.line, [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
						x: 1,
						y: 1,
						w: 6,
						h: 3,
						lineDataSymbolSize: 1,
					})
					const s2 = p.addSlide()
					s2.addChart(p.ChartType.line, [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
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
						p.ChartType.bar,
						[
							{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] },
							{ name: 'S2', labels: ['A', 'B', 'C'], values: [2, 3, 1] },
						],
						{ x: 1, y: 1, w: 6, h: 3, barGapWidthPct: 600, barOverlapPct: 200 }
					)
					// holeSize 200 (>90) violates ST_HoleSize; firstSliceAng 400 (>360) violates ST_FirstSliceAng.
					p.addSlide().addChart(p.ChartType.doughnut, [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
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
					s.addShape(p.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, fill: { color: 'FF0000', transparency: 150 } })
					s.addShape(p.ShapeType.rect, { x: 4, y: 1, w: 2, h: 1, fill: { color: '00FF00', transparency: -20 } })
					// line width 2000pt -> w 25.4M EMU (>20116800), and a negative width, both violate ST_LineWidth.
					s.addShape(p.ShapeType.rect, { x: 1, y: 3, w: 2, h: 1, line: { color: '0000FF', width: 2000 } })
					s.addShape(p.ShapeType.rect, { x: 4, y: 3, w: 2, h: 1, line: { color: '0000FF', width: -5 } })
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
				s.addChart(p.ChartType.bar, [{ name: 'Series 1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
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
				s.addChart(p.ChartType.bar, [{ name: 'Series 1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
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
		name: 'picture bullet (buBlip image)',
		fn: async () => {
			const b64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
			const { buf } = await build((p) => {
				const s = p.addSlide()
				// shape-level picture bullet shared across runs
				s.addText('star item', {
					x: 1,
					y: 1,
					w: 4,
					h: 0.5,
					bullet: { image: { data: 'image/png;base64,' + b64 }, size: 120 },
				})
				// second box re-using the same data must register its own slide rel
				s.addText('another item', {
					x: 1,
					y: 2,
					w: 4,
					h: 0.5,
					bullet: { image: { data: 'image/png;base64,' + b64 } },
				})
			})
			await expectNoSchemaErrors(buf, 'picture-bullet')
		},
	},
	{
		name: 'SVG picture bullet (buBlip + svgBlip ext)',
		fn: async () => {
			const svg =
				'image/svg+xml;base64,' +
				Buffer.from(
					'<svg xmlns="http://www.w3.org/2000/svg" width="8" height="8"><circle cx="4" cy="4" r="4"/></svg>'
				).toString('base64')
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addText('svg bullet', { x: 1, y: 1, w: 4, h: 0.5, bullet: { image: { data: svg }, size: 120 } })
				// second box re-using the same SVG data must register its own dual rel pair
				s.addText('another svg', { x: 1, y: 2, w: 4, h: 0.5, bullet: { image: { data: svg } } })
			})
			await expectNoSchemaErrors(buf, 'svg-picture-bullet')
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
				p.addSlide().addChart(p.ChartType.line, [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
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
				p.addSlide().addChart(
					p.ChartType.line,
					[{ name: 'S1', labels: ['A', 'B', 'C', 'D'], values: [1, null, 3, 4] }],
					{
						x: 1,
						y: 1,
						w: 6,
						h: 3,
					}
				)
			})
			await expectNoSchemaErrors(buf, 'line-chart-null-values-gap')
		},
	},
	{
		name: 'line chart with per-series lineDashValues',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.ChartType.line,
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
							type: p.ChartType.bar,
							data: [{ name: 'Revenue', labels: ['Q1', 'Q2', 'Q3'], values: [10, 20, 30] }],
							options: {},
						},
						{
							type: p.ChartType.line,
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
							type: p.ChartType.bar,
							data: [{ name: 'Bottom', labels, values: [17, 26, 53, 10, 4] }],
							options: { barDir: 'bar', barGrouping: 'clustered' },
						},
						{
							type: p.ChartType.scatter,
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
					p.ChartType.bar,
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
						dataLabelFormatCode: '#,##0',
						seriesOptions: [
							{ color: 'FF0000', dataLabelColor: 'FFFFFF', dataLabelFontBold: true, dataLabelFormatCode: '0.00%' },
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
					p.ChartType.line,
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
		name: 'chart error bars (bar percentage/cust, line stdDev, scatter x+y)',
		fn: async () => {
			const { buf } = await build((p) => {
				// BAR: percentage error bars on one series, custom per-point on another
				p.addSlide().addChart(
					p.ChartType.bar,
					[
						{
							name: 'Pct',
							labels: ['Q1', 'Q2', 'Q3'],
							values: [10, 20, 30],
							errorBars: { valueType: 'percentage', value: 5, color: 'FF0000', size: 1 },
						},
						{
							name: 'Cust',
							labels: ['Q1', 'Q2', 'Q3'],
							values: [15, 25, 5],
							errorBars: { valueType: 'cust', plusValues: [1, 2, 1], minusValues: [0.5, 1, 0.5], noEndCap: true },
						},
					],
					{ x: 1, y: 1, w: 6, h: 3 }
				)
				// LINE: standard-deviation error bars, plus-only
				p.addSlide().addChart(
					p.ChartType.line,
					[
						{
							name: 'StdDev',
							labels: ['Jan', 'Feb', 'Mar'],
							values: [1, 2, 3],
							errorBars: { valueType: 'stdDev', value: 1, barType: 'plus' },
						},
					],
					{ x: 1, y: 1, w: 6, h: 3 }
				)
				// SCATTER: both X and Y error bars on the Y series
				p.addSlide().addChart(
					p.ChartType.scatter,
					[
						{ name: 'X-Axis', values: [1, 2, 3, 4] },
						{
							name: 'Y-Value',
							values: [13, 20, 21, 25],
							errorBars: [
								{ direction: 'x', valueType: 'fixedVal', value: 0.5 },
								{ direction: 'y', valueType: 'stdErr' },
							],
						},
					],
					{ x: 1, y: 1, w: 6, h: 3 }
				)
			})
			await expectNoSchemaErrors(buf, 'chart-error-bars')
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
				s.addShape(p.ShapeType.round2SameRect, {
					x: 0.5,
					y: 0.5,
					w: 3,
					h: 2,
					rectRadius: 0.1,
					fill: { color: '4472C4' },
				})
				s.addShape(p.ShapeType.round2DiagRect, {
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
				s.addShape(p.ShapeType.chevron, {
					x: 0.5,
					y: 0.5,
					w: 3,
					h: 1,
					shapeAdjust: { name: 'adj', value: 0.25 },
					fill: { color: '4472C4' },
				})
				// Array form on a rounded-rectangle adjust handle
				s.addShape(p.ShapeType.roundRect, {
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
					p.ChartType.bar,
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
				p.addSlide({ masterTitle: 'ROUNDRECT_MASTER' })
			})
			await expectNoSchemaErrors(buf, 'slide-master-roundrect')
		},
	},
	{
		name: 'bar chart with per-point customLabels',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.ChartType.bar,
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
					p.ChartType.pie,
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
					p.ChartType.bar,
					[
						{
							name: 'Status',
							labels: ['A', 'B', 'C', 'D'],
							values: [10, 20, 38, 2],
							pointStyles: [
								{ border: { width: 2, color: 'FF0000' } },
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
		name: 'bar chart with per-point pattern fills (a:pattFill)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.ChartType.bar,
					[
						{
							name: 'Status',
							labels: ['A', 'B', 'C', 'D'],
							values: [10, 20, 38, 2],
							pointStyles: [
								// fgColor defaults to the resolved point fill -> hatched bar color
								{ fill: '00B050', pattern: { preset: 'ltUpDiag' } },
								// explicit fg/bg colors
								{ pattern: { preset: 'diagCross', fgColor: 'C00000', bgColor: 'FFFFFF' } },
								// pattern alongside a border
								{ pattern: { preset: 'pct25' }, border: { width: 2, color: '404040' } },
								{},
							],
						},
					],
					{ x: 1, y: 1, w: 6, h: 3 }
				)
			})
			await expectNoSchemaErrors(buf, 'bar-chart-pattern-fills')
		},
	},
	{
		name: 'pie chart with per-point pointStyles (border + fill)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					p.ChartType.pie,
					[
						{
							name: 'Status',
							labels: ['Red', 'Amber', 'Green'],
							values: [10, 30, 60],
							pointStyles: [{ border: { width: 3, color: 'C00000' } }, {}, { fill: '70AD47' }],
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
					p.ChartType.pie,
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
				p.addSlide().addChart(p.ChartType.doughnut, [{ name: 'Status', labels: ['A', 'B', 'C'], values: [1, 2, 3] }], {
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
					p.ChartType.line,
					[
						{
							name: 'Series 1',
							labels: ['A', 'B', 'C', 'D'],
							values: [4, 8, 6, 10],
							pointStyles: [
								{},
								{ border: { width: 2, color: 'FF0000' } },
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
					p.ChartType.area,
					[
						{
							name: 'Series 1',
							labels: ['A', 'B', 'C'],
							values: [5, 9, 7],
							pointStyles: [{ fill: 'FFC000' }, {}, { border: { width: 1, color: '404040' } }],
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
					p.ChartType.scatter,
					[
						{ name: 'X-Axis', values: [1, 2, 3, 4] },
						{
							name: 'Y-Values',
							values: [3, 6, 2, 8],
							pointStyles: [{ border: { width: 2, color: 'FF0000' } }, {}, { fill: '00B050' }, {}],
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
		// upstream #1288: theme East Asian (<a:ea>) and complex-script (<a:cs>) font slots must stay
		// schema-valid when populated from ThemeProps for both the major and minor fonts.
		name: 'theme East Asian / complex-script font faces',
		fn: async () => {
			const { buf } = await build((p) => {
				p.theme = {
					headFontFace: 'Arial Narrow',
					bodyFontFace: 'Arial',
					headFontFaceEA: 'Yu Gothic',
					bodyFontFaceEA: 'Yu Mincho',
					headFontFaceCS: 'Arial',
					bodyFontFaceCS: 'Times New Roman',
				}
				p.addSlide().addText('テーマ', { x: 1, y: 1, w: 4, h: 0.5 })
			})
			await expectNoSchemaErrors(buf, 'theme-ea-cs-fonts')
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
				// Shape binding: <a:stCxn>/<a:endCxn> in <p:cNvCxnSpPr> must stay schema-valid.
				s.addShape('rect', { x: 0.5, y: 6.5, w: 1, h: 0.5, objectName: 'cxnBoxA' })
				s.addShape('rect', { x: 4, y: 6.5, w: 1, h: 0.5, objectName: 'cxnBoxB' })
				s.addConnector({
					type: 'elbow',
					x1: 1.5,
					y1: 6.75,
					x2: 4,
					y2: 6.75,
					startShape: 'cxnBoxA',
					startShapeIdx: 3,
					endShape: 'cxnBoxB',
					endShapeIdx: 1,
				})
			})
			await expectNoSchemaErrors(buf, 'connectors')
		},
	},
	{
		// Looping media (upstream-issue-1434): `loop`/`loopCount` emit a slide-level
		// <p:timing> tree with repeatCount on the media node's <p:cTn>. Asserts the
		// timing tree (tmRoot + p:video/cMediaNode) stays schema-valid.
		name: 'media loop and loopCount (p:timing repeatCount)',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addMedia({ type: 'video', data: 'video/mp4;base64,AAAA', x: 1, y: 1, w: 3, h: 2, loop: true })
				s.addMedia({ type: 'video', data: 'video/mp4;base64,BBBB', x: 5, y: 1, w: 3, h: 2, loopCount: 3 })
				// audio loops via <a:audioFile> + <p:audio> timing node
				s.addMedia({ type: 'audio', data: 'audio/mp3;base64,CCCC', x: 1, y: 4, w: 3, h: 2, loop: true })
			})
			await expectNoSchemaErrors(buf, 'media-loop')
		},
	},
	{
		// Slide transitions (docs/animations-and-transitions.md, Phase 1): p:transition
		// is emitted between p:clrMapOvr and p:timing. Bare form for a speed bucket;
		// mc:AlternateContent (p14 Choice with p14:dur + base Fallback) for an exact
		// duration. Asserts both forms — and their type-variant attrs — stay schema-valid.
		name: 'slide transitions (p:transition bare + mc:AlternateContent)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().transition = { type: 'fade' }
				p.addSlide().transition = { type: 'push', durationMs: 1250, variant: { dir: 'd' } }
				p.addSlide().transition = { type: 'wipe', speed: 'med', variant: { dir: 'u' } }
				p.addSlide().transition = { type: 'dissolve', durationMs: 2000, speed: 'slow' }
				p.addSlide().transition = { type: 'fade', speed: 'med', advanceOnClick: false, advanceAfterMs: 3000 }
			})
			await expectNoSchemaErrors(buf, 'slide-transitions')
		},
	},
	{
		// Preset build animations (docs/animations-and-transitions.md, Phase 1): the
		// p:timing mainSeq is assembled from captured preset templates, grouped into
		// click steps by trigger, with one p:bldP per animated shape. Mirrors the rich
		// fixture (entrance/emphasis/exit x click/after/with). Asserts schema validity.
		name: 'preset build animations (p:timing mainSeq + p:bldLst)',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				for (const nm of ['fadeShape', 'flyShape', 'growShape', 'exitShape'])
					s.addText(nm, { x: 1, y: 1, w: 3, h: 1, objectName: nm })
				s.addAnimation({ preset: 'fadeIn', shapeIndex: 0, trigger: 'onClick' })
				s.addAnimation({ preset: 'flyIn', shapeIndex: 1, trigger: 'afterPrevious' })
				s.addAnimation({ preset: 'grow', shapeIndex: 2, trigger: 'withPrevious' })
				s.addAnimation({ preset: 'fadeOut', objectName: 'exitShape', trigger: 'onClick' })
			})
			await expectNoSchemaErrors(buf, 'preset-animations')
		},
	},
	{
		// Phase 2 capability B (docs/animations-and-transitions.md): the expanded
		// preset set adds appear/wipe (entr), spin (emph), and flyOut (exit) on top
		// of the Phase 1 four. Emits one on-click effect per preset across all three
		// classes (mirrors slide-animation-presets.pptx). Asserts the timing tree +
		// p:bldLst for the new templates stays schema-valid.
		name: 'expanded preset build animations (appear/wipe/spin/flyOut)',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				const presets = ['appear', 'wipe', 'spin', 'flyOut']
				presets.forEach((preset, i) => {
					const nm = `shape-${preset}`
					s.addText(nm, { x: 1, y: 1 + i, w: 3, h: 1, objectName: nm })
					s.addAnimation({ preset, shapeIndex: i, trigger: 'onClick' })
				})
			})
			await expectNoSchemaErrors(buf, 'expanded-preset-animations')
		},
	},
	{
		// Phase 2 capability C (docs/animations-and-transitions.md): transition sounds.
		// A start sound (p:sndAc/p:stSnd/p:snd r:embed) pulls in an audio relationship,
		// an embedded WAV media part, and a wav=audio/x-wav Default content type; the
		// looped form adds @loop and the stop-previous form is a bare p:endSnd (no rel/
		// part). Mirrors slide-transition-sound.pptx. Asserts the whole package — slide
		// XML, slide rels, and content types — stays schema-valid.
		name: 'slide transition sounds (p:sndAc start/loop/stop + audio rel graph)',
		fn: async () => {
			// Minimal valid 16-bit PCM mono WAV (silent), enough to embed a real audio part.
			const wav = 'data:audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEAQB8AAIA+AAACABAAZGF0YQAAAAA='
			const { buf } = await build((p) => {
				p.addSlide().transition = { type: 'fade', durationMs: 2000, sound: { data: wav, name: 'ding.wav' } }
				p.addSlide().transition = { type: 'fade', durationMs: 2000, sound: { data: wav, name: 'ding.wav', loop: true } }
				p.addSlide().transition = { type: 'fade', durationMs: 2000, sound: { stopPrevious: true } }
			})
			await expectNoSchemaErrors(buf, 'transition-sounds')
		},
	},
	{
		// Speaker-notes hyperlinks + rich runs (upstream-issue-1250): notes runs carry
		// inline formatting and external `url` hyperlinks. The hyperlink emits an
		// <a:hlinkClick> in the notes body and an external relationship in the notes
		// part's rels (rId3+). Asserts the notes part + its rels stay schema-valid.
		name: 'speaker notes with hyperlink and rich runs',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello', { x: 1, y: 1, w: 4, h: 0.5 })
				s.addNotes([
					{ text: 'Intro. ' },
					{
						text: 'bold link',
						options: { bold: true, hyperlink: { url: 'https://gitbrent.github.io/PptxGenJS/', tooltip: 'Docs' } },
					},
					{ text: '\nNext line ' },
					{ text: 'red', options: { color: 'FF0000', italic: true } },
				])
			})
			await expectNoSchemaErrors(buf, 'notes-hyperlinks')
		},
	},
	{
		// upstream-issue-1301: a custom `fontFace` fills the Latin (<a:latin>) + complex-script (<a:cs>)
		// slots only, and `fontFaceEA` adds an explicit East Asian (<a:ea>) face. Lock in that the
		// resulting run properties stay schema-valid (correct CT_TextCharacterProperties child order).
		name: 'custom fontFace + fontFaceEA emit schema-valid latin/ea/cs runs (upstream #1301)',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addText('Latin only', { x: 1, y: 1, w: 4, h: 0.5, fontFace: 'Jost Light' })
				s.addText('東アジア', { x: 1, y: 2, w: 4, h: 0.5, fontFace: 'Jost Light', fontFaceEA: '游ゴシック' })
			})
			await expectNoSchemaErrors(buf, 'text-fontface-ea-cs')
		},
	},
	{
		// upstream-issue-1165: a hyperlink run with no color inherits the theme hyperlink
		// color, so it must emit a bare <a:hlinkClick/> (no solidFill, no hlinkClr override);
		// a hyperlink with an explicit color emits solidFill + ahyp:hlinkClr. Lock in that both
		// the theme-colored and explicitly-colored hyperlink runs stay schema-valid.
		name: 'slide hyperlink runs stay schema-valid with and without color (upstream #1165)',
		fn: async () => {
			const { buf } = await build((p) => {
				const s = p.addSlide()
				s.addText('theme link', {
					x: 1,
					y: 1,
					w: 4,
					h: 0.5,
					hyperlink: { url: 'https://example.com', tooltip: 'Example' },
				})
				s.addText('red link', { x: 1, y: 2, w: 4, h: 0.5, color: 'FF0000', hyperlink: { url: 'https://example.com' } })
				s.addText('jump', { x: 1, y: 3, w: 4, h: 0.5, hyperlink: { slide: 1 } })
			})
			await expectNoSchemaErrors(buf, 'slide-hyperlink-theme-colors')
		},
	},
	{
		// upstream-issue-307: a flat group (<p:grpSp>) wrapping a shape, a text box, and an image.
		// Identity child coordinate space; children keep their slide-absolute coordinates.
		name: 'flat group of shape + text + image (addGroup)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addGroup(
					[
						{ rect: { x: 1, y: 1, w: 2, h: 1, fill: { color: 'CC0000' } } },
						{ text: { text: 'Grouped', options: { x: 1.2, y: 1.2, w: 1.6, h: 0.6, color: 'FFFFFF' } } },
						{
							image: {
								x: 3.5,
								y: 1,
								w: 1,
								h: 1,
								data: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
							},
						},
					],
					{ objectName: 'SchemaGroup' }
				)
			})
			await expectNoSchemaErrors(buf, 'flat-group')
		},
	},
	{
		// upstream-issue-307: a nested group (<p:grpSp> inside <p:grpSp>). Identity child coordinate
		// space at every depth; children keep their slide-absolute coordinates.
		name: 'nested group of rect + (group of rect + text) (addGroup)',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addGroup(
					[
						{ rect: { x: 1, y: 1, w: 2, h: 1, fill: { color: 'CC0000' } } },
						{
							group: {
								children: [
									{ rect: { x: 4, y: 1, w: 1, h: 1, fill: { color: '00CC00' } } },
									{ text: { text: 'Nested', options: { x: 4, y: 1, w: 1, h: 1, color: 'FFFFFF' } } },
								],
								options: { objectName: 'InnerGroup' },
							},
						},
					],
					{ objectName: 'OuterGroup' }
				)
			})
			await expectNoSchemaErrors(buf, 'nested-group')
		},
	},
	{
		// upstream-pr-1447: native (legacy ISO/IEC 29500 §13) PowerPoint comments. One author,
		// one comment: assert the comment part, the commentAuthors part, both Content-Types
		// Overrides, and the slide->comments / presentation->commentAuthors relationships.
		name: 'slide comment (single author, single comment)',
		fn: async () => {
			const { buf, zip } = await build((p) => {
				p.addSlide().addComment({
					author: 'Ada Lovelace',
					initials: 'AL',
					text: 'Tighten this headline',
					x: 1,
					y: 0.5,
					date: '2026-06-24T10:00:00Z',
				})
			})
			await expectNoSchemaErrors(buf, 'comment-single')

			const commentXml = await readEntry(zip, 'ppt/comments/comment1.xml')
			assertIncludes(commentXml, '<p:cm authorId="0" dt="2026-06-24T10:00:00Z" idx="1">', 'comment cm attrs')
			assertIncludes(commentXml, '<p:pos x="914400" y="457200"/>', 'comment pos in EMU')
			assertIncludes(commentXml, '<p:text>Tighten this headline</p:text>', 'comment text')

			const authorsXml = await readEntry(zip, 'ppt/commentAuthors.xml')
			assertIncludes(
				authorsXml,
				'<p:cmAuthor id="0" name="Ada Lovelace" initials="AL" lastIdx="1" clrIdx="0"/>',
				'commentAuthor entry'
			)

			const ctXml = await readEntry(zip, '[Content_Types].xml')
			assertIncludes(
				ctXml,
				'<Override PartName="/ppt/comments/comment1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml"/>',
				'comments Override'
			)
			assertIncludes(
				ctXml,
				'<Override PartName="/ppt/commentAuthors.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"/>',
				'commentAuthors Override'
			)

			const slideRels = await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels')
			assertIncludes(
				slideRels,
				'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"',
				'slide->comments rel'
			)

			const presRels = await readEntry(zip, 'ppt/_rels/presentation.xml.rels')
			assertIncludes(
				presRels,
				'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" Target="commentAuthors.xml"',
				'presentation->commentAuthors rel'
			)
		},
	},
	{
		// upstream-pr-1447: two authors across two slides with multiple comments — pins per-author
		// idx numbering (each author counts from 1) and lastIdx on the author entries.
		name: 'slide comments (two authors, per-author idx numbering)',
		fn: async () => {
			const { buf, zip } = await build((p) => {
				p.addSlide()
					.addComment({ author: 'Ada Lovelace', initials: 'AL', text: 'First by Ada' })
					.addComment({ author: 'Alan Turing', initials: 'AT', text: 'First by Alan' })
					.addComment({ author: 'Ada Lovelace', initials: 'AL', text: 'Second by Ada' })
				p.addSlide().addComment({ author: 'Alan Turing', initials: 'AT', text: 'Second by Alan' })
			})
			await expectNoSchemaErrors(buf, 'comment-multi')

			const authorsXml = await readEntry(zip, 'ppt/commentAuthors.xml')
			assertIncludes(
				authorsXml,
				'<p:cmAuthor id="0" name="Ada Lovelace" initials="AL" lastIdx="2" clrIdx="0"/>',
				'author 0 lastIdx=2'
			)
			assertIncludes(
				authorsXml,
				'<p:cmAuthor id="1" name="Alan Turing" initials="AT" lastIdx="2" clrIdx="1"/>',
				'author 1 lastIdx=2'
			)

			const c1 = await readEntry(zip, 'ppt/comments/comment1.xml')
			assertIncludes(
				c1,
				'<p:cm authorId="0" idx="1"><p:pos x="457200" y="457200"/><p:text>First by Ada</p:text></p:cm>',
				'Ada idx=1'
			)
			assertIncludes(
				c1,
				'<p:cm authorId="1" idx="1"><p:pos x="457200" y="457200"/><p:text>First by Alan</p:text></p:cm>',
				'Alan idx=1'
			)
			assertIncludes(
				c1,
				'<p:cm authorId="0" idx="2"><p:pos x="457200" y="457200"/><p:text>Second by Ada</p:text></p:cm>',
				'Ada idx=2'
			)

			const c2 = await readEntry(zip, 'ppt/comments/comment2.xml')
			assertIncludes(
				c2,
				'<p:cm authorId="1" idx="2"><p:pos x="457200" y="457200"/><p:text>Second by Alan</p:text></p:cm>',
				'Alan idx=2 on slide 2'
			)
		},
	},
]
