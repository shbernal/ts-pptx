// Recovering a probe's code from the function object, for `docs/comparison-syntax.md`.
//
// The page prints the corpus rather than a second hand-written copy of it, which only works
// while the extraction is exact: a body that loses a line, keeps the arrow wrapper, or comes
// out with its indentation flattened is a documentation page saying something the harness
// did not run. The dedent is the part most likely to rot, because it is the only step that
// looks at more than one line at a time.
//
// The elision threshold matters for a different reason. A snapshot is committed, so a
// constant that is a path has to render relative to the repository root on every machine,
// and a 6 kB data URL has to not render at all.

import path from 'node:path'
import { describe, expect, test } from 'vitest'
import { PROGRAMS, programFrame, programModule, programSource } from '../../scripts/comparison/programs.mjs'
import { PROBES, probeSource, SUBJECTS } from '../../scripts/comparison/probes.mjs'
import { functionBody, literal, renderSource } from '../../scripts/comparison/source.mjs'
import { ROOT } from '../../scripts/script-utils.mjs'

describe('functionBody', () => {
	test('drops the arrow wrapper and the indentation the corpus nests it in', () => {
		const build = (pres) => {
			pres.addSlide()
		}
		expect(functionBody(build)).toBe('pres.addSlide()')
	})

	test('keeps relative depth inside a multi-line body', () => {
		const build = (pres) => {
			pres.addSlide().addText('probe', {
				x: 1,
			})
		}
		expect(functionBody(build)).toBe("pres.addSlide().addText('probe', {\n\tx: 1,\n})")
	})

	test('keeps the await an async arm needs', () => {
		const build = async (pres) => {
			await pres.embedFont({ typeface: 'Silkscreen' })
		}
		expect(functionBody(build)).toBe("await pres.embedFont({ typeface: 'Silkscreen' })")
	})

	// A concise body renders as an expression with no statement around it, and a `function`
	// keyword brings its own `this`. Neither is wrong code; both would print as something the
	// page's stated frame does not wrap, so the corpus is held to one shape.
	test('refuses a shape it cannot print', () => {
		expect(() => functionBody((pres) => pres.addSlide())).toThrow(/block-bodied arrow function/)
	})
})

describe('literal', () => {
	test('elides a string past the threshold and marks the cut', () => {
		const rendered = literal('x'.repeat(200))
		expect(rendered).toBe("'" + 'x'.repeat(60) + "…'")
	})

	test('leaves a string at the threshold whole', () => {
		expect(literal('y'.repeat(60))).toBe("'" + 'y'.repeat(60) + "'")
	})

	// The alternative is a committed snapshot carrying the directory layout of whichever
	// machine measured it, which would also make the page's own diff machine-dependent.
	test('writes a path under the repository root relative to it, with forward slashes', () => {
		expect(literal(path.join(ROOT, 'demos', 'common', 'media', 'cube.glb'))).toBe("'demos/common/media/cube.glb'")
	})

	test('renders the nested shape a chart series has', () => {
		expect(literal([{ name: 'Revenue', values: [12, 19], stacked: false }])).toBe(
			"[{ name: 'Revenue', values: [12, 19], stacked: false }]"
		)
	})

	test('quotes a key that is not an identifier', () => {
		expect(literal({ 'data-uri': 1 })).toBe("{ 'data-uri': 1 }")
	})

	test('has no form for a value the corpus should not be declaring', () => {
		expect(() => literal(() => 1)).toThrow(/no literal form/)
	})
})

describe('renderSource', () => {
	test('declares only the constants the body names', () => {
		const PNG = 'data:image/png;base64,AAA'
		const build = (pres) => {
			pres.addSlide().addImage({ data: PNG })
		}
		expect(renderSource(build, { PNG, UNUSED: 'no' })).toBe(
			"const PNG = 'data:image/png;base64,AAA'\n\npres.addSlide().addImage({ data: PNG })"
		)
	})

	// Declaration order, not order of first use: the syntax page compares two arms' rendered
	// sources to report how many probes call both libraries identically, and a preamble
	// ordered by the body would let that number move on a formatting change.
	test('declares constants in the order the corpus declares them', () => {
		const FIRST = 'a'
		const SECOND = 'b'
		const build = (pres) => {
			pres.addSlide().addImage({ data: SECOND, alt: FIRST })
		}
		const rendered = renderSource(build, { FIRST, SECOND })
		expect(rendered.split('\n').slice(0, 2)).toEqual(["const FIRST = 'a'", "const SECOND = 'b'"])
	})
})

describe('the corpus itself', () => {
	test('every arm renders, and no arm renders empty', () => {
		for (const probe of PROBES)
			for (const subject of SUBJECTS) {
				const source = probeSource(probe, subject)
				if (probe.build[subject] === null) expect(source, `${probe.id}/${subject}`).toBeNull()
				else expect(source, `${probe.id}/${subject}`).toMatch(/\S/)
			}
	})

	// Two arms rendering to the same string is a fact the page prints as a count, so it has to
	// mean the same thing every run. The shared baseline is where it is expected to happen.
	test('the arms that agree are the ones the page counts', () => {
		const agree = PROBES.filter(
			(probe) => probe.build['ts-pptx'] && probeSource(probe, 'ts-pptx') === probeSource(probe, 'pptxgenjs')
		).map((probe) => probe.id)
		expect(agree).not.toContain('bar-chart')
		expect(agree).not.toContain('image')
		expect(agree).toContain('text-run')
	})
})

// The bundle corpus reuses the same extraction for a second purpose, and that purpose has
// the opposite requirement: `programSource` writes for a reader and elides, `programModule`
// is compiled and must not. A data URL truncated on the way into the bundler would be a
// program the library was never given, measured as though it were.
describe('the bundle corpus', () => {
	test('every program builds with both libraries, and no arm renders empty', () => {
		for (const program of PROGRAMS)
			for (const subject of SUBJECTS) {
				expect(program.build[subject], `${program.id}/${subject}`).toBeTypeOf('function')
				expect(programSource(program, subject), `${program.id}/${subject}`).toMatch(/\S/)
			}
	})

	// The id is the bundle table's join key and the entry file's name, so a duplicate would
	// pair one library's chart deck with the other's table deck and overwrite a bundle.
	test('program ids are unique', () => {
		const ids = PROGRAMS.map((program) => program.id)
		expect(new Set(ids).size).toBe(ids.length)
	})

	test('a compiled program carries the frame, the constants and the body', () => {
		const program = PROGRAMS.find((entry) => entry.id === 'full-deck')
		const module = programModule(program, 'ts-pptx')
		expect(module.startsWith("import TsPptx from 'pptx-ts'\n")).toBe(true)
		expect(module).toContain('const pres = new TsPptx()')
		expect(module.trimEnd().endsWith("console.log(await pres.write({ outputType: 'arraybuffer' }))")).toBe(true)
	})

	test('the compiled program keeps a data URL the page elides', () => {
		const program = PROGRAMS.find((entry) => entry.id === 'full-deck')
		expect(programSource(program, 'ts-pptx')).toContain('…')
		expect(programModule(program, 'ts-pptx')).not.toContain('…')
		expect(programModule(program, 'ts-pptx')).toContain('R9awAAAABJRU5ErkJggg==')
	})

	// The page prints the frame with the program cut out of it. Built from the same two
	// strings the module is compiled from, so a reader cannot be shown an import the
	// measurement did not use.
	test('the frame the page prints is the frame that was compiled', () => {
		for (const subject of SUBJECTS) {
			const frame = programFrame(subject)
			const module = programModule(PROGRAMS[0], subject)
			for (const line of frame.split('\n').filter((line) => line.trim() !== '' && !line.startsWith('//')))
				expect(module, subject).toContain(line)
		}
	})
})
