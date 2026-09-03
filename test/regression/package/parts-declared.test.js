// Every part the package writes is declared in `[Content_Types].xml`, and every declaration
// names a part that exists.
//
// The two enumerations are built in different modules from different literals — the assembler
// writes `ppt/slides/slide1.xml`, content types declares `/ppt/slides/slide1.xml` — and a
// disagreement either way is a PowerPoint repair prompt rather than a wrong-looking deck:
// an undeclared part has no content type, and an Override for a part nobody wrote dangles.
// Nothing compared them, so the whole class was invisible to a suite of 3,900 tests.

import { defineRegressionSuite, build, readEntry, listEntries, assert, assertEqual } from '../../helpers.js'
import { ChartType } from '../../../dist/node.js'

const PNG_DATA =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP8z8DwHwAFAAH/Re1ZlAAAAABJRU5ErkJggg=='

/** `Extension` of every `Default`, and `PartName` of every `Override`, in `[Content_Types].xml`. */
function declarations(xml) {
	const attr = (tag, name) => new RegExp(`${name}="([^"]*)"`).exec(tag)?.[1] ?? ''
	const tags = (name) => xml.match(new RegExp(`<${name}\\b[^>]*/>`, 'g')) || []
	return {
		defaults: new Set(tags('Default').map((tag) => attr(tag, 'Extension').toLowerCase())),
		overrides: tags('Override').map((tag) => attr(tag, 'PartName')),
	}
}

/**
 * A deck reaching every part kind the writer emits: more slides than layouts so the two
 * enumerations cannot accidentally line up, comments on one slide only (that part is written
 * per slide that has them, not per slide), a chart with its embedded workbook, and media.
 */
async function everyPartKind() {
	return build((p) => {
		p.defineSlideMaster({ title: 'custom', background: { color: 'F1F1F1' } })
		const first = p.addSlide({ masterTitle: 'custom' })
		first.addImage({ data: PNG_DATA, x: 1, y: 1, w: 1, h: 1 })
		first.addComment({ author: 'Ada Lovelace', text: 'only this slide has one' })
		first.addNotes('speaker notes')
		const second = p.addSlide()
		second.addChart([{ name: 'S', labels: ['a', 'b'], values: [1, 2] }], {
			type: ChartType.bar,
			x: 1,
			y: 1,
			w: 4,
			h: 3,
		})
		const third = p.addSlide()
		third.addMedia({ type: 'video', data: 'video/mp4;base64,AAAA', x: 1, y: 1, w: 4, h: 3 })
	})
}

defineRegressionSuite('package parts and their content-type declarations agree', [
	{
		name: 'every written part is declared, by Override or by its extension Default',
		fn: async () => {
			const { zip } = await everyPartKind()
			const { defaults, overrides } = declarations(await readEntry(zip, '[Content_Types].xml'))
			const declared = new Set(overrides)
			const undeclared = listEntries(zip).filter((name) => {
				// The content-types part declares the others, never itself (ECMA-376 Part 2 §10.1.2).
				if (name === '[Content_Types].xml') return false
				const extension = (name.split('.').pop() ?? '').toLowerCase()
				// An `.xml` part needs an Override of its own. The package does carry a
				// `Default Extension="xml"` -- `application/xml`, what an unrecognised XML part is --
				// so a missing Override is not a missing content type, it is the *generic* one, and
				// PowerPoint offers to repair a deck whose slide announces itself as plain XML.
				// That makes the Default useless as evidence here, which is exactly how an earlier
				// draft of this test passed while a whole enumeration of Overrides was suppressed.
				if (extension === 'xml') return !declared.has(`/${name}`)
				return !defaults.has(extension) && !declared.has(`/${name}`)
			})
			assertEqual(undeclared.length, 0, `parts written but not declared: ${undeclared.join(', ')}`)
		},
	},
	{
		name: 'every Override names a part that was written',
		fn: async () => {
			const { zip } = await everyPartKind()
			const { overrides } = declarations(await readEntry(zip, '[Content_Types].xml'))
			const written = new Set(listEntries(zip).map((name) => `/${name}`))
			const dangling = overrides.filter((partName) => !written.has(partName))
			assertEqual(dangling.length, 0, `Overrides with no part behind them: ${dangling.join(', ')}`)
		},
	},
	{
		name: 'no part is declared twice',
		fn: async () => {
			// A `Default` and an `Override` may both cover a part -- the Override wins -- but two
			// Overrides for one part name is a malformed content-types part, and the shape of the
			// bug that produces one is a second enumeration writing the same name.
			const { zip } = await everyPartKind()
			const { overrides } = declarations(await readEntry(zip, '[Content_Types].xml'))
			const seen = new Set()
			const repeated = overrides.filter((partName) => (seen.has(partName) ? true : (seen.add(partName), false)))
			assertEqual(repeated.length, 0, `part names declared more than once: ${repeated.join(', ')}`)
		},
	},
	{
		name: 'every relationship target resolves to a written part',
		fn: async () => {
			// The third spelling of a part path, and the one with a dialect of its own: a target is
			// relative to the `_rels` folder's parent, so `../slides/slide1.xml` from a notes slide
			// and `slides/slide1.xml` from presentation.xml name the same part.
			const { zip } = await everyPartKind()
			const written = new Set(listEntries(zip))
			const missing = []
			for (const name of listEntries(zip)) {
				if (!name.endsWith('.rels')) continue
				// `ppt/slides/_rels/slide1.xml.rels` -> `ppt/slides/`; the root `_rels/.rels` -> ``.
				const base = name.slice(0, name.lastIndexOf('_rels/'))
				const xml = await readEntry(zip, name)
				for (const tag of xml.match(/<Relationship\b[^>]*\/>/g) || []) {
					if (/TargetMode="External"/.test(tag)) continue
					const target = /Target="([^"]*)"/.exec(tag)?.[1] ?? ''
					// An absolute target is a part name; a relative one resolves against `base`.
					const resolved = target.startsWith('/') ? target.slice(1) : (base + target).replace(/[^/]+\/\.\.\//g, '')
					if (!written.has(resolved)) missing.push(`${name} -> ${target}`)
				}
			}
			assert(missing.length === 0, `relationships pointing at nothing: ${missing.join(', ')}`)
		},
	},
])
