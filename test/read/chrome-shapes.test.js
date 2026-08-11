// Shape accessors on the shared chrome: SlideMaster.shapes / SlideLayout.shapes,
// and the showMasterSp flag that says whether a slide draws them (issue #12).
//
// Before this, `SlideMaster` and `SlideLayout` exposed only `placeholders` — a
// filtered view — so a template's non-placeholder content (the header band, the
// rule under the title, the logo, the wordmark) had no modeled path out of the read
// API at all. `Slide.shapes` and `GroupShape.shapes` already returned `AnyShape[]`;
// the two chrome classes now return the same union from the same `buildShapes`
// dispatch, which is what lets a consumer's shape-walking code apply unchanged.
//
// Two oracles, deliberately separate:
//
//   * **PowerPoint-authored** (`mixed.pptx`, `read-stress.pptx`). This is the arm
//     that matters — the write API puts almost nothing on a layout and *nothing* on
//     a master (`defineSlideMaster` creates a layout under the shared master), so
//     only a genuine desktop-PowerPoint deck exercises master furniture, decorative
//     text on a layout, nested groups in a layout tree, and `schemeClr` fills that
//     have to resolve against the *master's own* colour map rather than a slide's.
//     Every EMU/hex below was read off those fixtures' own XML.
//
//   * **Write→read** (`defineSlideMaster({ objects })`). Every non-`placeholder`
//     member of that union — `rect`, `text`, `line`, `image`, … — writes a shape
//     into the layout part. That arm proves the accessor returns what this library
//     itself authors, and that `placeholders` still filters them out.
//
// `showMasterSp` is `xsd:boolean` defaulting to `true` on both `p:sld` and
// `p:sldLayout` (ECMA-376 attributeGroup AG_ChildSlide). `mixed.pptx` and
// `read-stress.pptx` each carry a genuine `showMasterSp="0"` on their title layout,
// so the layout arm — the one PowerPoint actually writes — has a real oracle. No
// fixture slide sets it, so the slide arm's `false` is asserted through a round trip
// of the attribute written via the `element_` hatch; the `true` default is asserted
// against every fixture slide.

import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { authorRead, schemaErrors, validatorInstalled } from './authored.js'
import { openFixture } from './corpus.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

/** The shapes of `host` that are not placeholders — a template's decorative furniture. */
function decorative(host) {
	return host.shapes.filter((shape) => shape.placeholder === null)
}

describe('SlideMaster.shapes — PowerPoint-authored master furniture (mixed.pptx)', () => {
	test('returns the whole spTree, placeholders included, in document order', async () => {
		const master = (await openFixture('mixed')).masters()[0]

		assertEqual(master.shapes.length, 12, 'the master spTree holds twelve shapes')
		assertEqual(master.placeholders.length, 5, 'five of them carry a p:ph')
		assertEqual(decorative(master).length, 7, 'the other seven are the decoration placeholders never surfaced')

		// Document order: the seven decorative rects precede the five placeholders.
		assertEqual(
			master.shapes.map((shape) => shape.name).join(','),
			'Rectangle 2,Rectangle 3,Rectangle 4,Rectangle 5,Rectangle 6,Rectangle 7,Rectangle 8,' +
				'Rectangle 9,Rectangle 10,Rectangle 11,Rectangle 12,Rectangle 13',
			'shapes come back in document order'
		)
		assert(
			master.shapes.every((shape) => shape.host === master),
			'master shapes back-reference the master as their host'
		)
	})

	test('a master shape carries the paint surface, not only its box', async () => {
		const master = (await openFixture('mixed')).masters()[0]
		const band = master.shapes.find((shape) => shape.name === 'Rectangle 7')
		assert(band, 'the master has a "Rectangle 7" bar')

		// Geometry (read off the fixture's own slideMaster1.xml a:xfrm).
		assertEqual(band.left, 635000, 'left EMU')
		assertEqual(band.top, 504825, 'top EMU')
		assertEqual(band.width, 31750, 'width EMU')
		assertEqual(band.height, 1052513, 'height EMU')
		assertEqual(band.presetGeometry, 'rect', 'preset geometry')

		// The fill is a schemeClr, so a literal hex proves the master's *own* clrMap +
		// its theme were threaded into the shape — a slide's context is not involved.
		assertEqual(band.fillSchemeColor, 'bg2', 'the raw fill reference is a scheme colour')
		assertEqual(band.resolvedFill?.effectiveHex, '1C1C1C', 'bg2 resolves through the master clrMap to a literal hex')

		const accent = master.shapes.find((shape) => shape.name === 'Rectangle 2')
		assertEqual(accent.fillSchemeColor, 'accent2', 'a second decorative rect fills from accent2')
		assertEqual(accent.resolvedFill?.effectiveHex, 'FFCF01', 'accent2 resolves to the theme yellow')

		// A title placeholder in the same tree sets an explicit no-fill; the flag
		// separates that from "inherits a fill", which resolvedFill cannot.
		const title = master.shapes.find((shape) => shape.placeholder?.type === 'title')
		assertEqual(title.fillNoFill, true, 'the title placeholder sets an explicit a:noFill')
		assertEqual(title.resolvedFill, null, 'and so resolves no fill colour')
	})

	test('placeholders is the same tree, filtered — same elements, same ids', async () => {
		const master = (await openFixture('mixed')).masters()[0]

		const fromShapes = master.shapes.filter((shape) => shape.placeholder !== null)
		assertEqual(
			fromShapes.map((shape) => shape.id).join(','),
			master.placeholders.map((ph) => ph.id).join(','),
			'the placeholder subset of shapes matches placeholders, in the same order'
		)
		assert(
			fromShapes.every((shape, i) => shape.element_ === master.placeholders[i].element_),
			'both views hand out the same live p:sp elements'
		)
		assertEqual(
			fromShapes.map((shape) => shape.placeholder.type).join(','),
			master.placeholders.map((ph) => ph.type).join(','),
			'and agree on p:ph@type'
		)
	})

	test('shapeByIdDeep finds a master shape by drawing id', async () => {
		const master = (await openFixture('mixed')).masters()[0]
		assertEqual(master.shapeByIdDeep(1031)?.name, 'Rectangle 7', 'a top-level master shape resolves by id')
		assertEqual(master.shapeByIdDeep(999999), undefined, 'an id no shape carries resolves to undefined')
	})
})

describe('SlideLayout.shapes — PowerPoint-authored layout furniture', () => {
	test('a layout group recurses, and its children compose to slide-absolute frames', async () => {
		const master = (await openFixture('mixed')).masters()[0]
		const title = master.layouts.find((layout) => layout.name === 'Diapositive de titre')
		assert(title, 'the title layout reads back')

		assertEqual(title.shapes.length, 6, 'the layout spTree holds six shapes')
		assertEqual(decorative(title).length, 1, 'one of them is a non-placeholder group')

		const group = decorative(title)[0]
		assertEqual(group.shapeType, 'group', 'the decoration is a p:grpSp')
		assertEqual(group.name, 'Group 2', 'the group name')
		assertEqual(group.shapes.length, 5, 'the group nests five children')

		// A child's own box is in the group's child coordinate space; absoluteFrame maps
		// it out through the group chain, exactly as it does for a slide-level group.
		const rule = group.shapes.find((shape) => shape.name === 'Rectangle 11')
		assertEqual(rule.left, 199, 'the child box is in child-space units, not EMU')
		assertEqual(
			JSON.stringify(rule.absoluteFrame),
			JSON.stringify({
				left: 315913,
				top: 3260725,
				width: 8693150,
				height: 55563,
				rotation: 0,
				flipH: false,
				flipV: true,
			}),
			'the child composes to a slide-absolute frame, flipV carried'
		)

		// The deep lookup descends into the layout's groups.
		assertEqual(title.shapeByIdDeep(group.shapes[0].id)?.name, 'Group 3', 'a nested group resolves by id')
	})

	test('decorative text on a layout reads its text frame (read-stress.pptx)', async () => {
		const pres = await openFixture('read-stress')
		const quote = pres
			.masters()
			.flatMap((master) => master.layouts)
			.find((layout) => layout.name === 'Quote with Caption')
		assert(quote, 'the "Quote with Caption" layout reads back')

		assertEqual(quote.shapes.length, 8, 'the layout spTree holds eight shapes')
		assertEqual(quote.placeholders.length, 6, 'six carry a p:ph')

		const marks = decorative(quote)
		assertEqual(marks.length, 2, 'the two quotation-mark text boxes are the decoration')
		assertEqual(
			marks.map((shape) => shape.text).join(''),
			'“”',
			'their text — invisible through placeholders — reads back'
		)
		assertEqual(marks[0].left, 541870, 'the opening mark keeps its own geometry')
	})
})

describe('showMasterSp — whether the master shapes are drawn', () => {
	test('a layout that suppresses the master reads false; its siblings read true', async () => {
		for (const [fixture, suppressing] of [
			['mixed', 'Diapositive de titre'],
			['read-stress', 'Title Slide'],
		]) {
			const layouts = (await openFixture(fixture)).masters().flatMap((master) => master.layouts)
			const hidden = layouts.filter((layout) => !layout.showMasterSp)
			assertEqual(hidden.length, 1, `${fixture}: exactly one layout sets showMasterSp="0"`)
			assertEqual(hidden[0].name, suppressing, `${fixture}: it is the title layout`)
			assert(
				layouts.filter((layout) => layout !== hidden[0]).every((layout) => layout.showMasterSp),
				`${fixture}: every other layout omits the attribute and so reads true`
			)
		}
	})

	test('a slide with no attribute reads true (absent ⇒ shown)', async () => {
		for (const fixture of ['mixed', 'read-stress', 'textbox']) {
			const slides = (await openFixture(fixture)).slides
			assert(
				slides.length > 0 && slides.every((slide) => slide.showMasterSp),
				`${fixture}: every slide defaults to showing the master shapes`
			)
		}
	})

	test('a slide that sets showMasterSp="0" reads false across a round trip', async () => {
		const pres = await openFixture('mixed')
		const slide = pres.slides[0]
		slide.element_.setAttribute('showMasterSp', '0')
		slide.markDirty()
		assertEqual(slide.showMasterSp, false, 'the getter reads the attribute off p:sld')

		const reopened = await Presentation.load(await pres.save())
		assertEqual(reopened.slides[0].showMasterSp, false, 'and it survives serialization')
		assertEqual(reopened.slides[1].showMasterSp, true, 'a sibling slide is untouched')
	})
})

describe('SlideLayout.shapes — write→read fidelity for defineSlideMaster objects', () => {
	/**
	 * A master definition whose `objects` mixes a decorative rect, a decorative text
	 * box, and a real placeholder. The first two are the members that had no accessor.
	 */
	function authorBrandedDeck() {
		return authorRead((pres) => {
			pres.layout = 'LAYOUT_16x9'
			pres.defineSlideMaster({
				title: 'Branded',
				background: { color: 'FFFFFF' },
				objects: [
					{ rect: { x: 0, y: 0, w: 13.33, h: 0.6, fill: { color: '250F6B' } } },
					{ text: { text: 'ACME', options: { x: 11.8, y: 0.05, w: 1.4, h: 0.5, color: 'FFFFFF' } } },
					{ placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 1.2, w: 12, h: 1.2 } } },
				],
			})
			pres.addSlide({ masterTitle: 'Branded' }).addText('own content', { x: 0.5, y: 3, w: 6, h: 1 })
		})
	}

	test('the band and the wordmark come back, and stay out of placeholders', async () => {
		const { presentation } = await authorBrandedDeck()
		const layout = presentation.slides[0].layout
		assert(layout, 'the slide binds to the Branded layout')

		assertEqual(layout.shapes.length, 3, 'all three authored objects are reachable')
		assertEqual(layout.placeholders.length, 1, 'only the placeholder is in the filtered view')

		const band = layout.shapes[0]
		assertEqual(band.placeholder, null, 'the band is not a placeholder')
		assertEqual(band.fillColor, '250F6B', 'its authored fill round-trips')
		assertEqual(band.resolvedFill?.effectiveHex, '250F6B', 'and resolves to the same literal hex')
		assertEqual(band.top, 0, 'its top EMU round-trips')
		assertEqual(band.width, 12188952, 'its width round-trips (13.33in × 914400 EMU)')

		const wordmark = layout.shapes[1]
		assertEqual(wordmark.text, 'ACME', 'the wordmark text round-trips')
		assertEqual(wordmark.placeholder, null, 'and it is not a placeholder either')

		assertEqual(layout.shapes[2].placeholder?.type, 'title', 'the third object is the title placeholder')
		assertEqual(
			layout.shapes[2].element_,
			layout.placeholders[0].element_,
			'both views hand out the same placeholder element'
		)
	})

	test('the master under an authored deck has an empty tree, and says so', async () => {
		const { presentation } = await authorBrandedDeck()
		const master = presentation.slides[0].master
		assert(master, 'the slide resolves its master')
		// `defineSlideMaster` creates a *layout* under the shared master, so an authored
		// deck puts nothing on the master's own spTree at all. The accessor reports that
		// as `[]` rather than reaching into the layout — which is exactly why the master
		// arm of this feature is proven against PowerPoint output above, not here.
		assertEqual(master.shapes.length, 0, 'an authored master carries no shapes of its own')
		assertEqual(master.placeholders.length, 0, 'and therefore no placeholders either')
	})

	test.skipIf(!validatorInstalled)('the authored branded deck is schema-valid', async () => {
		const { buf } = await authorBrandedDeck()
		assertEqual((await schemaErrors(buf)).length, 0, 'branded deck validates')
	})
})
