import { describe, expect, test } from 'vitest'
import { slideObjectToXml, slideObjectRelationsToXml } from '../../../src/gen/slide/object.ts'
import { SlideObjectType } from '../../../src/enums.ts'

// Byte-pin for Slide / Section / Summary Zoom (dn-zoom-links). These emit `<mc:AlternateContent>`
// wrappers the demo deck never produces, so the byte-identity harness is blind to them — pin the
// load-bearing structure (namespaces, target ids, zmPr attrs, fallback hlink) here. Ground truth:
// a PowerPoint-authored fixture (plan `foamy-imagining-narwhal`); emitter verified to reopen in
// PowerPoint with live-regenerated previews.

const LAYOUT = { name: 'test', width: 12192000, height: 6858000 }

const mkSlide = (objects, extra = {}) => ({
	_slideNum: 5,
	_slideObjects: objects,
	_presLayout: LAYOUT,
	_rels: [],
	_relsChart: [],
	_relsMedia: [],
	...extra,
})
const render = (objects, extra = {}) => slideObjectToXml(mkSlide(objects, extra))

const zoomObj = (variant, zoom, options = {}) => ({
	_type: SlideObjectType.zoom,
	options: { x: 1, y: 1, w: 3, h: 1.7, objectName: `${variant} zoom`, ...options },
	zoom: { variant, returnToParent: false, transitionDur: 1000, ...zoom },
})

describe('Slide Zoom', () => {
	const xml = render([
		zoomObj('slide', { tiles: [{ sldId: 257, previewRid: 2, fallbackSlideRid: 3, zmPrId: '{AAA}' }] }),
	])

	test('wrapped in mc:AlternateContent with a pslz Choice', () => {
		expect(xml).toContain(
			'<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
		)
		expect(xml).toContain('xmlns:pslz="http://schemas.microsoft.com/office/powerpoint/2016/slidezoom" Requires="pslz"')
		expect(xml).toContain('uri="http://schemas.microsoft.com/office/powerpoint/2016/slidezoom"')
	})
	test('sldZmObj carries the target sldId; zmPr has returnToParent + transitionDur', () => {
		expect(xml).toContain('<pslz:sldZm><pslz:sldZmObj sldId="257">')
		expect(xml).toContain('<pslz:zmPr id="{AAA}" returnToParent="0" transitionDur="1000">')
	})
	test('preview blip + framed spPr in the 2016/6/main namespace', () => {
		expect(xml).toContain(
			'<p166:blipFill xmlns:p166="http://schemas.microsoft.com/office/powerpoint/2016/6/main"><a:blip r:embed="rId2"/>'
		)
		expect(xml).toContain('<a:ln w="3175"><a:solidFill><a:prstClr val="ltGray"/></a:solidFill></a:ln>')
	})
	test('mc:Fallback is a hyperlinked picture (hlinksldjump) reusing the preview blip', () => {
		expect(xml).toContain('<mc:Fallback><p:pic>')
		expect(xml).toContain('<a:hlinkClick r:id="rId3" action="ppaction://hlinksldjump"/>')
	})
})

describe('Section Zoom', () => {
	const xml = render([
		zoomObj('section', { tiles: [{ sectionId: '{SEC}', previewRid: 2, fallbackSlideRid: 3, zmPrId: '{BBB}' }] }),
	])
	test('sectionZmObj carries the section GUID', () => {
		expect(xml).toContain('Requires="psez"')
		expect(xml).toContain('<psez:sectionZm><psez:sectionZmObj sectionId="{SEC}">')
	})
	test('section zmPr OMITS returnToParent (matches PowerPoint)', () => {
		expect(xml).toContain('<psez:zmPr id="{BBB}" transitionDur="1000">')
		expect(xml).not.toContain('psez:zmPr id="{BBB}" returnToParent')
	})
})

describe('Summary Zoom', () => {
	const xml = render([
		zoomObj(
			'summary',
			{
				tiles: [
					{
						sectionId: '{S1}',
						previewRid: 2,
						fallbackSlideRid: 3,
						zmPrId: '{C1}',
						grid: { x: 0, y: 0, cx: 100, cy: 56 },
					},
					{
						sectionId: '{S2}',
						previewRid: 2,
						fallbackSlideRid: 4,
						zmPrId: '{C2}',
						grid: { x: 120, y: 0, cx: 100, cy: 56 },
					},
				],
			},
			{ w: 10, h: 4 }
		),
	])
	test('one summaryZmObj per section, terminated by gridLayout', () => {
		expect(xml).toContain('<psuz:summaryZm><psuz:summaryZmObj sectionId="{S1}">')
		expect(xml).toContain('<psuz:summaryZmObj sectionId="{S2}">')
		expect(xml).toContain('<psuz:gridLayout/></psuz:summaryZm>')
	})
	test('each tile uses its own grid cell as the frame-local spPr xfrm', () => {
		expect(xml).toContain('<a:off x="0" y="0"/><a:ext cx="100" cy="56"/>')
		expect(xml).toContain('<a:off x="120" y="0"/><a:ext cx="100" cy="56"/>')
	})
	test('fallback is a p:grpSp of hyperlinked pics (one per tile)', () => {
		expect(xml).toContain('<mc:Fallback><p:grpSp>')
		expect(xml).toContain('<a:hlinkClick r:id="rId3" action="ppaction://hlinksldjump"/>')
		expect(xml).toContain('<a:hlinkClick r:id="rId4" action="ppaction://hlinksldjump"/>')
	})
})

describe('Zoom relationships', () => {
	test('a data:"slide" rel emits a .../slide relationship to slideN.xml', () => {
		const rels = slideObjectRelationsToXml(
			mkSlide([], { _rels: [{ type: SlideObjectType.hyperlink, data: 'slide', rId: 3, Target: '2' }] }),
			[]
		)
		expect(rels).toContain(
			'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide2.xml"'
		)
	})
})
