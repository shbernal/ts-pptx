import { describe, expect, test } from 'vitest'
import { slideObjectToXml, slideObjectRelationsToXml } from '../../src/gen/slide/object.ts'
import { SlideObjectType } from '../../src/core-enums.ts'

// Characterization tests for slide-object XML that the byte-identity harness CANNOT see. The demo
// deck emits ZERO parts containing `<a:duotone>`, `<a:stCxn>`, `mc:AlternateContent`,
// `<a:tableStyleId>` or any of the `<a:tblPr>` flags (rtl/firstRow/lastRow/bandRow/bandCol/
// firstCol/lastCol), and only one or two parts for several more — so a green gate says nothing
// about them. These pin the exact bytes so the next refactor is not blind.
//
// Pinning is not endorsement — where the current output is a latent bug it is called out as such.

const LAYOUT = { name: 'test', width: 9144000, height: 6858000 }

const mkSlide = (objects, extra = {}) => ({
	_slideNum: 1,
	_slideObjects: objects,
	_presLayout: LAYOUT,
	_rels: [],
	_relsChart: [],
	_relsMedia: [],
	...extra,
})

const render = (objects, extra = {}) => slideObjectToXml(mkSlide(objects, extra))

const textObj = (options = {}) => ({
	_type: SlideObjectType.text,
	text: [{ text: 'x' }],
	shape: 'rect',
	options: { objectName: 'T', ...options },
})

describe('escaping: cNvPrOpen leaves objectName as-is (caller escapes upstream); cSld-name is escaped here', () => {
	// This test drives `slideObjectToXml` directly with an already-raw `options.objectName`,
	// bypassing the define layer (`addTextDefinition` et al.) that normally escapes it first via
	// `encodeXmlEntities(validateObjectName(...))`. So "NOT escaped" here describes cNvPrOpen in
	// isolation — intentional, since the real `addText()` API escapes once upstream (see
	// `cNvPrOpen` in src/gen/slide/object.ts) and escaping again here would double-encode it.
	test('objectName is NOT escaped (by this layer — the define layer escapes it upstream)', () => {
		expect(render([textObj({ objectName: 'Q&A' })])).toContain('name="Q&A"')
	})

	test('altText IS escaped, in the same element', () => {
		expect(render([textObj({ altText: 'a & <b>' })])).toContain('descr="a &amp; &lt;b&gt;"')
	})

	// Unlike objectName, `_name` (-> `<p:cSld name>`) is escaped HERE, at this render layer, not
	// upstream: `_name` doubles as the raw lookup key `addSlide({masterTitle})` matches against the
	// caller's `title` (pptxgen.ts, `layout._name === masterSlideName`), so it must stay unescaped
	// until emission or that match breaks for a title containing XML metacharacters. Fixed for
	// backlog `fork-slidemaster-title-unescaped`.
	test('the slide name IS escaped, at this render layer', () => {
		expect(render([], { _name: 'R&D' })).toContain('<p:cSld name="R&amp;D">')
	})

	test('a hyperlink tooltip IS escaped, and exactly once', () => {
		const xml = render([textObj({ hyperlink: { url: 'http://a.b', _rId: 3, tooltip: 'a & b' } })])
		expect(xml).toContain('tooltip="a &amp; b"')
		expect(xml).not.toContain('&amp;amp;')
	})
})

describe('table properties (ZERO baseline parts for every flag below)', () => {
	const table = (options) => ({
		_type: SlideObjectType.table,
		arrTabRows: [[{ _type: SlideObjectType.tablecell, text: 'a', options: {} }]],
		options: { objectName: 'T', ...options },
	})

	test('no flags emits a self-closing tblPr', () => {
		expect(render([table({})])).toContain('<a:tblPr/>')
	})

	test('a table style makes tblPr PAIRED, not self-closing', () => {
		const xml = render([table({ tableStyle: '{GUID}' })])
		expect(xml).toContain('<a:tblPr><a:tableStyleId>{GUID}</a:tableStyleId></a:tblPr>')
	})

	test('flag ORDER is rtl, firstRow, lastRow, bandRow, bandCol, firstCol, lastCol', () => {
		const xml = render([
			table({
				rtl: true,
				hasHeader: true,
				hasFooter: true,
				hasBandedRows: true,
				hasBandedColumns: true,
				hasFirstColumn: true,
				hasLastColumn: true,
			}),
		])
		expect(xml).toContain(
			'<a:tblPr rtl="1" firstRow="1" lastRow="1" bandRow="1" bandCol="1" firstCol="1" lastCol="1"/>'
		)
	})

	test('flags combine with a style id', () => {
		expect(render([table({ hasHeader: true, tableStyle: '{G}' })])).toContain(
			'<a:tblPr firstRow="1"><a:tableStyleId>{G}</a:tableStyleId></a:tblPr>'
		)
	})
})

describe('table cells', () => {
	const cell = (text, options = {}) => ({ _type: SlideObjectType.tablecell, text, options })
	const table = (rows, options = {}) => ({
		_type: SlideObjectType.table,
		arrTabRows: rows,
		options: { objectName: 'T', ...options },
	})

	test('tcPr attribute order is margins, then anchor, then vert', () => {
		const xml = render([table([[cell('a', { valign: 'middle', textDirection: 'vert' })]])])
		expect(xml).toMatch(/<a:tcPr marL="\d+" marR="\d+" marT="\d+" marB="\d+" anchor="ctr" vert="vert">/)
	})

	test('a real cell carries byte-significant indentation before its closing tags', () => {
		expect(render([table([[cell('a')]])])).toContain('  </a:tcPr> </a:tc>')
	})

	test('a covered (merged) cell is FLAT, unlike a real cell', () => {
		const xml = render([
			table([
				[cell('a', { colspan: 2 }), cell('c')],
				[cell('d'), cell('e'), cell('f')],
			]),
		])
		expect(xml).toContain('<a:tc hMerge="1"><a:tcPr></a:tcPr></a:tc>')
	})

	test('span attribute order is rowSpan, gridSpan, vMerge, hMerge', () => {
		const xml = render([table([[cell('a', { colspan: 2, rowspan: 2 }), cell('c')], [cell('d')]])])
		expect(xml).toContain('<a:tc rowSpan="2" gridSpan="2">')
		expect(xml).toContain('<a:tc rowSpan="2" hMerge="1">')
	})

	test('textDirection horz is omitted rather than emitted', () => {
		expect(render([table([[cell('a', { textDirection: 'horz' })]])])).not.toContain('vert="horz"')
	})
})

describe('connector shape bindings (stCxn has ZERO baseline parts)', () => {
	const withConnector = (options) =>
		render([
			textObj({ objectName: 'A' }),
			textObj({ objectName: 'B' }),
			{ _type: SlideObjectType.connector, shape: 'bentConnector3', options: { objectName: 'C', ...options } },
		])

	test('start and end bindings emit in schema order', () => {
		const xml = withConnector({ _startCxn: { name: 'A', idx: 0 }, _endCxn: { name: 'B', idx: 3 } })
		expect(xml).toContain('<p:cNvCxnSpPr><a:stCxn id="2" idx="0"/><a:endCxn id="3" idx="3"/></p:cNvCxnSpPr>')
	})

	test('no bindings emits a self-closing cNvCxnSpPr', () => {
		expect(withConnector({})).toContain('<p:cNvCxnSpPr/>')
	})

	test('an unresolvable binding is dropped, not emitted as a dangling id', () => {
		const xml = withConnector({ _startCxn: { name: 'NOPE', idx: 0 } })
		expect(xml).toContain('<p:cNvCxnSpPr/>')
		expect(xml).not.toContain('a:stCxn')
	})

	test('connector adjust values emit as adj1..adjN', () => {
		const xml = withConnector({ _connectorAdj: [25000, 75000] })
		expect(xml).toContain('<a:avLst><a:gd name="adj1" fmla="val 25000"/><a:gd name="adj2" fmla="val 75000"/></a:avLst>')
	})
})

describe('image blip effects (duotone has ZERO baseline parts)', () => {
	const RASTER = [{ rId: 5, type: 'image/png', Target: '../media/i.png', extn: 'png' }]
	const SVG = [
		{ rId: 4, type: 'image/png', Target: '../media/i.png', extn: 'png' },
		{ rId: 5, type: 'image/svg+xml', Target: '../media/i.svg', extn: 'svg' },
	]
	const image = (options = {}) => ({
		_type: SlideObjectType.image,
		image: 'i.png',
		imageRid: 5,
		options: { objectName: 'I', ...options },
	})

	test('duotone emits shadow then highlight, after alphaModFix', () => {
		const xml = render([image({ transparency: 25, duotone: { shadow: 'FF0000', highlight: '00FF00' } })], {
			_relsMedia: RASTER,
		})
		expect(xml).toContain(
			'<a:alphaModFix amt="75000"/><a:duotone><a:srgbClr val="FF0000"/><a:srgbClr val="00FF00"/></a:duotone>'
		)
	})

	test('QUIRK: the SVG branch prefixes alphaModFix with a space and the raster branch does not', () => {
		const raster = render([image({ transparency: 25 })], { _relsMedia: RASTER })
		const svg = render([image({ transparency: 25 })], { _relsMedia: SVG })
		expect(raster).toContain('"><a:alphaModFix')
		expect(svg).toContain('"> <a:alphaModFix')
	})

	test('the SVG branch binds the raster fallback at rId-1 and the SVG at rId', () => {
		const xml = render([image()], { _relsMedia: SVG })
		expect(xml).toContain('<a:blip r:embed="rId4">')
		expect(xml).toContain('r:embed="rId5"/>')
	})
})

describe('media (audio/video/online share one body)', () => {
	const media = (mtype) => render([{ _type: SlideObjectType.media, mtype, mediaRid: 6, options: { objectName: 'M' } }])

	test('audio uses a:audioFile, bound by r:link', () => {
		expect(media('audio')).toContain('<a:audioFile r:link="rId6"/>')
	})

	test('video uses a:videoFile', () => {
		expect(media('video')).toContain('<a:videoFile r:link="rId6"/>')
	})

	test('an embedded file binds p14:media by r:embed', () => {
		expect(media('video')).toContain(`r:embed="rId7"/>`)
	})

	test('online video is a:videoFile but binds p14:media by r:link', () => {
		const xml = media('online')
		expect(xml).toContain('<a:videoFile r:link="rId6"/>')
		expect(xml).toContain(`r:link="rId7"/>`)
	})

	test('the preview image is always required, at rId+2', () => {
		expect(media('audio')).toContain('<a:blip r:embed="rId8"/>')
	})
})

describe('equation shapes (mc:AlternateContent has ZERO baseline parts)', () => {
	test('a math run wraps the whole shape in an a14 markup-compatibility envelope', () => {
		const xml = render([
			{
				_type: SlideObjectType.text,
				text: [{ text: 'E=mc^2', math: '<m:oMath><m:r><m:t>E=mc^2</m:t></m:r></m:oMath>' }],
				shape: 'rect',
				options: { objectName: 'M' },
			},
		])
		expect(xml).toContain(
			'<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
		)
		expect(xml).toContain('Requires="a14">')
		expect(xml).toContain('</mc:Choice></mc:AlternateContent>')
	})

	test('a plain text shape emits no envelope', () => {
		expect(render([textObj()])).not.toContain('mc:AlternateContent')
	})
})

describe('slide number placeholder', () => {
	const sn = (props) => render([], { _slideNumberProps: { x: 1, y: 1, ...props } })

	test('bold emits b="0" when unset — it is NOT omitted like other run properties', () => {
		expect(sn({})).toContain('<a:rPr b="0" lang="en-US"/>')
		expect(sn({ bold: true })).toContain('<a:rPr b="1" lang="en-US"/>')
	})

	test('margins map [Top, Right, Bottom, Left] onto lIns/tIns/rIns/bIns', () => {
		// The array is CSS order but the attributes emit l/t/r/b, so the indexing is 3/0/1/2.
		const xml = sn({ margin: [0.1, 0.2, 0.3, 0.4] })
		expect(xml).toContain(`<a:bodyPr lIns="${Math.round(0.4 * 914400)}" tIns="${Math.round(0.1 * 914400)}"`)
		expect(xml).toContain(`rIns="${Math.round(0.2 * 914400)}" bIns="${Math.round(0.3 * 914400)}"/>`)
	})

	test('QUIRK: a master (no _slideNum) emits the literal text "null"', () => {
		// LATENT BUG, pinned deliberately — this is what slideMaster1.xml actually contains today.
		// It is why the `<a:t>` child is built with an explicit `String(...)`: the element builder
		// skips a null child, which would silently change those bytes.
		const xml = slideObjectToXml({ ...mkSlide([], { _slideNumberProps: { x: 1, y: 1 } }), _slideNum: null })
		expect(xml).toContain('<a:t>null</a:t>')
	})

	test('no margin emits a bare bodyPr', () => {
		expect(sn({})).toContain('<a:bodyPr/>')
	})

	test('an unrecognized align falls back to left', () => {
		expect(sn({ align: 'weird' })).toContain('<a:pPr algn="l"/>')
		expect(sn({ align: 'center' })).toContain('<a:pPr algn="ctr"/>')
		expect(sn({ align: 'right' })).toContain('<a:pPr algn="r"/>')
	})

	test('a fontFace is escaped across latin/ea/cs', () => {
		const xml = sn({ fontFace: 'Arial & "X"' })
		expect(xml).toContain('<a:latin typeface="Arial &amp; &quot;X&quot;"/>')
		expect(xml).toContain('<a:cs typeface="Arial &amp; &quot;X&quot;"/>')
	})

	test('an empty lvl1pPr is still emitted when no font properties are set', () => {
		expect(sn({})).toContain('<a:lvl1pPr></a:lvl1pPr>')
	})
})

describe('groups', () => {
	const kid = (name, o = {}) => ({
		_type: SlideObjectType.text,
		text: [{ text: name }],
		shape: 'rect',
		options: { objectName: name, x: 1, y: 1, w: 2, h: 2, ...o },
	})

	test('an auto-sized group uses the bounding box of its children, with an identity child space', () => {
		const xml = render([
			{
				_type: SlideObjectType.group,
				_groupObjects: [kid('a'), kid('b', { x: 3, y: 4 })],
				options: { objectName: 'G' },
			},
		])
		// chOff/chExt mirror off/ext exactly, so children keep slide-absolute coordinates.
		expect(xml).toContain('<a:off x="914400" y="914400"/><a:ext cx="3657600" cy="4572000"/>')
		expect(xml).toContain('<a:chOff x="914400" y="914400"/><a:chExt cx="3657600" cy="4572000"/>')
	})

	test('a group with no locks emits a self-closing cNvGrpSpPr', () => {
		const xml = render([{ _type: SlideObjectType.group, _groupObjects: [kid('a')], options: { objectName: 'G' } }])
		expect(xml).toContain('<p:cNvGrpSpPr/>')
	})
})

describe('relationships', () => {
	const rels = (extra, defaults = []) => slideObjectRelationsToXml(mkSlide([], extra), defaults)

	test('an online-video pair shares one Target: ECMA video first, MS media second', () => {
		const target = 'https://y.t/?v=1&t=2'
		const xml = rels({
			_relsMedia: [
				{ rId: 4, type: 'online/youtube', Target: target },
				{ rId: 5, type: 'online/youtube', Target: target },
			],
		})
		// The dedupe compares the ESCAPED Target — an unescaped compare would mistype the pair.
		expect(xml).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"')
		expect(xml).toContain('Type="http://schemas.microsoft.com/office/2007/relationships/media"')
		expect(xml).toContain('Target="https://y.t/?v=1&amp;t=2" TargetMode="External"')
	})

	test('default rels start past the highest dynamic rId', () => {
		const xml = rels({ _rels: [{ rId: 7, type: SlideObjectType.hyperlink, Target: 'http://a' }] }, [
			{ target: 't.xml', type: 'http://x/theme' },
		])
		expect(xml).toContain('Id="rId8" Type="http://x/theme" Target="t.xml"')
	})
})
