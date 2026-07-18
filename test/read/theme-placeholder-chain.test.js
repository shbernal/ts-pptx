// Read-model coverage for the placeholder-inheritance CHAIN in
// src/read/oxml/theme.ts — the layout → master-placeholder → master-txStyles
// tiers of placeholderInheritedFill / placeholderInheritedDefRPrs /
// placeholderInheritedAnchor, plus findPlaceholder's idx / category fall-backs
// and the phCategory('other') mapping. The real fixtures resolve through a
// single deck's actual layout+master, so the individual tier-selection and
// fall-back branches never all fire. Here a synthetic FlattenContext carries
// hand-authored layout / master roots, and a placeholder TextFrame drives the
// chain through Run.resolvedColor / resolvedSizePt and TextFrame.resolvedAnchor.

import { DOMParser } from '@xmldom/xmldom'
import { describe, test } from 'vitest'
import { TextFrame } from '../../dist/read.js'
import { assertEqual } from '../helpers.js'

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

const stubPart = () => ({ markDirty() {} })

function parse(xml) {
	return new DOMParser().parseFromString(xml, 'text/xml').documentElement
}

/** A layout/master root (`p:sldLayout` / `p:sldMaster`) wrapping shape-tree XML. */
function root(local, spTreeXml) {
	return parse(
		`<p:${local} xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:cSld><p:spTree>${spTreeXml}</p:spTree></p:cSld>${''}</p:${local}>`
	)
}

/** A master root with a `p:txStyles` block appended after the shape tree. */
function masterWithTxStyles(spTreeXml, txStylesXml) {
	return parse(
		`<p:sldMaster xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:cSld><p:spTree>${spTreeXml}</p:spTree></p:cSld>${txStylesXml}</p:sldMaster>`
	)
}

/** A placeholder shape carrying an a:lstStyle (and optional a:bodyPr) in its txBody. */
function phSp(type, idx, { lstStyle = '', bodyPr = '<a:bodyPr/>' } = {}) {
	const phAttrs = `${type === null ? '' : `type="${type}" `}idx="${idx}"`
	return (
		`<p:sp><p:nvSpPr><p:cNvPr id="2" name="ph"/><p:cNvSpPr/><p:nvPr><p:ph ${phAttrs}/></p:nvPr></p:nvSpPr>` +
		`<p:spPr/><p:txBody>${bodyPr}<a:lstStyle>${lstStyle}</a:lstStyle><a:p/></p:txBody></p:sp>`
	)
}

/** A lvl1 defRPr solidFill lstStyle fragment. */
const lvl1Fill = (hex) =>
	`<a:lvl1pPr><a:defRPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></a:defRPr></a:lvl1pPr>`
const lvl1Size = (sz) => `<a:lvl1pPr><a:defRPr sz="${sz}"/></a:lvl1pPr>`

/** The first run of a placeholder TextFrame resolving against `flatten`. */
function phRun(flatten, ph = { type: 'body', idx: '0' }) {
	const txBody = parse(
		`<p:txBody xmlns:p="${P_NS}" xmlns:a="${A_NS}"><a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p></p:txBody>`
	)
	return new TextFrame(txBody, /** @type {any} */ (stubPart()), flatten, { ph, flatten }).paragraphs[0].runs[0]
}

function ctx(overrides) {
	return {
		clrMap: new Map(),
		clrScheme: new Map(),
		fmtScheme: null,
		fontScheme: null,
		layoutRoot: null,
		masterRoot: null,
		...overrides,
	}
}

describe('placeholderInheritedFill — tier selection', () => {
	test('the layout placeholder lstStyle is the first tier', () => {
		const layoutRoot = root('sldLayout', phSp('body', '0', { lstStyle: lvl1Fill('AA0000') }))
		assertEqual(phRun(ctx({ layoutRoot })).resolvedColor.hex, 'AA0000', 'colour comes from the layout placeholder')
	})

	test('the master placeholder lstStyle is used when the layout defines none', () => {
		const masterRoot = root('sldMaster', phSp('body', '0', { lstStyle: lvl1Fill('00AA00') }))
		assertEqual(phRun(ctx({ masterRoot })).resolvedColor.hex, '00AA00', 'colour comes from the master placeholder')
	})

	test('the master p:txStyles category style is the final tier', () => {
		// No placeholder in the master shape tree, only a titleStyle — reached via phCategory.
		const masterRoot = masterWithTxStyles(
			'',
			`<p:txStyles><p:titleStyle>${lvl1Fill('0000AA')}</p:titleStyle></p:txStyles>`
		)
		assertEqual(
			phRun(ctx({ masterRoot }), { type: 'title', idx: '0' }).resolvedColor.hex,
			'0000AA',
			'title runs read titleStyle'
		)
	})

	test("an 'other'-category placeholder (e.g. ftr) reads p:otherStyle", () => {
		const masterRoot = masterWithTxStyles(
			'',
			`<p:txStyles><p:otherStyle>${lvl1Fill('AA00AA')}</p:otherStyle></p:txStyles>`
		)
		assertEqual(
			phRun(ctx({ masterRoot }), { type: 'ftr', idx: '0' }).resolvedColor.hex,
			'AA00AA',
			'ftr → other category'
		)
	})
})

describe('findPlaceholder — idx / category fall-backs', () => {
	test('a same-idx placeholder of a different category is the idx fall-back', () => {
		// Run is body/idx0; the layout only has a title/idx0 placeholder → idx match.
		const layoutRoot = root('sldLayout', phSp('title', '0', { lstStyle: lvl1Fill('123456') }))
		assertEqual(phRun(ctx({ layoutRoot })).resolvedColor.hex, '123456', 'falls back to the same-idx placeholder')
	})

	test('a same-category placeholder at a different idx is the category fall-back', () => {
		// Run is body/idx5; the layout only has a body/idx0 placeholder → category match.
		const layoutRoot = root('sldLayout', phSp('body', '0', { lstStyle: lvl1Fill('654321') }))
		assertEqual(
			phRun(ctx({ layoutRoot }), { type: 'body', idx: '5' }).resolvedColor.hex,
			'654321',
			'falls back on category'
		)
	})
})

describe('inherited size + anchor through the chain', () => {
	test('resolvedSizePt walks the master txStyles defRPr sz', () => {
		const masterRoot = masterWithTxStyles('', `<p:txStyles><p:bodyStyle>${lvl1Size('3200')}</p:bodyStyle></p:txStyles>`)
		assertEqual(phRun(ctx({ masterRoot })).resolvedSizePt, 32, 'sz 3200 → 32pt from bodyStyle')
	})

	test('resolvedAnchor skips a null layout root and reads the master placeholder bodyPr', () => {
		// layoutRoot null forces the loop to skip its first (null) tier before the master.
		const masterRoot = root('sldMaster', phSp('body', '0', { bodyPr: '<a:bodyPr anchor="ctr"/>' }))
		const txBody = parse(`<p:txBody xmlns:p="${P_NS}" xmlns:a="${A_NS}"><a:bodyPr/><a:p/></p:txBody>`)
		const flatten = ctx({ layoutRoot: null, masterRoot })
		const frame = new TextFrame(txBody, /** @type {any} */ (stubPart()), flatten, {
			ph: { type: 'body', idx: '0' },
			flatten,
		})
		assertEqual(frame.resolvedAnchor, 'ctr', 'anchor inherited from the master placeholder bodyPr')
	})
})
