/**
 * Presence checks for constructs the read model has no accessor for.
 *
 * These reach through the documented raw hatch (`element_`) rather than through a typed
 * accessor, which needs justifying. The rule they follow: **detect, never decode.** Each
 * function answers only "is this here?", so the answer can become a {@link FidelityNote}.
 * Nothing here interprets a value or feeds one into a call — that would be building a
 * private read path alongside the public one, and it would rot the moment the element's
 * meaning changed.
 *
 * The alternative is worse than the coupling. Without these checks, a deck with an embedded
 * video converts to a static poster image and a deck with an equation converts to an empty
 * box, both silently and both with a clean round-trip, because a diff cannot see what
 * neither side read. A loss you can name is a bug report; a loss nothing observes is a
 * defect that ships.
 *
 * Each check here is a standing argument for a real accessor. When one lands, the
 * corresponding function should be deleted, not kept in sync.
 */

/** DrawingML main namespace — hosts `a:videoFile` / `a:audioFile` under `p:nvPr`. */
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

/** Office MathML namespace — the equation body itself, wrapped by `a14:m` in a text run. */
const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

/** PresentationML main namespace — `p:grpSpPr`. */
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'

/** A minimal structural view of a DOM element, so this module needs no DOM lib types. */
interface ElementLike {
	namespaceURI: string | null
	localName: string | null
	firstChild: NodeLike | null
	getAttribute(name: string): string | null
	getElementsByTagNameNS(namespaceURI: string, localName: string): { length: number }
}

interface NodeLike {
	nodeType: number
	nextSibling: NodeLike | null
}

const ELEMENT_NODE = 1

function has(element: unknown, namespaceURI: string, localName: string): boolean {
	const el = element as ElementLike | null
	if (!el || typeof el.getElementsByTagNameNS !== 'function') return false
	return el.getElementsByTagNameNS(namespaceURI, localName).length > 0
}

/**
 * The first *direct* child with this name.
 *
 * A subtree search is wrong for structural lookups, in two ways that both bite here.
 * `a:ext` names two unrelated elements — a shape extent (`@cx`/`@cy`) and an extension-list
 * entry (`@uri`) — so a subtree search finds whichever comes first in document order. And on
 * a group it descends into the group's own children, so a nested group's transform answers
 * for its parent's.
 */
function child(element: ElementLike | null, namespaceURI: string, localName: string): ElementLike | null {
	if (!element) return null
	for (let node = element.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const el = node as unknown as ElementLike
		if (el.namespaceURI === namespaceURI && el.localName === localName) return el
	}
	return null
}

/**
 * `true` when this shape is an embedded audio or video rather than a still picture.
 *
 * PowerPoint authors media as a `p:pic` whose `p:nvPr` carries an `a:videoFile` or
 * `a:audioFile`, with the poster frame as the picture's own blip. The read model reports
 * only the poster — `mediaKind` distinguishes raster from SVG, not still from moving — so
 * without this check a video converts to a static image that looks deliberate.
 */
export function isAudioVideo(element: unknown): boolean {
	return has(element, A_NS, 'videoFile') || has(element, A_NS, 'audioFile')
}

/**
 * `true` when this shape's text body contains an OMML equation.
 *
 * An equation lives in `m:oMath` inside an `a14:m` run, which contributes nothing to
 * `TextFrame.text`, so an equation-only shape reads as an empty text frame. The write API
 * *can* author one (`TextProps.math`, plus the `ts-pptx/math` subpath for LaTeX/MathML
 * input), which makes this a read-side gap rather than a hard limit.
 */
export function hasEquation(element: unknown): boolean {
	return has(element, MATH_NS, 'oMath')
}

/**
 * `true` when this `p:sp` is a text box rather than an auto shape.
 *
 * The discriminator is one attribute, `p:cNvSpPr/@txBox`, and nothing reads it. The
 * tempting substitute — "it has no preset geometry" — is wrong for the common case:
 * PowerPoint gives every text box an explicit `<a:prstGeom prst="rect"/>`, so that test
 * calls every real text box an auto shape. The difference is not cosmetic. PowerPoint
 * autofits, wraps and resizes a text box by different rules, and offers a different
 * right-click menu for it.
 *
 * Direct child rather than a subtree search: `p:cNvSpPr` sits under `p:nvSpPr`, and a group
 * would otherwise answer for its first child.
 */
export function isTextBox(element: unknown): boolean {
	const nvSpPr = child(element as ElementLike | null, P_NS, 'nvSpPr')
	return child(nvSpPr, P_NS, 'cNvSpPr')?.getAttribute('txBox') === '1'
}

/**
 * `true` when a group's child coordinate space equals its own frame, so its children are
 * already in slide-absolute units.
 *
 * `addGroup` always emits an identity child space; a source group whose `a:chOff`/`a:chExt`
 * differ scales or shifts its contents, and that behaviour cannot be reproduced. The read
 * model has no accessor for the child space — `absoluteFrame` deliberately composes it away,
 * which is right for locating a shape and wrong for spotting this difference.
 *
 * A group with no explicit transform inherits an identity one, so a missing `a:xfrm` is
 * `true` rather than unknown.
 */
export function hasIdentityChildSpace(element: unknown): boolean {
	const xfrm = child(child(element as ElementLike | null, P_NS, 'grpSpPr'), A_NS, 'xfrm')
	if (!xfrm) return true

	const off = child(xfrm, A_NS, 'off')
	const ext = child(xfrm, A_NS, 'ext')
	const chOff = child(xfrm, A_NS, 'chOff')
	const chExt = child(xfrm, A_NS, 'chExt')
	if (!off || !ext || !chOff || !chExt) return true

	return (
		off.getAttribute('x') === chOff.getAttribute('x') &&
		off.getAttribute('y') === chOff.getAttribute('y') &&
		ext.getAttribute('cx') === chExt.getAttribute('cx') &&
		ext.getAttribute('cy') === chExt.getAttribute('cy')
	)
}
