/**
 * Mechanical proof that two XML parts differ **only** in inert inter-element whitespace.
 *
 * AGENTS.md: "Whitespace-only byte diffs are a STOP, not a known-divergence" — because
 * waving a diff through on the grounds that it *looks* like whitespace is the same
 * reasoning that would wave through a content change. This module exists so that the one
 * deliberate exception (flattening `src/gen/chart/`, see `docs/chart-whitespace-flatten.md`)
 * is settled by a program rather than by reading a diff. The STOP stays; what changes is
 * that there is now a way to discharge it that a human cannot get tired and do wrong.
 *
 * It is deliberately *stricter* than an XML-canonicalisation or DOM comparison, in four
 * ways that each correspond to a bug the emitters have actually been able to produce:
 *
 * 1. **Raw text, never decoded text.** A DOM parser turns `&amp;` and `&#38;` into the same
 *    character, so a DOM diff cannot see an escaping regression — the exact class of bug
 *    `gen/oxml/el.ts` was built to centralise away. Attribute values and character data are
 *    compared as the bytes on disk.
 * 2. **Attribute order is significant.** Inert per the XML spec, but not what is being
 *    proved here: this asserts a *whitespace-only* change, and a reordered attribute is not
 *    one. `chart-parts.ts` carries two `<a:defRPr>` orderings that would be nice to unify;
 *    that is a different change needing different evidence, and this must not silently
 *    absorb it.
 * 3. **Self-closing form is significant.** `<x/>` and `<x></x>` are the same element and
 *    different bytes. `el()` vs `voidEl()` decides this by arity precisely so it cannot
 *    drift on a value, so a change here is a real regression, not a formatting choice.
 * 4. **Whitespace only relaxes where it cannot be content.** See {@link isTextFrozen}.
 *
 * The comparison reports what it relaxed, so a caller can print the shape of the change
 * rather than just a verdict.
 */

/** Text that is XML whitespace and nothing else. Empty counts. */
const WHITESPACE = /^[\t\n\r ]*$/

/**
 * Elements whose content is character data in the parts this repo emits.
 *
 * This is the second lock, not the first. The load-bearing rule is structural — an element
 * with no element children never has its text relaxed (see {@link isTextFrozen}) — and it
 * already covers every one of these. The list guards the case the structural rule cannot
 * see on its own: an element that acquires an element child *and* carries significant text,
 * where relaxing "just the whitespace between the children" would eat content. Nothing in
 * the corpus is shaped that way today, which is exactly why a future change to one of these
 * should have to delete a name from this list on purpose.
 *
 * `si` and `is` are containers rather than leaves, and are here because the xlsx string
 * tables are not in scope for any flatten: whitespace between their `<t>` children is
 * frozen so a change there reports instead of being relaxed.
 */
const TEXT_BEARING = new Set([
	'a:t',
	'c:v',
	'c:f',
	'c:formatCode',
	'c:pt',
	'vt:lpstr',
	't',
	'si',
	'is',
	'dc:title',
	'dc:subject',
	'dc:creator',
	'dc:description',
	'cp:keywords',
	'cp:lastModifiedBy',
	'cp:category',
	'Company',
	'Manager',
	'Application',
	'AppVersion',
])

/** A part that could not be tokenized. Never a comparison verdict — always a stop. */
export class XmlSyntaxError extends Error {
	/**
	 * @param {string} message
	 * @param {number} offset byte offset into the source
	 */
	constructor(message, offset) {
		super(message + ' (at offset ' + offset + ')')
		this.name = 'XmlSyntaxError'
		this.offset = offset
	}
}

/**
 * @typedef {{name: string, gap: string, value: string, quote: string}} XmlAttr
 * @typedef {{kind: 'decl', raw: string}} DeclToken
 * @typedef {{kind: 'open', name: string, attrs: XmlAttr[], endGap: string, selfClosing: boolean}} OpenToken
 * @typedef {{kind: 'close', name: string}} CloseToken
 * @typedef {{kind: 'text', raw: string}} TextToken
 * @typedef {DeclToken | OpenToken | CloseToken | TextToken} XmlToken
 */

/**
 * @typedef {{type: 'element', name: string, attrs: XmlAttr[], endGap: string, selfClosing: boolean, children: XmlNode[]}} ElementNode
 * @typedef {{type: 'text', raw: string}} TextNode
 * @typedef {ElementNode | TextNode} XmlNode
 */

/**
 * Scan XML into a flat token stream, preserving every byte that is not structure.
 *
 * Strict on purpose: anything this does not fully understand — a comment, a CDATA section,
 * a DOCTYPE, a foreign processing instruction — throws rather than being skipped or passed
 * through. A lenient scanner would be the whole failure mode this module exists to avoid,
 * since the way it fails is by quietly agreeing that two parts match. The emitted corpus
 * contains none of these constructs (1163 XML parts checked); if one appears, that is news
 * and it should stop the gate.
 * @param {string} text
 * @returns {XmlToken[]}
 */
export function tokenizeXml(text) {
	/** @type {XmlToken[]} */
	const tokens = []
	let i = 0

	while (i < text.length) {
		if (text.charAt(i) !== '<') {
			const next = text.indexOf('<', i)
			const end = next === -1 ? text.length : next
			tokens.push({ kind: 'text', raw: text.slice(i, end) })
			i = end
			continue
		}

		if (text.startsWith('<!--', i)) throw new XmlSyntaxError('comment', i)
		if (text.startsWith('<![CDATA[', i)) throw new XmlSyntaxError('CDATA section', i)
		if (text.startsWith('<!', i)) throw new XmlSyntaxError('declaration (<!…)', i)

		if (text.startsWith('<?', i)) {
			if (!text.startsWith('<?xml', i)) throw new XmlSyntaxError('processing instruction', i)
			const end = text.indexOf('?>', i)
			if (end === -1) throw new XmlSyntaxError('unterminated XML declaration', i)
			tokens.push({ kind: 'decl', raw: text.slice(i, end + 2) })
			i = end + 2
			continue
		}

		if (text.startsWith('</', i)) {
			const end = text.indexOf('>', i)
			if (end === -1) throw new XmlSyntaxError('unterminated end tag', i)
			const name = text.slice(i + 2, end).trim()
			if (!name) throw new XmlSyntaxError('end tag with no name', i)
			tokens.push({ kind: 'close', name })
			i = end + 1
			continue
		}

		i = scanStartTag(text, i, tokens)
	}

	return tokens
}

/**
 * Scan one start tag beginning at `start`, push it, and return the offset just past `>`.
 * @param {string} text
 * @param {number} start offset of the `<`
 * @param {XmlToken[]} tokens
 * @returns {number}
 */
function scanStartTag(text, start, tokens) {
	let i = start + 1
	const nameStart = i
	while (i < text.length && !/[\s/>]/.test(text.charAt(i))) i++
	const name = text.slice(nameStart, i)
	if (!name) throw new XmlSyntaxError('start tag with no name', start)

	/** @type {XmlAttr[]} */
	const attrs = []
	for (;;) {
		const gapStart = i
		while (i < text.length && /\s/.test(text.charAt(i))) i++
		const gap = text.slice(gapStart, i)
		if (i >= text.length) throw new XmlSyntaxError('unterminated start tag', start)

		if (text.charAt(i) === '>') return pushOpen(tokens, name, attrs, gap, false, i + 1)
		if (text.startsWith('/>', i)) return pushOpen(tokens, name, attrs, gap, true, i + 2)

		const attrNameStart = i
		while (i < text.length && !/[\s=/>]/.test(text.charAt(i))) i++
		const attrName = text.slice(attrNameStart, i)
		if (!attrName) throw new XmlSyntaxError('malformed attribute', i)

		while (i < text.length && /\s/.test(text.charAt(i))) i++
		// An attribute with no value is HTML, not XML; refuse rather than invent one.
		if (text.charAt(i) !== '=') throw new XmlSyntaxError('attribute "' + attrName + '" has no value', i)
		i++
		while (i < text.length && /\s/.test(text.charAt(i))) i++

		const quote = text.charAt(i)
		if (quote !== '"' && quote !== "'") throw new XmlSyntaxError('unquoted attribute value', i)
		const valueEnd = text.indexOf(quote, i + 1)
		if (valueEnd === -1) throw new XmlSyntaxError('unterminated attribute value', i)
		attrs.push({ name: attrName, gap, value: text.slice(i + 1, valueEnd), quote })
		i = valueEnd + 1
	}
}

/**
 * @param {XmlToken[]} tokens
 * @param {string} name
 * @param {XmlAttr[]} attrs
 * @param {string} endGap whitespace between the last attribute and the closing delimiter
 * @param {boolean} selfClosing
 * @param {number} next offset just past the tag
 * @returns {number} `next`
 */
function pushOpen(tokens, name, attrs, endGap, selfClosing, next) {
	tokens.push({ kind: 'open', name, attrs, endGap, selfClosing })
	return next
}

/**
 * Build a node tree from a token stream.
 *
 * Returns the document's top-level children (declaration excluded) plus the declaration's
 * raw text, so the prolog is compared as bytes rather than reconstructed.
 * @param {XmlToken[]} tokens
 * @returns {{decl: string | null, children: XmlNode[]}}
 */
export function buildTree(tokens) {
	/** @type {XmlNode[]} */
	const roots = []
	/** @type {ElementNode[]} */
	const stack = []
	let decl = null

	for (const token of tokens) {
		const top = stack[stack.length - 1]
		const into = top ? top.children : roots
		if (token.kind === 'decl') {
			if (decl !== null || roots.length > 0 || stack.length > 0)
				throw new XmlSyntaxError('XML declaration is not the first thing in the document', 0)
			decl = token.raw
		} else if (token.kind === 'text') {
			into.push({ type: 'text', raw: token.raw })
		} else if (token.kind === 'open') {
			/** @type {ElementNode} */
			const node = {
				type: 'element',
				name: token.name,
				attrs: token.attrs,
				endGap: token.endGap,
				selfClosing: token.selfClosing,
				children: [],
			}
			into.push(node)
			if (!token.selfClosing) stack.push(node)
		} else {
			const open = stack.pop()
			if (!open) throw new XmlSyntaxError('end tag </' + token.name + '> with no open element', 0)
			if (open.name !== token.name)
				throw new XmlSyntaxError('end tag </' + token.name + '> closes <' + open.name + '>', 0)
		}
	}

	const unclosed = stack[stack.length - 1]
	if (unclosed) throw new XmlSyntaxError('unclosed element <' + unclosed.name + '>', 0)
	return { decl, children: roots }
}

/**
 * May this element's direct text children be treated as inert layout?
 *
 * Only when whitespace there cannot possibly be content, which takes all three:
 *
 * - it has at least one **element child** — an element whose only children are text is a
 *   leaf carrying a value, and `<a:t>  </a:t>` (a run of two literal spaces) is a
 *   perfectly ordinary thing for a deck to contain;
 * - it has no **non-whitespace** direct text child — anything else is mixed content, where
 *   the spaces around the text are part of the text;
 * - its name is not in {@link TEXT_BEARING}.
 *
 * The check is run against both sides and frozen if *either* says so, so a change cannot
 * unfreeze an element by removing the very text that made it text-bearing.
 * @param {ElementNode} node
 * @returns {boolean}
 */
export function isTextFrozen(node) {
	if (TEXT_BEARING.has(node.name)) return true
	const hasElementChild = node.children.some((child) => child.type === 'element')
	if (!hasElementChild) return true
	return node.children.some((child) => child.type === 'text' && !WHITESPACE.test(child.raw))
}

/**
 * @typedef {{path: string, from: string, to: string}} Relaxation
 */

/**
 * A proof. `ok` is spelled as a `@property` with a literal type rather than inline in an object
 * typedef, because only this form makes it a discriminant a caller can narrow the union on.
 * @typedef {object} ProofPass
 * @property {true} ok
 * @property {Relaxation[]} relaxations the whitespace positions that differed
 */

/**
 * A refusal, naming the first difference that was not inert whitespace.
 * @typedef {object} ProofFail
 * @property {false} ok
 * @property {string} path element path to the difference
 * @property {string} reason what differed
 */

/**
 * Prove that `current` differs from `base` only in inert inter-element whitespace.
 *
 * A pass means: identical element tree, identical attribute names, raw values, quote
 * characters and order, identical self-closing forms, identical character data byte for
 * byte, and every difference accounted for by a whitespace-only text node in a position
 * where whitespace cannot be content. Anything else is a fail naming the path and the
 * reason. Throws {@link XmlSyntaxError} if either side will not tokenize — a part that
 * cannot be read has not been proved equivalent to anything.
 * @param {string} base
 * @param {string} current
 * @returns {ProofPass | ProofFail}
 */
export function proveWhitespaceOnly(base, current) {
	const a = buildTree(tokenizeXml(base))
	const b = buildTree(tokenizeXml(current))

	if (a.decl !== b.decl)
		return {
			ok: false,
			path: '/',
			reason: 'XML declaration changed: ' + JSON.stringify(a.decl) + ' -> ' + JSON.stringify(b.decl),
		}

	/** @type {Relaxation[]} */
	const relaxations = []
	// The document itself is element-only by definition: whitespace outside the root element
	// is not content in any XML document, so the top level relaxes unconditionally.
	const fail = compareChildren(a.children, b.children, '', false, relaxations)
	return fail ?? { ok: true, relaxations }
}

/**
 * @param {XmlNode[]} baseKids
 * @param {XmlNode[]} curKids
 * @param {string} path
 * @param {boolean} frozen text children must match exactly
 * @param {Relaxation[]} relaxations
 * @returns {ProofFail | null}
 */
function compareChildren(baseKids, curKids, path, frozen, relaxations) {
	if (frozen) {
		if (baseKids.length !== curKids.length)
			return { ok: false, path, reason: 'child count changed: ' + baseKids.length + ' -> ' + curKids.length }
		for (const [i, baseKid] of baseKids.entries()) {
			const curKid = curKids[i]
			if (!curKid) return { ok: false, path, reason: 'child ' + i + ' vanished' }
			const fail = compareNode(baseKid, curKid, path, i, true, relaxations)
			if (fail) return fail
		}
		return null
	}

	// Not frozen: whitespace-only text children are layout. Drop them from both sides and
	// require what is left to correspond exactly — which is what makes an inserted element,
	// a removed one or a reordered one still a failure.
	const baseKept = baseKids.filter((child) => !isInertText(child))
	const curKept = curKids.filter((child) => !isInertText(child))
	const baseWs = baseKids.filter(isInertText).map((child) => /** @type {TextNode} */ (child).raw)
	const curWs = curKids.filter(isInertText).map((child) => /** @type {TextNode} */ (child).raw)
	if (baseWs.join('|') !== curWs.join('|'))
		relaxations.push({ path: path || '/', from: JSON.stringify(baseWs.join('|')), to: JSON.stringify(curWs.join('|')) })

	if (baseKept.length !== curKept.length)
		return {
			ok: false,
			path,
			reason: 'non-whitespace child count changed: ' + baseKept.length + ' -> ' + curKept.length,
		}
	for (const [i, baseKid] of baseKept.entries()) {
		const curKid = curKept[i]
		if (!curKid) return { ok: false, path, reason: 'child ' + i + ' vanished' }
		const fail = compareNode(baseKid, curKid, path, i, false, relaxations)
		if (fail) return fail
	}
	return null
}

/**
 * @param {XmlNode} child
 * @returns {boolean}
 */
function isInertText(child) {
	return child.type === 'text' && WHITESPACE.test(child.raw)
}

/**
 * @param {XmlNode} baseNode
 * @param {XmlNode} curNode
 * @param {string} path
 * @param {number} index
 * @param {boolean} frozen
 * @param {Relaxation[]} relaxations
 * @returns {ProofFail | null}
 */
function compareNode(baseNode, curNode, path, index, frozen, relaxations) {
	if (baseNode.type !== curNode.type)
		return { ok: false, path: path + '/[' + index + ']', reason: baseNode.type + ' became ' + curNode.type }

	if (baseNode.type === 'text') {
		const cur = /** @type {TextNode} */ (curNode)
		if (baseNode.raw !== cur.raw)
			return {
				ok: false,
				path: path + '/text()[' + index + ']',
				reason: 'character data changed: ' + JSON.stringify(baseNode.raw) + ' -> ' + JSON.stringify(cur.raw),
			}
		return null
	}

	const cur = /** @type {ElementNode} */ (curNode)
	const here = path + '/' + baseNode.name + '[' + index + ']'
	if (baseNode.name !== cur.name) return { ok: false, path: here, reason: 'element renamed to ' + cur.name }
	if (baseNode.selfClosing !== cur.selfClosing)
		return {
			ok: false,
			path: here,
			reason: baseNode.selfClosing ? '<x/> became <x></x>' : '<x></x> became <x/>',
		}

	const attrFail = compareAttrs(baseNode.attrs, cur.attrs, here)
	if (attrFail) return attrFail
	// Whitespace *inside* a start tag is inert too, and is deliberately NOT relaxed: this
	// module proves a claim about whitespace between elements, and `chart-parts.ts` has a
	// known double-space before `b=` (an empty `sz` interpolation) that must stay visible
	// rather than be absorbed by a gate aimed somewhere else.
	if (baseNode.endGap !== cur.endGap)
		return {
			ok: false,
			path: here,
			reason:
				'whitespace before the closing delimiter changed: ' +
				JSON.stringify(baseNode.endGap) +
				' -> ' +
				JSON.stringify(cur.endGap),
		}

	// `frozen` is inherited: once inside character data, nothing below relaxes.
	const childFrozen = frozen || isTextFrozen(baseNode) || isTextFrozen(cur)
	return compareChildren(baseNode.children, cur.children, here, childFrozen, relaxations)
}

/**
 * @param {XmlAttr[]} baseAttrs
 * @param {XmlAttr[]} curAttrs
 * @param {string} path
 * @returns {ProofFail | null}
 */
function compareAttrs(baseAttrs, curAttrs, path) {
	if (baseAttrs.length !== curAttrs.length)
		return {
			ok: false,
			path,
			reason:
				'attribute count changed: [' +
				baseAttrs.map((attr) => attr.name).join(' ') +
				'] -> [' +
				curAttrs.map((attr) => attr.name).join(' ') +
				']',
		}
	for (const [i, before] of baseAttrs.entries()) {
		const after = curAttrs[i]
		if (!after) return { ok: false, path, reason: 'attribute ' + i + ' vanished' }
		if (before.name !== after.name)
			return { ok: false, path, reason: 'attribute ' + i + ' is ' + after.name + ', was ' + before.name }
		if (before.value !== after.value)
			return {
				ok: false,
				path,
				reason: '@' + before.name + ' changed: ' + JSON.stringify(before.value) + ' -> ' + JSON.stringify(after.value),
			}
		if (before.quote !== after.quote) return { ok: false, path, reason: '@' + before.name + ' quote character changed' }
		if (before.gap !== after.gap)
			return {
				ok: false,
				path,
				reason:
					'whitespace before @' +
					before.name +
					' changed: ' +
					JSON.stringify(before.gap) +
					' -> ' +
					JSON.stringify(after.gap),
			}
	}
	return null
}
