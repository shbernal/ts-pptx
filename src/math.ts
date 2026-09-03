/**
 * `@shbernal/ts-pptx/math` — author display equations in LaTeX or MathML and get
 * OMML for the `math:` option on `addText` (see {@link TextProps.math}).
 *
 * Pipeline: LaTeX --temml--> MathML --mathml2omml--> OMML.
 *
 * Both converters are **optional peer dependencies** (temml is MIT, mathml2omml is
 * LGPL-3.0-or-later); install them to use this subpath:
 *
 * ```sh
 * npm install temml mathml2omml
 * ```
 *
 * This subpath is **Node-only**: it loads the converters synchronously via
 * `node:module`'s `createRequire`, so `latexToOmml()` / `mathmlToOmml()` stay
 * synchronous. Both display and inline (in-sentence) equations are supported: a
 * default `latexToOmml()` returns a centered display `<m:oMathPara>`, while
 * `latexToOmml(tex, { display: false })` (and `mathmlToOmml()`) return the bare
 * `<m:oMath>` form you pass to a text item with `inline: true` to flow it
 * mid-paragraph (see {@link TextProps.inline}).
 *
 * @module
 */
import { createRequire } from 'node:module'
import { InvalidOptionError, UnsupportedFeatureError } from './errors.js'

const nodeRequire = createRequire(import.meta.url)

/** Minimal shape of the temml API we depend on. */
interface TemmlModule {
	renderToString(latex: string, options?: { displayMode?: boolean; throwOnError?: boolean }): string
}
/** temml's `ParseError` carries a `position` (character offset into the source). */
interface TemmlParseError extends Error {
	position?: number
}

const MISSING_DEPS_MESSAGE =
	"@shbernal/ts-pptx/math requires the optional peer dependencies 'temml' and " +
	"'mathml2omml'. Install them with: npm install temml mathml2omml"

/** Require an optional peer, remapping a not-found error to the install hint. */
function requirePeer<T>(id: string): T {
	try {
		return nodeRequire(id) as T
	} catch (err) {
		if ((err as NodeJS.ErrnoException | undefined)?.code === 'MODULE_NOT_FOUND') {
			throw new UnsupportedFeatureError('math/missing-optional-peer', MISSING_DEPS_MESSAGE, { cause: err })
		}
		throw err
	}
}

let temmlModule: TemmlModule | undefined
let mml2ommlFn: ((mathml: string) => string) | undefined

function loadTemml(): TemmlModule {
	if (!temmlModule) {
		const mod = requirePeer<TemmlModule | { default: TemmlModule }>('temml')
		temmlModule = 'renderToString' in mod ? mod : mod.default
	}
	return temmlModule
}

function loadMml2omml(): (mathml: string) => string {
	if (!mml2ommlFn) {
		const mod = requirePeer<{ mml2omml: (mathml: string) => string }>('mathml2omml')
		mml2ommlFn = mod.mml2omml
	}
	return mml2ommlFn
}

/**
 * mathml2omml puts namespace declarations on the `<m:oMath>` opening tag: `xmlns:m`
 * (redundant — the `<a14:m>` envelope the `math:` option authors already declares it)
 * and `xmlns:w` (unused in the output it produces). Strip them so the result is a bare
 * `<m:oMath>` that drops into the `math:` option cleanly.
 */
function stripOMathNamespaces(omml: string): string {
	return omml.replace(/^(\s*<m:oMath)\b[^>]*>/, '$1>')
}

/** Centered-display paragraph props, matching what the `math:` emitter/oracle author. */
const DISPLAY_PARA_PR = '<m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr>'

/**
 * The operator-dictionary subset this module has to carry, keyed by the character temml
 * puts in the `<mo>` of an accent and valued by the combining mark OMML wants in `m:chr`.
 *
 * Two halves of one problem. **Which** `<mover>`s are accents: temml renders every accent
 * command as a bare `<mover>` with no `accent="true"`, which is correct for a browser —
 * MathML renderers derive accent positioning from the operator dictionary, so the attribute
 * is redundant there — but mathml2omml has no dictionary and keys strictly off the
 * attribute, mapping a bare `<mover>` to `<m:limUpp>` (an over-*limit*, with limit spacing)
 * instead of `<m:acc>`. Membership in this table is the dictionary lookup that is otherwise
 * missing. **Which character** to attach: mathml2omml passes the `<mo>` text straight
 * through to `m:chr`, and temml emits the *spacing* modifiers (U+02C6 MODIFIER LETTER
 * CIRCUMFLEX, U+2192 RIGHTWARDS ARROW) while ECMA-376 §22.1.2.20 says an `accPr` character
 * "should be within the range of (U+0300–U+036F) or (U+20D0–U+20EF)" — the combining marks,
 * which is also what Word itself writes. Left alone, `\vec{v}` would hang a full-size arrow
 * over the base instead of an arrow accent.
 *
 * Keyed by character rather than by LaTeX command because that is all that survives into
 * the MathML. Derived by rendering every accent command temml documents and reading back
 * the `<mo>` it produced; the commands each entry serves are named in the comments.
 */
const ACCENT_CHARS = new Map<string, string>([
	['ˆ', '̂'], // ˆ  \hat, \^         → COMBINING CIRCUMFLEX ACCENT
	['~', '̃'], // ~  \tilde, \~       → COMBINING TILDE
	['´', '́'], // ´  \acute           → COMBINING ACUTE ACCENT
	['ˊ', '́'], // ˊ  \'               → COMBINING ACUTE ACCENT
	['`', '̀'], // `  \grave           → COMBINING GRAVE ACCENT
	['ˋ', '̀'], // ˋ  \`               → COMBINING GRAVE ACCENT
	['¨', '̈'], // ¨  \ddot, \"        → COMBINING DIAERESIS
	['˙', '̇'], // ˙  \dot, \.         → COMBINING DOT ABOVE
	['‾', '̄'], // ‾  \bar             → COMBINING MACRON
	['ˉ', '̄'], // ˉ  \=               → COMBINING MACRON
	['˘', '̆'], // ˘  \breve, \u       → COMBINING BREVE
	['ˇ', '̌'], // ˇ  \check, \v       → COMBINING CARON
	['˚', '̊'], // ˚  \mathring, \r    → COMBINING RING ABOVE
	['˝', '̋'], // ˝  \H               → COMBINING DOUBLE ACUTE ACCENT
	['→', '⃗'], // →  \vec             → COMBINING RIGHT ARROW ABOVE
	['…', '⃛'], // …  \dddot           → COMBINING THREE DOTS ABOVE
])

/** An `<mover>`'s last child, when that child is an `<mo>`: `[, attributes, text]`. */
const TRAILING_MO = /<mo\b([^>]*)>([^<]*)<\/mo>\s*$/

/**
 * Stamp `accent="true"` on the `<mover>`s that are accents, and swap their operator for the
 * combining mark OMML wants. See {@link ACCENT_CHARS} for why both are needed.
 *
 * Scoped as narrowly as the defect. Only `<mo stretchy="false">` qualifies, which is how
 * temml distinguishes a true accent from a *wide* one (`\widehat`, `\overrightarrow`) or a
 * group character (`\overbrace`, `\overgroup`) — those already reach `m:groupChr`, which is
 * a reasonable rendering, so they are left alone rather than churned. `\stackrel` is an
 * over-relation whose character comes from the caller, so it stays `m:limUpp` unless the
 * caller happened to stack a diacritic. An `<mover>` that already states `accent` is the
 * author's call and is never rewritten.
 *
 * `<munder>` is deliberately **not** touched. The symmetric-looking fix — stamping
 * `accentunder="true"` — makes mathml2omml emit `<m:acc>`, which is an *over*-accent: it
 * would move `\utilde`'s tilde above the base. OMML has no under-accent object, so the
 * `m:limLow` an under-accent already produces is the closest thing available.
 *
 * Implemented against the tag text rather than through a DOM. temml's output is
 * machine-generated and well-formed, so pairing `<mover>` with `</mover>` on a stack is
 * exact, and this subpath otherwise carries no XML parser — the read path's is a different
 * dependency and pulling it in here would grow the module for one rewrite.
 */
function markAccentedMovers(mathml: string): string {
	const edits: { at: number; length: number; text: string }[] = []
	const open: number[] = []
	const tags = /<(\/?)mover\b[^>]*>/g
	for (let tag = tags.exec(mathml); tag !== null; tag = tags.exec(mathml)) {
		if (tag[1] === '') {
			// `accent` already stated: the author has answered the question this asks.
			open.push(tag[0].includes('accent=') ? -1 : tag.index)
			continue
		}
		const openAt = open.pop()
		if (openAt === undefined || openAt === -1) continue

		const openEnd = mathml.indexOf('>', openAt) + 1
		const mo = TRAILING_MO.exec(mathml.slice(openEnd, tag.index))
		if (!mo) continue
		const [whole, attrs = '', operator = ''] = mo
		// `stretchy="false"` is how temml separates a true accent from a wide one
		// (`\widehat`) or a group character (`\overbrace`); those already map well.
		if (!attrs.includes('stretchy="false"')) continue
		const combining = ACCENT_CHARS.get(operator)
		// A multi-character operator (`\ddddot` renders as "….") has no single `m:chr`, so it
		// is not in the table and keeps the limit form rather than producing an invalid one.
		if (combining === undefined) continue

		// Applied right-to-left below, so earlier offsets stay valid.
		edits.push({
			at: openEnd + mo.index + whole.length - `${operator}</mo>`.length,
			length: operator.length,
			text: combining,
		})
		edits.push({ at: openAt + '<mover'.length, length: 0, text: ' accent="true"' })
	}

	let out = mathml
	for (const edit of edits.sort((a, b) => b.at - a.at)) {
		out = out.slice(0, edit.at) + edit.text + out.slice(edit.at + edit.length)
	}
	return out
}

/**
 * Convert a MathML string to OMML for the `math:` option on `addText`.
 *
 * Canonical output form: a bare `<m:oMath>…</m:oMath>` in the `m:` namespace, with no
 * namespace declarations (the `math:` option's envelope supplies the `m` prefix).
 *
 * @param mathml - a MathML string (e.g. `<math>…</math>`)
 * @returns OMML `<m:oMath>…</m:oMath>`
 */
export function mathmlToOmml(mathml: string): string {
	const mml2omml = loadMml2omml()
	return stripOMathNamespaces(mml2omml(mathml).trim())
}

export interface LatexToOmmlOptions {
	/**
	 * Display (block) math: render in `displayMode` and wrap the result in
	 * `<m:oMathPara>` (a centered display-math paragraph). Default `true`. When `false`,
	 * temml renders in inline mode and the result is a bare `<m:oMath>`.
	 */
	display?: boolean
}

/**
 * Convert a LaTeX string to OMML for the `math:` option on `addText`.
 *
 * Display math (the default) is returned as a full `<m:oMathPara>` centered display
 * paragraph; with `{ display: false }` a bare `<m:oMath>` is returned (both forms are
 * accepted by the `math:` option). Throws on invalid LaTeX, surfacing temml's parse
 * position.
 *
 * @param latex - a LaTeX math expression, e.g. `x = \frac{-b \pm \sqrt{b^2-4ac}}{2a}`
 * @param opts - see {@link LatexToOmmlOptions}
 * @returns OMML `<m:oMathPara>…</m:oMathPara>` (display) or `<m:oMath>…</m:oMath>` (inline)
 * @throws {Error} if the LaTeX cannot be parsed
 */
export function latexToOmml(latex: string, opts: LatexToOmmlOptions = {}): string {
	const display = opts.display ?? true
	const temml = loadTemml()
	let mathml: string
	try {
		mathml = temml.renderToString(latex, { displayMode: display, throwOnError: true })
	} catch (err) {
		const pe = err as TemmlParseError
		const pos = typeof pe.position === 'number' ? ` (position ${pe.position})` : ''
		throw new InvalidOptionError('math/invalid-latex', `Invalid LaTeX${pos}: ${pe.message}`, { cause: err })
	}
	// Applied here rather than in `mathmlToOmml`, because it compensates for something temml
	// does: a caller handing in their own MathML states `accent` themselves, and rewriting it
	// under them would be the module second-guessing an attribute the MathML spec gives them.
	const oMath = mathmlToOmml(markAccentedMovers(mathml))
	return display ? `<m:oMathPara>${DISPLAY_PARA_PR}${oMath}</m:oMathPara>` : oMath
}

// Error taxonomy — see `entry-errors.ts`. Re-exported from every entry so `instanceof`
// works whichever subpath a consumer imports.
export * from './entry-errors.js'

// Diagnostics — see `entry-diagnostics.ts`. Re-exported from every entry so a consumer of any
// subpath can install a handler for the warnings that subpath emits.
export * from './entry-diagnostics.js'
