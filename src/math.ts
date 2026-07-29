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
	const oMath = mathmlToOmml(mathml)
	return display ? `<m:oMathPara>${DISPLAY_PARA_PR}${oMath}</m:oMathPara>` : oMath
}

// Error taxonomy — every failure the library throws. The classes and their `code` are API;
// the message is not. Re-exported from every entry so `instanceof` works whichever subpath a
// consumer imports — they all resolve to one shared module, so the classes are identical.
export {
	TsPptxError,
	InvalidOptionError,
	UnsupportedFeatureError,
	PackageReadError,
	MediaError,
	InternalError,
	type TsPptxErrorOptions,
} from './errors.js'
export type {
	ErrorCode,
	TsPptxCode,
	InvalidOptionErrorCode,
	UnsupportedFeatureErrorCode,
	PackageReadErrorCode,
	MediaErrorCode,
	InternalErrorCode,
} from './codes.js'
