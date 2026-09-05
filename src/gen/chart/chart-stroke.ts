/**
 * ts-pptx: folding a chart's older stroke spellings onto {@link StrokeProps}.
 *
 * A stroke is one concept, and the chart option bag used to spell it three ways that the rest
 * of the library did not: `size` for the width, `style` for the dash, and a separate
 * `*AxisLineShow` flag for the state `type: 'none'` already names. Those spellings are kept
 * (deprecated) rather than removed, so this is where they collapse into the one shape the
 * emitters read.
 *
 * The rule is the same at every site: **a new key wins over its old counterpart**, and an
 * explicit `type` wins over an inferred one. Nothing here validates -- `define/chart.ts` warns
 * on an out-of-range width or cap before this runs, and `resolveDash` catches a dash outside
 * `ST_PresetLineDashVal` after it.
 */

/* oxlint-disable typescript/no-deprecated -- reading the deprecated spellings is this module's whole job. */

import { warn } from '../../diagnostics.js'
import type {
	ChartPropsAxisCat,
	ChartPropsAxisSer,
	ChartPropsAxisVal,
	OptsChartGridLine,
	StrokeProps,
} from '../../types/index.js'

/** The axis-line half of the chart option bag, in all three axes' spellings. */
type AxisLineOpts = ChartPropsAxisCat & ChartPropsAxisSer & ChartPropsAxisVal

/**
 * Fold a gridline's `size` and `style` onto `width`, `dashType` and `type`.
 *
 * `style: 'none'` and `type: 'none'` are the same statement; `style: 'dash'` is the literal
 * `dash` preset and maps onto `dashType`, *not* onto `type: 'dash'`, which would mean
 * `sysDash` and would change every deck that already spelled it the old way.
 * @param gl - the caller's gridline options
 * @returns a copy with only the {@link StrokeProps} keys carrying meaning
 */
export function gridLineStroke(gl: OptsChartGridLine): OptsChartGridLine {
	const out: OptsChartGridLine = { ...gl }
	if (out.width === undefined && out.size !== undefined) out.width = out.size
	if (out.dashType === undefined && gl.type === undefined && out.style !== undefined && out.style !== 'none')
		out.dashType = out.style
	if (out.type === undefined && out.style === 'none') out.type = 'none'
	return out
}

/**
 * Whether a gridline (or a stacked bar's series lines) asks not to be drawn at all.
 *
 * Three axis builders and `createSerLinesElement` each spelled this as
 * `opts.style !== 'none'`, which stopped being the whole test once `type: 'none'` could say it
 * too.
 * @param gl - the caller's gridline options, if any
 * @returns true when no element should be emitted
 */
export function gridLineSuppressed(gl: OptsChartGridLine | undefined): boolean {
	if (!gl) return true
	const stroke = gridLineStroke(gl)
	return stroke.type === 'none'
}

/**
 * The one stroke an axis' line emitter reads, folded from whichever spelling the caller used.
 *
 * `*AxisLineShow: false` is the state `type: 'none'` names, and it is the only one of the four
 * flat keys that was not simply a different word for a {@link StrokeProps} key. What it emits
 * is `<a:noFill/>` on the axis' own `<c:spPr>` -- the axis line element is always present, so
 * "hidden" and "absent" are genuinely different here.
 *
 * All twelve deprecated reads live in this one function rather than at the three axis
 * builders, which is also what stops the three from drifting again: the series axis got its
 * width and dash hardcoded precisely because each builder spelled this out by hand.
 * @param opts - the chart option bag
 * @param axis - which axis' line to resolve
 * @returns the stroke to emit
 */
export function axisLineStroke(opts: AxisLineOpts, axis: 'cat' | 'val' | 'ser'): StrokeProps {
	const nested = opts[`${axis}AxisLine`]
	const color = opts[`${axis}AxisLineColor`]
	const size = opts[`${axis}AxisLineSize`]
	const style = opts[`${axis}AxisLineStyle`]
	const show = opts[`${axis}AxisLineShow`]

	const out: StrokeProps = { ...nested }
	if (out.color === undefined && color !== undefined) out.color = color
	if (out.width === undefined && size !== undefined) out.width = size
	// `show` and `style` are folded independently: a hidden axis line still writes its
	// `<a:prstDash>`, so letting the inferred `type: 'none'` swallow the dash would change
	// the bytes of every deck that set both.
	if (out.dashType === undefined && nested?.type === undefined && style !== undefined) out.dashType = style
	if (out.type === undefined && show === false) out.type = 'none'
	return out
}

/**
 * Scrub a gridline-shaped stroke's out-of-range values so a default applies instead of a
 * PowerPoint-invalid one reaching the part.
 *
 * `width` and `size` are the same measure under two names, so both are checked and both are
 * dropped together -- leaving one behind would let the scrubbed value come back through the
 * fold. `dashType` needs nothing here: `resolveDash` checks it against the whole
 * `ST_PresetLineDashVal` set at emit, which is the check `style` never had.
 *
 * `option` exists because `c:serLines` takes the same {@link OptsChartGridLine} shape through a
 * different option, and a diagnostic naming `chart.gridLine.width` for a `barSeriesLine` mistake
 * points the caller at an option they did not set. Every entry point that reaches
 * `chartFurnitureLine` is scrubbed here, which is also what lets that emitter default a missing
 * width with `??` -- an unscrubbed caller was the one path a zero could reach the part on.
 * @param glOpts - the caller's line options, scrubbed in place
 * @param option - the option path to name in a diagnostic, as the caller spells it
 */
export function scrubGridLine(glOpts: OptsChartGridLine, option = 'chart.gridLine'): void {
	if (!glOpts || gridLineSuppressed(glOpts)) return
	const width = glOpts.width ?? glOpts.size
	if (width !== undefined && (!Number.isFinite(Number(width)) || width <= 0)) {
		warn('chart/invalid-grid-line-size', `${option}.width must be greater than 0.`)
		delete glOpts.width // delete prop to used defaults
		delete glOpts.size
	}
	if (glOpts.style && !['solid', 'dash', 'dot'].includes(glOpts.style)) {
		warn('chart/invalid-grid-line-style', `${option}.style options: \`solid\`, \`dash\`, \`dot\`.`)
		delete glOpts.style
	}
	if (glOpts.cap && !['flat', 'square', 'round'].includes(glOpts.cap)) {
		warn('chart/invalid-grid-line-cap', `${option}.cap options: \`flat\`, \`square\`, \`round\`.`)
		delete glOpts.cap
	}
}
