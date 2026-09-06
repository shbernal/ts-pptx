/**
 * Keep the corpora's shared data out of the libraries' reach.
 *
 * Both corpora declare a handful of values once and hand the same objects to arm after arm:
 * a chart series, a table's rows, a data URL. That is deliberate, because the page prints
 * those declarations beside the code that used them. It is also a hazard, because a library
 * may edit what it is given. pptxgenjs does: `addChart` wraps a series' `labels` in an array
 * and stamps a `_dataIndex` on it, in place, on the caller's own object. ts-pptx leaves the
 * argument alone.
 *
 * Two things go wrong when nothing guards against that, and both were live before this
 * module existed:
 *
 *   - **The page printed the library's internals as though the corpus had written them.**
 *     The declarations are rendered from the live values, so a run that had already built a
 *     chart published `labels: [['Q1', 'Q2', 'Q3']], _dataIndex: 0` as the corpus's own
 *     source. Which is also to say the snapshot depended on how many arms had run first.
 *   - **A later arm was measured on what an earlier one left behind.** Every arm has to see
 *     the data the corpus declares, not the data the previous library edited, or the
 *     measurement is about the order of the loop.
 *
 * So the rendered copy is cloned at load, before any arm runs, and {@link reset} puts the
 * live values back between arms. Strings and other immutable values are passed over: there
 * is nothing for a library to edit in them.
 */

/**
 * A corpus's shared values, split into the copy the page renders and a way back to it.
 * @template {Record<string, unknown>} T
 * @param {T} values - the live constants, exactly as the arms name them
 * @returns {{constants: T, reset: () => void}} `constants` for rendering, `reset` for between arms
 */
export function corpusData(values) {
	const pristine = structuredClone(values)
	return {
		constants: /** @type {T} */ (structuredClone(pristine)),
		reset() {
			for (const [name, value] of Object.entries(pristine)) {
				const live = values[name]
				if (!Array.isArray(live) || !Array.isArray(value)) continue
				live.length = 0
				live.push(...structuredClone(value))
			}
		},
	}
}
