/**
 * The package names a chart's parts are written under, from its package-wide id.
 *
 * Four modules spelled them: the chart definer's placeholder, the write-time pass in
 * `package/assemble.ts` that assigns the authoritative id, the chart part contributor's content-type
 * overrides, and the part writer in `embed-xlsx.ts`. A name that drifted in one of them would write a
 * part no relationship or content type names.
 *
 * Imports nothing, so `package/assemble.ts`, which every write reaches, can use it without bringing
 * chart code into a deck that has no chart.
 */

/** The names one chart's parts take. */
export interface ChartPartNames {
	/** The chart part's file name, `chart{N}.xml` or `chartEx{N}.xml`. */
	fileName: string
	/** The chart part's absolute package path, as a slide relationship and a content type name it. */
	target: string
	/** A chartEx chart's style sidecar, `style{N}.xml`, beside the chart part. */
	styleName: string
	/** A chartEx chart's colors sidecar, `colors{N}.xml`, beside the chart part. */
	colorsName: string
	/** The embedded workbook's file name, in `ppt/embeddings/`. */
	workbookName: string
}

/**
 * The names for the chart with package-wide id `globalId`.
 * @param globalId - the chart's id, unique across every chart part in the package
 * @param isChartEx - whether it is a chartEx chart, which uses the `chartEx{N}.xml` name family;
 *   absent means a classic chart, as `SlideRelChart.isChartEx` does
 */
export function chartPartNames(globalId: number, isChartEx?: boolean): ChartPartNames {
	const base = isChartEx ? `chartEx${globalId}` : `chart${globalId}`
	return {
		fileName: `${base}.xml`,
		target: `/ppt/charts/${base}.xml`,
		styleName: `style${globalId}.xml`,
		colorsName: `colors${globalId}.xml`,
		workbookName: `Microsoft_Excel_Worksheet${globalId}.xlsx`,
	}
}
