/**
 * PptxGenJS: Chart Generation (barrel)
 *
 * Re-exports the chart package-part builders now split under `gen/chart/`:
 *   - `gen/chart/chart-xml.ts`   the chart.xml DrawingML (`makeXmlCharts` + fragments)
 *   - `gen/chart/embed-xlsx.ts`  the embedded `.xlsx` workbook backing the cached data
 *   - `gen/chart/data-refs.ts`   the series↔worksheet-cell mapping the two sides share
 *
 * Kept so existing `./gen-charts.js` import sites (the `pptxgen.ts` `genCharts.*`
 * namespace) keep resolving unchanged.
 */

export { makeXmlCharts } from './gen/chart/chart-xml.js'
export { buildEmbeddedWorksheet, buildChartRelsXml, createExcelWorksheet } from './gen/chart/embed-xlsx.js'
