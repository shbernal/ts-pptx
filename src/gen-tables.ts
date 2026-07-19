/**
 * PptxGenJS: Table Generation — barrel
 *
 * The auto-paging core lives in `gen/table/autopage.ts` (`getSlidesForTableRows`);
 * the browser-only tableToSlides() DOM path lives in `gen/table/html-dom.ts`. This
 * file re-exports both so existing `./gen-tables.js` import sites keep resolving.
 */

export { getSlidesForTableRows } from './gen/table/autopage.js'
export { htmlBorderToProps, resolveHtmlColWidth, genTableToSlides } from './gen/table/html-dom.js'
