/**
 * ts-pptx: `ppt/tableStyles.xml`
 *
 * Emit the table-styles part. The part carries only the default style id: it defines no
 * styles of its own, because PowerPoint never reads one out of it.
 *
 * `<a:tableStyleId>` is resolved against PowerPoint's *own* table-style gallery, not against
 * this part. Verified by render (PowerPoint desktop 16.0): a deck pointed at a built-in GUID
 * this part does not define paints correctly, while a PowerPoint-authored deck whose style
 * GUID is rewritten to a novel value — in both this part and the slide, bytes otherwise
 * identical — loses its styling entirely and falls back to the black hairline grid. The same
 * holds for a definition placed inline in the slide's `<a:tblPr>` as `<a:tableStyle>`, and for
 * one nominated by this part's `def=`. So a custom `<a:tblStyle>` is unreachable markup in the
 * supported target, whatever it says and wherever it sits. Brand styling belongs in direct
 * cell formatting (`headerRow`, `columns[i]`, per-cell options); see `docs/tables.md`.
 *
 * The part itself still ships: PowerPoint expects the relationship and content-type override
 * to be present, and the read side (`src/read/api/ops/table-styles.ts`) merges real definitions out of
 * *imported* decks into it, which is a different thing from authoring one here.
 */

import { TableStyle } from '../../enums.js'
import { CRLF, XML_DECL } from '../../constants-internal.js'
import { voidEl } from '../oxml/el.js'
import { OOXML_NS } from '../../ooxml/namespaces.js'

/**
 * Create `ppt/tableStyles.xml`
 * @see: http://openxmldeveloper.org/discussions/formats/f/13/p/2398/8107.aspx
 * @return {string} XML
 */
export function makeXmlTableStyles(): string {
	return (
		XML_DECL +
		CRLF +
		voidEl('a:tblStyleLst', {
			'xmlns:a': OOXML_NS.a,
			def: TableStyle.MEDIUM_STYLE_2_ACCENT_1,
		})
	)
}
