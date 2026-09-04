/**
 * The `TableRow` read/edit proxy (`a:tr`).
 *
 * A row is a thin walk: its cells, its height, and the two escape hatches. It is its own file
 * only because it sits between the other two in the chain `Table -> TableRow -> TableCell`,
 * and a chain is what keeps the three modules acyclic.
 */
import type { Part } from '../opc/part.js'
import type { Relationships } from '../opc/relationships.js'
import { attr, type Element, getElements, numberValue } from '../oxml/dom.js'
import type { ThemeContext } from '../oxml/theme.js'
import { TableCell, type TableCellStyleContext } from './table-cell.js'
/** One table row (`a:tr`). */
export class TableRow {
	constructor(
		private readonly tr: Element,
		private readonly part: Part,
		/** The owning slide's theme colour context, threaded to each {@link TableCell}. */
		private readonly themeContext?: ThemeContext,
		/** The table's style-resolution context, threaded to each cell for {@link TableCell.resolvedFill}. */
		private readonly style?: TableCellStyleContext | null,
		/** This row's zero-based index in the table, for style-graph banding/edge conditions. */
		private readonly rowIndex = 0,
		/** The owning slide's relationships, threaded to each cell for {@link TableCell.pictureFill}. */
		private readonly rels?: Relationships
	) {}

	/** The row's cells (`a:tc`) in left-to-right order. */
	get cells(): TableCell[] {
		return getElements(this.tr, 'a:tc').map(
			(tc, colIndex) => new TableCell(tc, this.part, this.themeContext, this.style, this.rowIndex, colIndex, this.rels)
		)
	}

	/** Row height in EMU (`a:tr/@h`), or `null` if unset. */
	get heightEmu(): number | null {
		return numberValue(attr(this.tr, 'h'))
	}

	/** Escape hatch: the underlying `a:tr` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.tr
	}

	/** Mark the owning part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}
}
