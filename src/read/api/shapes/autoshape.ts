/**
 * An auto shape, text box, or placeholder (`p:sp`) — the only shape kind that holds text.
 */

import { attr, firstChild } from '../../oxml/dom.js'
import { resolveStyleFontRef, type PlaceholderRef } from '../theme-context.js'
import { TextFrame } from '../text.js'
import { Shape } from './base.js'
import { nvPrOf } from '../../oxml/slide-dom.js'

/** An auto shape, text box, or placeholder (`p:sp`). The only kind that holds text. */
export class AutoShape extends Shape {
	readonly shapeType = 'autoShape' as const

	override get hasTextFrame(): boolean {
		return firstChild(this.element, 'p:txBody') !== null
	}

	override get textFrame(): TextFrame | null {
		const txBody = firstChild(this.element, 'p:txBody')
		if (!txBody) return null
		const ctx = this.host.themeContext()
		// Every shape resolves its `p:style/a:fontRef` and the presentation's
		// `p:defaultTextStyle`, so an inheritance is always supplied; `ph` is null for a
		// non-placeholder shape, which then skips only the layout/master placeholder tiers.
		const inherit = { ph: this.placeholder, fontRef: resolveStyleFontRef(this.element, ctx) }
		return new TextFrame(txBody, { part: this.host.part, ctx, rels: this.host.relationships, inherit })
	}

	/**
	 * This shape's placeholder identity (`p:ph` `type`/`idx`), or `null` when it is
	 * not a placeholder. `idx` defaults to `'0'` when the attribute is absent, as
	 * PowerPoint does. Use {@link Slide.placeholder} to find a placeholder by type.
	 */
	override get placeholder(): PlaceholderRef | null {
		const nvPr = nvPrOf(this.element)
		const ph = nvPr && firstChild(nvPr, 'p:ph')
		return ph ? { type: attr(ph, 'type'), idx: attr(ph, 'idx') ?? '0' } : null
	}
}
