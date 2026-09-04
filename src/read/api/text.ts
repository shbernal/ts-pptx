/**
 * Read/write proxies for a shape's text: `TextFrame → Paragraph[] → Run[]`.
 *
 * Each proxy wraps a live DOM element (`a:txBody`, `a:p`, `a:r`) and holds the
 * owning `Part`, so a setter can mutate the node in place and call
 * `part.markDirty()` — that single flag is what makes `save()` reserialize the
 * part. Getters compute from the DOM on each access rather than caching.
 *
 * The three proxies share no state, so they are three files under `text/`, plus the two free
 * body mutators `TableCell` and `DiagramPoint` also reach for. This module is the barrel: it
 * keeps the import surface every other module already uses.
 */
export {
	Run,
	type BulletDetail,
	type BulletStyle,
	type LineSpacing,
	type PlaceholderTextContext,
	type RunHyperlink,
} from './text/run.js'
export { Paragraph } from './text/paragraph.js'
export { TextFrame, type AutofitMode, type BodyProperties } from './text/frame.js'
export { setParagraphText, setTextBodyText } from './text/edit.js'
