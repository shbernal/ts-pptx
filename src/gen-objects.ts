/**
 * PptxGenJS: Slide Object Generators (barrel)
 *
 * The `add*Definition` layer behind the public `Slide` methods (`slide.addText`, `addChart`,
 * `addTable`, …) now lives under `gen/define/`, one module per object kind. Each function
 * validates and normalizes caller options into an internal `SlideObject` and pushes it onto the
 * slide model (and registers media/hyperlink rels); the XML is emitted later by `gen-xml.ts`.
 *
 * This barrel re-exports the public entry points so existing `./gen-objects.js` import sites
 * (the `genObj.*` namespace in `pptxgen.ts` / `slide.ts`) keep resolving unchanged.
 *
 *   - gen/define/group.ts        addGroupDefinition / groupObjectsDefinition (+ child dispatch)
 *   - gen/define/master.ts       createSlideMaster
 *   - gen/define/chart.ts        addChartDefinition
 *   - gen/define/image.ts        addImageDefinition (+ image-fill media)
 *   - gen/define/media.ts        addMediaDefinition
 *   - gen/define/notes.ts        addNotesDefinition
 *   - gen/define/comment.ts      addCommentDefinition
 *   - gen/define/shape.ts        addShapeDefinition
 *   - gen/define/connector.ts    addConnectorDefinition
 *   - gen/define/table.ts        addTableDefinition
 *   - gen/define/text.ts         addTextDefinition
 *   - gen/define/placeholder.ts  addPlaceholdersToSlideLayouts
 *   - gen/define/background.ts   addBackgroundDefinition
 *   - gen/define/object-name.ts  nextObjectNameIdx (shared name counter)
 *   - gen/define/hyperlinks.ts   createHyperlinkRels (shared rel registration)
 */
export { addGroupDefinition, groupObjectsDefinition } from './gen/define/group.js'
export { createSlideMaster } from './gen/define/master.js'
export { addChartDefinition } from './gen/define/chart.js'
export { addImageDefinition } from './gen/define/image.js'
export { addMediaDefinition } from './gen/define/media.js'
export { addNotesDefinition } from './gen/define/notes.js'
export { addCommentDefinition } from './gen/define/comment.js'
export { addShapeDefinition } from './gen/define/shape.js'
export { addConnectorDefinition } from './gen/define/connector.js'
export { addTableDefinition } from './gen/define/table.js'
export { addTextDefinition } from './gen/define/text.js'
export { addPlaceholdersToSlideLayouts } from './gen/define/placeholder.js'
export { addBackgroundDefinition } from './gen/define/background.js'
