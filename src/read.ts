/**
 * `pptxgenjs/read` — open an existing `.pptx`, inspect its OPC structure, and
 * save it back with untouched parts byte-identical (lossless round-trip).
 *
 * This subsystem is isomorphic: bytes in, bytes out, no `node:fs`. File I/O
 * is the caller's job.
 */

// OPC layer (Phase 1) — low-level package, parts, content types, relationships.
export { OpcPackage, type OpcInput } from './read/opc/package.js'
export { Part } from './read/opc/part.js'
export { ContentTypes } from './read/opc/content-types.js'
export { Relationships, type Relationship } from './read/opc/relationships.js'
export { resolveRelativePartName, relsPartNameFor } from './read/opc/partnames.js'

// Read object model (Phase 2) — navigable presentation → slides → shapes → text.
export {
	Presentation,
	type SlideSize,
	type ImportSlideOptions,
	type ImportShapeOptions,
	type ImportSlideMastersOptions,
	type ImportedSlideMaster,
	type LayoutHandle,
	type AppendSlidesOptions,
	type SlideSource,
	type ExtractedSlide,
	type ExtractedSlides,
} from './read/api/presentation.js'
export { Slide, type AddTextBoxOptions, type AddPictureOptions } from './read/api/slide.js'
export {
	Shape,
	AutoShape,
	Picture,
	Connector,
	GraphicFrame,
	GroupShape,
	type ShapeType,
	type GradientStop,
	type GradientFill,
	type LineEnd,
	type LineEnds,
	type OuterShadow,
	type CustomGeometry,
	type CustomGeometryPath,
	type GeometryCommand,
	type AbsoluteFrame,
	type Recolor,
	type RecolorColor,
} from './read/api/shapes.js'
export { TextFrame, Paragraph, Run, type BodyProperties } from './read/api/text.js'
export { Table, TableRow, TableCell } from './read/api/table.js'
export { Chart, ChartSeries } from './read/api/chart.js'

// Theme colour resolution (schemeClr → literal hex) shared by the colour getters.
export { type ResolvedColor, resolveColorElement } from './read/api/theme-context.js'
export type { ColorContext } from './read/oxml/theme.js'
// DrawingML colour-transform application (base hex + transforms → effective hex).
export { applyColorTransforms, type ColorTransform, type EffectiveColor } from './read/oxml/color-transform.js'
