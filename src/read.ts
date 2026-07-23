/**
 * `pptxgenjs/read` — open an existing `.pptx`, inspect its OPC structure, and
 * save it back with untouched parts byte-identical (lossless round-trip).
 *
 * This subsystem is isomorphic: bytes in, bytes out, no `node:fs`. File I/O
 * is the caller's job.
 */

// OPC layer — low-level package, parts, content types, relationships.
export { OpcPackage, type OpcInput } from './read/opc/package.js'
export { Part } from './read/opc/part.js'
export { ContentTypes } from './read/opc/content-types.js'
export { Relationships, type Relationship } from './read/opc/relationships.js'
export { resolveRelativePartName, relsPartNameFor } from './read/opc/partnames.js'

// Read object model — navigable presentation → slides → shapes → text.
export { Presentation } from './read/api/presentation.js'
export type {
	SlideSize,
	ImportSlideOptions,
	ImportShapeOptions,
	ImportSlideMastersOptions,
	ImportedSlideMaster,
	LayoutHandle,
	AppendSlidesOptions,
	FromTemplateOptions,
	SlideSource,
	ExtractedSlide,
	ExtractedSlides,
} from './read/api/presentation-types.js'
export { Slide, type AddTextBoxOptions, type AddPictureOptions } from './read/api/slide.js'
export { NotesSlide, NotesPlaceholder } from './read/api/notes.js'
export { type SlideBackground, type BackgroundSource } from './read/api/slide-background.js'
export { type TransitionInfo, type TransitionInput, type TransitionSpeed } from './read/api/transition.js'
export {
	Shape,
	AutoShape,
	Picture,
	Connector,
	GraphicFrame,
	GroupShape,
	isAutoShape,
	isPicture,
	isConnector,
	isGraphicFrame,
	isGroupShape,
	type AnyShape,
	type ShapeType,
	type GradientStop,
	type GradientFill,
	type LineEnd,
	type LineEnds,
	type ConnectionSite,
	type OuterShadow,
	type InnerShadow,
	type Glow,
	type Reflection,
	type SoftEdge,
	type PatternFill,
	type CustomGeometry,
	type CustomGeometryPath,
	type GeometryCommand,
	type AbsoluteFrame,
	type Recolor,
	type RecolorColor,
} from './read/api/shapes.js'
export {
	TextFrame,
	Paragraph,
	Run,
	type BodyProperties,
	type AutofitMode,
	type RunHyperlink,
	type LineSpacing,
} from './read/api/text.js'
export { Table, TableRow, TableCell, type CellBorder, type CellBorders } from './read/api/table.js'
export {
	Chart,
	ChartSeries,
	ChartAxis,
	type AxisNumberFormat,
	type ChartLegend,
	type ChartDataLabels,
	type ChartFill,
	type ChartLine,
} from './read/api/chart.js'
export { ChartEx, ChartExSeries, ChartExAxis, type ChartExLegend, type ChartExDataLabels } from './read/api/chartex.js'

// Theme colour resolution (schemeClr → literal hex) shared by the colour getters.
export { type ResolvedColor, resolveColorElement } from './read/api/theme-context.js'
export type { ColorContext } from './read/oxml/theme.js'
// DrawingML colour-transform application (base hex + transforms → effective hex).
export { applyColorTransforms, type ColorTransform, type EffectiveColor } from './read/oxml/color-transform.js'
