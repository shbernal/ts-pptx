/**
 * Object-identity types: placeholders, object names/alt text and DrawingML lock flags.
 *
 * Re-exported by `../core-interfaces.js`, which is the import site for the rest of `src/`.
 */
import type { PLACEHOLDER_TYPE, SHAPE_NAME } from '../core-enums.js'
import type { Margin, PositionProps } from './core.js'
import type { TextBaseProps } from './text.js'

export interface PlaceholderProps extends PositionProps, TextBaseProps, ObjectNameProps {
	name: string
	type: PLACEHOLDER_TYPE
	/**
	 * margin (inches) — text-frame internal margin; a value `>= 1` warns as a likely legacy points value
	 */
	margin?: Margin
	/**
	 * Preset shape geometry for this placeholder (e.g. `'roundRect'`)
	 * @default 'rect'
	 */
	shape?: SHAPE_NAME
	/**
	 * Rounded rectangle corner radius (inches) when `shape: 'roundRect'`
	 * - range: 0.0 to slide height/2
	 */
	rectRadius?: number
}
export interface ObjectNameProps {
	/**
	 * Object name
	 * - used instead of default "Object N" name
	 * - PowerPoint: Home > Arrange > Selection Pane...
	 * @default 'Object 1'
	 * @example 'Antenna Design 9'
	 */
	objectName?: string
	/**
	 * Alt Text value ("How would you describe this object and its contents to someone who is blind?")
	 * - serialized to the generated object's `p:cNvPr` `descr` attribute
	 * - PowerPoint: [right-click on the object] > "Edit Alt Text..."
	 * @example 'Quarterly revenue bar chart'
	 */
	altText?: string
	/**
	 * Object lock flags (DrawingML `a:spLocks` / `a:picLocks` / `a:graphicFrameLocks`)
	 * - restrict how the object can be manipulated in PowerPoint (e.g. prevent moving, resizing, or grouping)
	 * - each flag maps 1:1 to the OOXML attribute of the same name; only flags set to `true` are emitted
	 * - PowerPoint UI: Selection Pane / right-click protections (most locks are honored at edit time, not as a password)
	 * - flags only apply to the object types that support them (see each flag); flags set on an unsupported
	 *   object type are ignored with a console warning
	 * @example { noMove: true, noResize: true } // pin an object in place
	 * @example { noGrp: true } // exclude from grouping
	 */
	objectLock?: ObjectLockProps
}
/**
 * Object lock flags. Maps to DrawingML locking elements:
 * - shapes / text boxes / placeholders → `a:spLocks`
 * - images / media → `a:picLocks`
 * - tables → `a:graphicFrameLocks`
 *
 * Each property mirrors the OOXML attribute name. A flag is only serialized for object types whose
 * locking element defines it (noted per-flag); setting an unsupported flag logs a warning and is ignored.
 */
export interface ObjectLockProps {
	/** Disallow grouping/ungrouping with other objects. (shapes, images, tables) */
	noGrp?: boolean
	/** Disallow selecting the object. (shapes, images, tables) */
	noSelect?: boolean
	/** Disallow moving the object. (shapes, images, tables) */
	noMove?: boolean
	/** Disallow resizing the object. (shapes, images, tables) */
	noResize?: boolean
	/** Disallow changing the aspect ratio. (shapes, images, tables) */
	noChangeAspect?: boolean
	/** Disallow rotating the object. (shapes, images) */
	noRot?: boolean
	/** Disallow editing the freeform/custom-geometry points. (shapes, images) */
	noEditPoints?: boolean
	/** Disallow moving the shape's adjustment handles. (shapes, images) */
	noAdjustHandles?: boolean
	/** Disallow changing arrowheads. (shapes, images) */
	noChangeArrowheads?: boolean
	/** Disallow changing the shape type (preset geometry). (shapes, images) */
	noChangeShapeType?: boolean
	/** Disallow editing the text body. (shapes / text boxes) */
	noTextEdit?: boolean
	/** Disallow cropping the picture. (images) */
	noCrop?: boolean
	/** Disallow drilling down into the graphical object (e.g. chart data). (tables) */
	noDrilldown?: boolean
}
