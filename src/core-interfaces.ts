/**
 * PptxGenJS Interfaces — re-export barrel
 *
 * The public typed contract (plus the internal `*Internal` shapes the generators pass
 * around) lives in `src/types/`, split by domain. This module re-exports all of it so
 * that `./core-interfaces.js` remains the single import site for consumers and for the
 * rest of `src/`.
 *
 * Where things live:
 *   - `types/core.ts`      Coord/PositionProps, colors, gradient/pattern/image fills, geometry points
 *   - `types/style.ts`     borders, shadows, shape fill/line, connectors, hyperlinks
 *   - `types/object.ts`    PlaceholderProps, object name/lock
 *   - `types/theme.ts`     ThemeColorScheme, ThemeProps
 *   - `types/text.ts`      TextBaseProps, TextPropsOptions/TextProps, measurement, notes, comments
 *   - `types/media.ts`     MediaType, ImageProps, MediaProps
 *   - `types/shape.ts`     ShapeProps and adjust values
 *   - `types/table.ts`     TableProps/TableCell(+Props), table styles, tableToSlides, layout results
 *   - `types/chart.ts`     OptsChartData(+Internal), per-axis and per-type chart props, ChartOpts
 *   - `types/animation.ts` transitions and slide animations
 *   - `types/master.ts`    slide-master objects, bullets and per-level text styles
 *   - `types/slide.ts`     groups, ObjectOptions, the SlideLayout/PresSlide authoring surfaces
 *   - `types/pres.ts`      WriteProps, sections, PresLayout, presentation props
 *   - `types/internal.ts`  generator-internal wire shapes — NOT public contract
 *
 * Note: this barrel is not types-only — `types/text.ts` also exports the `textRun` /
 * `textRuns` run-array helpers.
 */

export * from './types/core.js'
export * from './types/style.js'
export * from './types/object.js'
export * from './types/theme.js'
export * from './types/text.js'
export * from './types/media.js'
export * from './types/shape.js'
export * from './types/table.js'
export * from './types/chart.js'
export * from './types/animation.js'
export * from './types/master.js'
export * from './types/slide.js'
export * from './types/pres.js'
export * from './types/internal.js'
