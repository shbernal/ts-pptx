/**
 * ts-pptx Interfaces — re-export barrel
 *
 * The public typed contract lives in `src/types/`, split by domain. This module re-exports
 * all of it so that `./core-interfaces.js` remains the single import site for consumers and
 * for the rest of `src/`.
 *
 * The generator-internal `*Internal` wire shapes live alongside it in `types/internal.ts`
 * but are deliberately NOT re-exported here — they are not part of the published surface.
 * Internal code imports them straight from `./types/internal.js`.
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
 *   - `types/chart.ts`     OptsChartData, per-axis and per-type chart props, ChartOpts
 *   - `types/animation.ts` transitions and slide animations
 *   - `types/master.ts`    slide-master objects, bullets and per-level text styles
 *   - `types/slide.ts`     groups, ObjectOptions, the SlideLayout/Slide authoring surfaces
 *   - `types/pres.ts`      WriteProps, sections, PresLayout, presentation props
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
export * from './types/zoom.js'
