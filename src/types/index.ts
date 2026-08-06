/**
 * ts-pptx Interfaces — re-export barrel
 *
 * The public typed contract lives in `src/types/`, split by domain. This module re-exports
 * all of it so that `types/index.js` is the single import site for consumers and for the
 * rest of `src/`.
 *
 * The generator-internal `*Internal` wire shapes live alongside it in `types/internal.ts`
 * but are deliberately NOT re-exported here — they are not part of the published surface.
 * Internal code imports them straight from `./internal.js`.
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
 *   - `types/zoom.ts`      Slide/Section/Summary Zoom links
 *   - `types/model3d.ts`   embedded 3D models (`.glb`)
 *
 * Note: this barrel is not types-only — `types/text.ts` also exports the `textRun` /
 * `textRuns` run-array helpers.
 */

export * from './core.js'
export * from './style.js'
export * from './object.js'
export * from './theme.js'
export * from './text.js'
export * from './media.js'
export * from './shape.js'
export * from './table.js'
export * from './chart.js'
export * from './animation.js'
export * from './master.js'
export * from './slide.js'
export * from './pres.js'
export * from './zoom.js'
export * from './model3d.js'
