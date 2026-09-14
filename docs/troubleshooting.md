---
doc-schema-version: 1
title: "Troubleshooting"
summary: "Symptoms seen in a generated deck or at runtime, grouped by where they show up, each with its cause, its fix and the guide section that documents it, and where to report a problem the fixes do not cover."
read_when:
  - PowerPoint offers to repair a generated deck
  - A table border, table style, picture, 3D model, OLE object or embedded font does not look the way the options suggest
  - Text overflows its box although fit is set
  - Importing the package, writeFile or Presentation.load fails in one runtime
  - Deciding whether a problem is worth reporting
doc_type: "troubleshooting"
---

# Troubleshooting

Each table lists what you see, what causes it, and what to change. The last column links the guide section that documents the behavior. [Errors and warnings](errors-and-warnings.md) lists every error class and code.

## Opening the file

| Symptom | Cause | Fix | Guide |
| --- | --- | --- | --- |
| PowerPoint offers to repair the deck, or opening it fails with `0x80070570` | The package holds markup or a part PowerPoint rejects. A deck the library wrote should open with no repair prompt, so this is a library bug. | Reduce it to the shortest script that builds a failing deck, and [report it](#report-a-problem). | [Introduction](getting-started/introduction.md#what-works-means) |

`npx ooxml-validate deck.pptx` runs Microsoft's Open XML SDK validator over a file and prints what it finds as JSON. It downloads the validator the first time it runs. Attach the output to the report. A clean result does not prove PowerPoint opens the file, because PowerPoint rejects some markup the schema allows.

## Layout and text

| Symptom | Cause | Fix | Guide |
| --- | --- | --- | --- |
| A table border draws on every cell, not only around the table | `border` styles each cell's own four edges. | Set the outline with `outerBorder`, which draws the perimeter only. | [Tables](tables.md#draw-borders) |
| A table shows a black hairline grid instead of its table style | PowerPoint paints only its built-in styles and never reads a style definition from the file. Any other style id, from a cast or a template's custom style, falls back to no style, and the library does not warn. | Use a `TableStyle` member, or style the table with `headerRow`, `columns`, `fill` and `border`. | [Tables](tables.md#apply-a-built-in-table-style) |
| Text with `fit: 'shrink'` or `fit: 'resize'` overflows its box | No font metrics are registered, so the library writes a bare autofit flag. PowerPoint computes the fit only after the text or the box is edited. | Await `registerFontMetrics()` for the box's `fontFace` before writing the deck. | [Text that fits](text-fit.md#register-font-metrics) |
| Fitted text overflows, and `measure/shrink-unmeasured` or `measure/resize-unmeasured` is reported | The box has no `fontFace`, and the library does not resolve the theme font. | Set `fontFace` on fitted text. | [Text that fits](text-fit.md#handle-a-face-with-no-metrics) |
| Fitted text shrinks more than needed, and `measure/heuristic-metrics` is reported | The box's `fontFace` has no registered metrics, so it is measured with average character widths, which err wide. | Register that face, and each bold or italic variant the text uses. | [Text that fits](text-fit.md#handle-a-face-with-no-metrics) |
| Fitted text overflows by a line, and `measure/uncovered-codepoints` is reported | The registered face has no glyph for some characters. PowerPoint draws them from a substitute font, which can be wider. | Register a face that covers the text. | [Text that fits](text-fit.md#handle-a-face-with-no-metrics) |
| Text in a table cell does not shrink | `fit: 'shrink'` acts only on fixed-height rows. A row with no `rowH` entry and no table `h` grows with its text. | Fix the row height with `rowH` or `h`. | [Tables](tables.md#set-row-heights) |
| A photo looks squashed or stretched | With no `sizing`, a raster image stretches to its box. | Set `sizing: { type: 'cover' }` to crop, or `'contain'` to fit inside the box. | [Images in shapes](image-in-shape.md#fill-the-shape-without-distortion) |
| An SVG sits in its box with empty bands on two sides | With no `sizing`, an SVG that states its size or `viewBox` is letterboxed at its own aspect ratio. | Give the box the SVG's ratio, or set `sizing` to `cover` or `stretch`. | [Images in shapes](image-in-shape.md#place-an-svg-at-its-own-aspect-ratio) |
| An SVG comes out 1 inch square | Given neither `w` nor `h`, an SVG is placed at 1 in by 1 in. Its user units are not read as pixels. | Set `w` or `h`. The other side follows the SVG's ratio. | [Images in shapes](image-in-shape.md#place-an-svg-at-its-own-aspect-ratio) |

## Embedded content

| Symptom | Cause | Fix | Guide |
| --- | --- | --- | --- |
| A 3D model draws far too large, or the slide shows a wall of shading | The library does not measure the `.glb` and writes `meterPerModelUnit: 0.5`, which suits a model 2 units across. A larger model surrounds the default camera, 2.26 m from the centre. | Set `meterPerModelUnit` to 1 divided by the model's largest bounding-box dimension. | [3D models](3d-models.md#set-the-scale) |
| A 3D model draws from the wrong angle | The default camera sits on the z axis and looks at the origin head-on. | Set `camera.pos`, and `lookAt` if needed. | [3D models](3d-models.md#set-the-camera) |
| A gray square shows where a 3D model should be, outside PowerPoint 2019 and later | Those readers draw the preview picture, and with no `preview` that is a gray placeholder. `model3d/preview-missing` is reported. | Pass `preview` with a picture of the model at the frame's aspect ratio. | [3D models](3d-models.md#supply-a-preview-picture) |
| An OLE object shows as a gray square in a viewer that draws its picture | The library never opens the payload, so with no `cover` it embeds a gray placeholder and reports nothing. | Pass `cover` with a picture of the file, or an icon picture with `showAsIcon: true`. | [OLE embedded objects](ole-objects.md#supply-a-cover-picture) |
| Text does not use the embedded font | PowerPoint matches an embedded font by name, and the text's `fontFace` differs from the `typeface` given to `embedFont`. The library does not read the font file, so it does not warn. | Use the family name the font file declares, in both `typeface` and `fontFace`. | [Embedded fonts](embedded-fonts.md#name-the-typeface-your-text-uses) |
| A slide imported from another deck renders in a substitute font | Fonts belong to the presentation, and an import copies them only when asked. | Pass `embedFonts: true` to `importSlide`, `importSlides` or `importSlideMasters`. | [Embedded fonts](embedded-fonts.md#carry-fonts-when-importing-slides) |

## Installing and running

| Symptom | Cause | Fix | Guide |
| --- | --- | --- | --- |
| `require("pptx-ts")` returns an object, and `new` on it throws | `require()` returns the ES module's namespace, and the class is its default export. | Write `const { default: TsPptx } = require("pptx-ts")`. | [Installation](getting-started/installation.md#commonjs) |
| npm or pnpm reports an unsupported engine | The package declares `"engines": { "node": ">=24" }`. | Run Node.js 24 or later. | [Installation](getting-started/installation.md#requirements) |
| Importing a file under `pptx-ts/dist/` fails with `ERR_PACKAGE_PATH_NOT_EXPORTED` | The package's `exports` map publishes only its entry points, and Node refuses every other path. | Import an entry point, such as `pptx-ts/node` or `pptx-ts/read`. | [Where it runs](getting-started/runtime.md#entry-points) |
| `writeFile()` throws `runtime/file-output-unavailable` on Deno, Bun or an edge worker | These runtimes load the runtime-agnostic build, which has no filesystem and no page to download onto. | Call `write()` and store the bytes yourself. | [Where it runs](getting-started/runtime.md#deno-bun-and-edge-workers) |
| `Presentation.load("deck.pptx")` throws `zip/filesystem-unavailable` | A string is read as a file path, and only Node has a filesystem to read it from. | Pass the bytes as a `Uint8Array`, `ArrayBuffer` or `Blob`. | [Where it runs](getting-started/runtime.md#entry-points) |
| A method throws `family/not-composed` | The deck came from `createPresentation` without the construct family that method belongs to. | Add the family to `use`, as in `createPresentation({ use: [charts] })`. | [Smaller bundles](bundle-size.md#when-a-family-is-missing) |
| `latexToOmml()` or `mathmlToOmml()` throws `math/missing-optional-peer` | `temml` and `mathml2omml` are optional peer dependencies, and one is not installed. | Run `npm install temml mathml2omml`. | [Math equations](math-latex.md#install-the-converters) |

## Report a problem

Open an issue at <https://github.com/shbernal/ts-pptx/issues> when no fix above applies, or when [Errors and warnings](errors-and-warnings.md) says the failure is worth reporting. Include the shortest script that reproduces it. Never attach a deck from a real project.

An `InternalError` means an invariant of the library did not hold. Its message ends with the link to file it:

```text
This is a bug in ts-pptx, not in your deck or your code. Please report it:
https://github.com/shbernal/ts-pptx/issues/new?template=agent-report.yml
```

A check that fails while you work on the repository is covered in [When a check fails](https://github.com/shbernal/ts-pptx/blob/master/docs/contributing/development.md#when-a-check-fails).
