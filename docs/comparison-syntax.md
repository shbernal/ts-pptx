---
doc-schema-version: 1
title: "Side-By-Side Syntax"
summary: "Every intent in the comparison corpus as code: the calls ts-pptx 3.7.0 and pptxgenjs 4.0.1 were each given to produce the rows on the comparison page."
read_when:
  - Reading a comparison row and wanting the calls behind it
  - Porting a deck script from pptxgenjs to ts-pptx
  - Looking for the ts-pptx call that emits a particular construct
doc_type: "reference"
---

<!-- GENERATED FILE. Do not edit by hand.
     Regenerate with `pnpm run comparison:render`.
     Source: `scripts/comparison/snapshot.json`, written by `scripts/comparison/measure.mjs`. -->

# Side-By-Side Syntax

Every row of the [comparison](comparison.md) comes from running both libraries over a
corpus of deck intents. This page is that corpus as code: for each intent, the calls each
library was given. The harness lifts them out of the build functions as it measures and
records them in the snapshot beside the outcome they produced, so no snippet here can
illustrate a row that some earlier version of it produced.

Measured on 2026-09-05: ts-pptx 3.7.0 built from this repository, against pptxgenjs 4.0.1
installed from npm.

Each intent is written in the library's own idiom rather than transcribed from one into
the other. Transcribing is how a comparison of two APIs becomes a comparison of one API
and its translation. Where the two arms come out the same anyway the page prints one block
and says so, which is 8 of the 10 intents both libraries build. The rest are where a port
stops being a rename.

## How to read a snippet

Each block is the body of a build function. The harness puts the same frame around every
one of them, so this page states the frame once rather than repeating it on every intent:

```js
import TsPptx from 'pptx-ts/node'

const pres = new TsPptx()
// the snippet goes here
await pres.writeFile({ fileName: 'probe.pptx' })
```

and, for the other column:

```js
import PptxGenJS from 'pptxgenjs'

const pres = new PptxGenJS()
// the snippet goes here
await pres.writeFile({ fileName: 'probe.pptx' })
```

A snippet that needs a value the corpus declares once, such as a data URL or a chart
series, carries that declaration above it. Those lines come from the value the measurement
used, cut short with an ellipsis past 60 characters, and a path is written relative to the
repository root.

A library with no API for an intent has no block. The line under each heading says what
the harness read for both of them.

## Shared baseline

### Text run

`<a:t>` in `ppt/slides/slide1.xml`. ts-pptx: emitted. pptxgenjs: emitted.

Both libraries, called identically:

```js
pres.addSlide().addText('probe', { x: 1, y: 1, w: 4, h: 1 })
```

### Table

`<a:tbl>` in `ppt/slides/slide1.xml`. ts-pptx: emitted. pptxgenjs: emitted.

Both libraries, called identically:

```js
pres.addSlide().addTable(
	[
		[{ text: 'Region' }, { text: 'Units' }],
		[{ text: 'North' }, { text: '41' }],
	],
	{ x: 1, y: 1, w: 6 }
)
```

### Raster image

`<p:pic>` in `ppt/slides/slide1.xml`. ts-pptx: emitted. pptxgenjs: emitted.

**ts-pptx**

```js
const PNG_1PX_URL = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAA…'

pres.addSlide().addImage({ data: PNG_1PX_URL, x: 1, y: 1, w: 2, h: 2 })
```

**pptxgenjs**

```js
const PNG_1PX_BARE = 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcS…'

pres.addSlide().addImage({ data: PNG_1PX_BARE, x: 1, y: 1, w: 2, h: 2 })
```

### Bar chart

`<c:barChart>` in `ppt/charts/chart1.xml`. ts-pptx: emitted. pptxgenjs: emitted.

**ts-pptx**

```js
const BAR_DATA = [{ name: 'Revenue', labels: ['Q1', 'Q2', 'Q3'], values: [12, 19, 7] }]

pres.addSlide().addChart(BAR_DATA, { type: 'bar', x: 1, y: 1, w: 6, h: 4 })
```

**pptxgenjs**

```js
const BAR_DATA = [{ name: 'Revenue', labels: ['Q1', 'Q2', 'Q3'], values: [12, 19, 7] }]

pres.addSlide().addChart('bar', BAR_DATA, { x: 1, y: 1, w: 6, h: 4 })
```

### External hyperlink

`<a:hlinkClick` in `ppt/slides/slide1.xml`. ts-pptx: emitted. pptxgenjs: emitted.

Both libraries, called identically:

```js
pres.addSlide().addText('docs', { x: 1, y: 1, w: 4, h: 1, hyperlink: { url: 'https://example.com/' } })
```

### User-defined slide master

`<p:ph` in `ppt/slides/slide1.xml`. ts-pptx: emitted. pptxgenjs: emitted.

Both libraries, called identically:

```js
pres.defineSlideMaster({
	title: 'PROBE_MASTER',
	objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 1, y: 1, w: 8, h: 1 } } }],
})
pres.addSlide({ masterName: 'PROBE_MASTER' }).addText('probe', { placeholder: 'title' })
```

### Sections

`<p14:sectionLst` in `ppt/presentation.xml`. ts-pptx: emitted. pptxgenjs: emitted.

Both libraries, called identically:

```js
pres.addSection({ title: 'Findings' })
pres.addSlide({ sectionTitle: 'Findings' }).addText('probe', { x: 1, y: 1, w: 4, h: 1 })
```

### Speaker notes

`probe note` in `ppt/notesSlides/notesSlide1.xml`. ts-pptx: emitted. pptxgenjs: emitted.

Both libraries, called identically:

```js
pres.addSlide().addNotes('probe note')
```

### Preset-geometry shape

`<a:prstGeom` in `ppt/slides/slide1.xml`. ts-pptx: emitted. pptxgenjs: emitted.

Both libraries, called identically:

```js
pres.addSlide().addShape('roundRect', { x: 1, y: 1, w: 3, h: 2, fill: { color: '4472C4' } })
```

### Slide background colour

`<p:bg>` in `ppt/slides/slide1.xml`. ts-pptx: emitted. pptxgenjs: emitted.

Both libraries, called identically:

```js
const slide = pres.addSlide()
slide.background = { color: 'F2F2F2' }
slide.addText('probe', { x: 1, y: 1, w: 4, h: 1 })
```

## Motion

### Slide transition

`<p:transition` in `ppt/slides/slide1.xml`. ts-pptx: emitted. pptxgenjs: no API.

**ts-pptx**

```js
const slide = pres.addSlide()
slide.transition = { type: 'push', speed: 'slow', variant: { dir: 'd' } }
slide.addText('probe', { x: 1, y: 1, w: 4, h: 1 })
```

### Build animation on a shape

`<p:timing>` in `ppt/slides/slide1.xml`. ts-pptx: emitted. pptxgenjs: no API.

**ts-pptx**

```js
const slide = pres.addSlide()
slide.addText('probe', { x: 1, y: 1, w: 4, h: 1, objectName: 'Headline' })
slide.addAnimation({ preset: 'fadeIn', objectName: 'Headline' })
```

## Embedding

### Embedded OLE object

`<p:oleObj` in `ppt/slides/slide1.xml`. ts-pptx: emitted. pptxgenjs: no API.

**ts-pptx**

```js
const OLE_BLOB_B64 = 'dHMtcHB0eCBjb21wYXJpc29uIHByb2JlIHBheWxvYWQ='

pres.addSlide().addOleObject({ data: OLE_BLOB_B64, extn: 'bin', x: 1, y: 1, w: 4, h: 3 })
```

### 3D model

`am3d:model3d` in `ppt/slides/slide1.xml`. ts-pptx: emitted. pptxgenjs: no API.

**ts-pptx**

```js
const CUBE_GLB = 'demos/common/media/cube.glb'

pres.addSlide().addModel3d({ path: CUBE_GLB, meterPerModelUnit: 0.5, x: 1, y: 1, w: 4, h: 3 })
```

### Embedded font face

`<p:embeddedFontLst>` in `ppt/presentation.xml`. ts-pptx: emitted. pptxgenjs: no API.

**ts-pptx**

```js
const SILKSCREEN_TTF = 'test/read/fixtures/fonts/Silkscreen-Regular.ttf'

await pres.embedFont({ path: SILKSCREEN_TTF, typeface: 'Silkscreen' })
pres.addSlide().addText('probe', { x: 1, y: 1, w: 4, h: 1, fontFace: 'Silkscreen' })
```

## Shapes

### Connector between shapes

`<p:cxnSp>` in `ppt/slides/slide1.xml`. ts-pptx: emitted. pptxgenjs: no API.

**ts-pptx**

```js
const slide = pres.addSlide()
slide.addText('A', { objectName: 'BoxA', x: 1, y: 1, w: 1.5, h: 1 })
slide.addText('B', { objectName: 'BoxB', x: 5, y: 1, w: 1.5, h: 1 })
slide.addConnector({ type: 'elbow', x1: 2.5, y1: 1.5, x2: 5, y2: 1.5, startShape: 'BoxA', endShape: 'BoxB' })
```

## Text

### Inline equation

`<a14:m` in `ppt/slides/slide1.xml`. ts-pptx: emitted. pptxgenjs: no API.

**ts-pptx**

```js
const OMML_INLINE = '<m:oMath><m:r><m:t>n-1</m:t></m:r></m:oMath>'

pres.addSlide().addText([{ text: 'for all ' }, { math: OMML_INLINE, inline: true }, { text: ' terms' }], {
	x: 1,
	y: 1,
	w: 6,
	h: 1,
})
```

## Fills

### Gradient shape fill

`<a:gradFill` in `ppt/slides/slide1.xml`. ts-pptx: emitted. pptxgenjs: no API.

**ts-pptx**

```js
pres.addSlide().addShape('rect', {
	x: 1,
	y: 1,
	w: 4,
	h: 2,
	fill: {
		gradient: {
			kind: 'linear',
			angle: 45,
			stops: [
				{ position: 0, color: '4472C4' },
				{ position: 100, color: 'ED7D31' },
			],
		},
	},
})
```

The token is in the pptxgenjs bundle regardless. It appears only inside the bundled Office
theme XML, which no API parameterises.

## Tables

### 3D bevel on a table cell

`<a:cell3D` in `ppt/slides/slide1.xml`. ts-pptx: emitted. pptxgenjs: no API.

**ts-pptx**

```js
pres
	.addSlide()
	.addTable([[{ text: 'raised', options: { cell3D: { bevel: 'circle', width: 6, height: 6 } } }]], {
		x: 1,
		y: 1,
		w: 4,
	})
```

## Charts

### Funnel chart (chartEx)

`<cx:chart>` in `ppt/charts/chartEx1.xml`. ts-pptx: emitted. pptxgenjs: no API.

**ts-pptx**

```js
pres.addSlide().addChart([{ name: 'Stage', labels: ['Lead', 'Trial', 'Won'], values: [120, 48, 17] }], {
	type: 'funnel',
	x: 1,
	y: 1,
	w: 6,
	h: 4,
})
```

## Navigation

### Slide Zoom tile

`pslz:sldZm` in `ppt/slides/slide2.xml`. ts-pptx: emitted. pptxgenjs: no API.

**ts-pptx**

```js
pres.addSlide().addText('target', { x: 1, y: 1, w: 4, h: 1 })
pres.addSlide().addSlideZoom({ target: 1, x: 1, y: 1, w: 3, h: 1.7 })
```

## Diagrams

### SmartArt diagram (write side)

`<dgm:relIds` in `ppt/slides/slide1.xml`. ts-pptx: no API. pptxgenjs: no API.

Neither library has an API for this intent, so neither has a block.

## Adding one

The corpus is `scripts/comparison/probes.mjs`, one object per intent, and both arms of a
probe are ordinary code. A pull request that adds an intent is welcome, including one
ts-pptx fails. The harness reports the four outcomes it reads, and the comparison page
prints an intent upstream emits and we do not rather than dropping it.
