---
doc-schema-version: 1
title: "Side-By-Side Syntax"
summary: "Every intent in the comparison corpus as code: the calls ts-pptx 3.7.0 and pptxgenjs 4.0.1 were each given to produce the rows on the comparison page."
read_when:
  - Reading a comparison row and wanting the calls behind it
  - Porting a deck script from pptxgenjs to ts-pptx
  - Looking for the ts-pptx call that emits a particular construct
  - Reading a bundle size and wanting the program that was measured
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

Measured on 2026-09-06: ts-pptx 3.7.0 built from this repository, against pptxgenjs 4.0.1
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

## The bundle corpus

The [comparison](comparison.md) also reports what a bundler leaves after tree-shaking,
over whole decks rather than single constructs. These are those programs. Each one builds
the same deck with both libraries, using only constructs the shared baseline shows both of
them emitting, and each is run before it is bundled.

The frame is a consumer program rather than the probe harness, so it imports the package
root and keeps the exported bytes:

```js
import TsPptx from 'pptx-ts'

const pres = new TsPptx()
// the program goes here

console.log(await pres.write({ outputType: 'arraybuffer' }))
```

and, for the other column:

```js
import PptxGenJS from 'pptxgenjs'

const pres = new PptxGenJS()
// the program goes here

console.log(await pres.write({ outputType: 'arraybuffer' }))
```

### Hello world

One slide with one text box.

Both libraries, called identically:

```js
pres.addSlide().addText('hello', { x: 1, y: 1, w: 4, h: 1 })
```

### Text deck

A defined master, two sections, formatted and bulleted text, a hyperlink, a slide
background and speaker notes.

**ts-pptx**

```js
pres.defineSlideMaster({
	title: 'NARRATIVE',
	objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.5, w: 8.8, h: 1 } } }],
})
pres.addSection({ title: 'Findings' })
pres.addSection({ title: 'Next steps' })

const findings = pres.addSlide({ masterTitle: 'NARRATIVE', sectionTitle: 'Findings' })
findings.background = { color: 'F2F2F2' }
findings.addText('What we found', { placeholder: 'title' })
findings.addText(
	[
		{ text: 'Revenue is up', options: { bullet: true, bold: true } },
		{ text: 'Retention held', options: { bullet: true } },
		{ text: 'Payback lengthened', options: { bullet: true, color: 'C00000' } },
	],
	{ x: 0.6, y: 1.8, w: 8.8, h: 3, fontSize: 18 }
)
findings.addNotes('Twenty minutes. Hold questions until the last slide.')

const next = pres.addSlide({ masterTitle: 'NARRATIVE', sectionTitle: 'Next steps' })
next.addText('Where to read the rest', { placeholder: 'title' })
next.addText('The full report', {
	x: 0.6,
	y: 1.8,
	w: 8.8,
	h: 0.6,
	fontSize: 16,
	hyperlink: { url: 'https://example.com/report' },
})
next.addNotes('The link is the internal report, not the public summary.')
```

**pptxgenjs**

```js
pres.defineSlideMaster({
	title: 'NARRATIVE',
	objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.5, w: 8.8, h: 1 } } }],
})
pres.addSection({ title: 'Findings' })
pres.addSection({ title: 'Next steps' })

const findings = pres.addSlide({ masterName: 'NARRATIVE', sectionTitle: 'Findings' })
findings.background = { color: 'F2F2F2' }
findings.addText('What we found', { placeholder: 'title' })
findings.addText(
	[
		{ text: 'Revenue is up', options: { bullet: true, bold: true } },
		{ text: 'Retention held', options: { bullet: true } },
		{ text: 'Payback lengthened', options: { bullet: true, color: 'C00000' } },
	],
	{ x: 0.6, y: 1.8, w: 8.8, h: 3, fontSize: 18 }
)
findings.addNotes('Twenty minutes. Hold questions until the last slide.')

const next = pres.addSlide({ masterName: 'NARRATIVE', sectionTitle: 'Next steps' })
next.addText('Where to read the rest', { placeholder: 'title' })
next.addText('The full report', {
	x: 0.6,
	y: 1.8,
	w: 8.8,
	h: 0.6,
	fontSize: 16,
	hyperlink: { url: 'https://example.com/report' },
})
next.addNotes('The link is the internal report, not the public summary.')
```

### Table deck

A titled slide and a bordered table with a header row, fixed column widths and per-cell
options.

**ts-pptx**

```js
const REGION_ROWS = [[{ text: 'Region' }, { text: 'Revenue' }, { text: 'Growth' }], [{ text: 'North America' }, { text: '24.9' }, { text: '14.2%' }], [{ text: 'EMEA' }, { text: '12.6' }, { text: '16.8%' }], [{ text: 'APAC' }, { text: '6.8' }, { text: '9.4%' }]]

const slide = pres.addSlide()
slide.addText('Revenue by region', { x: 0.6, y: 0.5, w: 8.8, h: 0.8, fontSize: 24, bold: true })
slide.addTable(REGION_ROWS, {
	x: 0.6,
	y: 1.5,
	w: 8.8,
	colW: [4, 2.4, 2.4],
	rowH: 0.4,
	border: { type: 'solid', width: 1, color: 'D9D9D9' },
	fill: { color: 'FFFFFF' },
	fontSize: 12,
	valign: 'middle',
})
```

**pptxgenjs**

```js
const REGION_ROWS = [[{ text: 'Region' }, { text: 'Revenue' }, { text: 'Growth' }], [{ text: 'North America' }, { text: '24.9' }, { text: '14.2%' }], [{ text: 'EMEA' }, { text: '12.6' }, { text: '16.8%' }], [{ text: 'APAC' }, { text: '6.8' }, { text: '9.4%' }]]

const slide = pres.addSlide()
slide.addText('Revenue by region', { x: 0.6, y: 0.5, w: 8.8, h: 0.8, fontSize: 24, bold: true })
slide.addTable(REGION_ROWS, {
	x: 0.6,
	y: 1.5,
	w: 8.8,
	colW: [4, 2.4, 2.4],
	rowH: 0.4,
	border: { type: 'solid', pt: 1, color: 'D9D9D9' },
	fill: { color: 'FFFFFF' },
	fontSize: 12,
	valign: 'middle',
})
```

### Chart deck

Three charts on three slides: a column chart with value labels, a two-series line chart,
and a pie chart with percentages.

**ts-pptx**

```js
const BAR_DATA = [{ name: 'Revenue', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [12, 19, 7, 24] }]
const LINE_DATA = [{ name: 'Net retention', labels: ['Jul', 'Aug', 'Sep'], values: [108, 111, 114] }, { name: 'Gross retention', labels: ['Jul', 'Aug', 'Sep'], values: [94, 94, 96] }]
const PIE_DATA = [{ name: 'Revenue mix', labels: ['Platform', 'Services', 'Licensing'], values: [25.8, 13.4, 7.2] }]

pres.addSlide().addChart(BAR_DATA, {
	type: 'bar',
	barDir: 'col',
	x: 0.6,
	y: 0.6,
	w: 8.8,
	h: 5,
	showValue: true,
	showLegend: true,
	legendPos: 'b',
})
pres.addSlide().addChart(LINE_DATA, {
	type: 'line',
	x: 0.6,
	y: 0.6,
	w: 8.8,
	h: 5,
	lineSmooth: true,
	showLegend: true,
	legendPos: 'b',
})
pres.addSlide().addChart(PIE_DATA, {
	type: 'pie',
	x: 0.6,
	y: 0.6,
	w: 8.8,
	h: 5,
	showPercent: true,
	dataLabelColor: 'FFFFFF',
})
```

**pptxgenjs**

```js
const BAR_DATA = [{ name: 'Revenue', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [12, 19, 7, 24] }]
const LINE_DATA = [{ name: 'Net retention', labels: ['Jul', 'Aug', 'Sep'], values: [108, 111, 114] }, { name: 'Gross retention', labels: ['Jul', 'Aug', 'Sep'], values: [94, 94, 96] }]
const PIE_DATA = [{ name: 'Revenue mix', labels: ['Platform', 'Services', 'Licensing'], values: [25.8, 13.4, 7.2] }]

pres.addSlide().addChart('bar', BAR_DATA, {
	barDir: 'col',
	x: 0.6,
	y: 0.6,
	w: 8.8,
	h: 5,
	showValue: true,
	showLegend: true,
	legendPos: 'b',
})
pres.addSlide().addChart('line', LINE_DATA, {
	x: 0.6,
	y: 0.6,
	w: 8.8,
	h: 5,
	lineSmooth: true,
	showLegend: true,
	legendPos: 'b',
})
pres.addSlide().addChart('pie', PIE_DATA, {
	x: 0.6,
	y: 0.6,
	w: 8.8,
	h: 5,
	showPercent: true,
	dataLabelColor: 'FFFFFF',
})
```

### Full deck

Every construct the shared baseline covers, in one deck: master, sections, background,
text, hyperlink, notes, a preset shape, an image, a table and a chart.

**ts-pptx**

```js
const PNG_1PX_URL = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAA…'
const REGION_ROWS = [[{ text: 'Region' }, { text: 'Revenue' }, { text: 'Growth' }], [{ text: 'North America' }, { text: '24.9' }, { text: '14.2%' }], [{ text: 'EMEA' }, { text: '12.6' }, { text: '16.8%' }], [{ text: 'APAC' }, { text: '6.8' }, { text: '9.4%' }]]
const BAR_DATA = [{ name: 'Revenue', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [12, 19, 7, 24] }]

pres.defineSlideMaster({
	title: 'REVIEW',
	objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.5, w: 8.8, h: 1 } } }],
})
pres.addSection({ title: 'Quarter' })

const cover = pres.addSlide({ masterTitle: 'REVIEW', sectionTitle: 'Quarter' })
cover.background = { color: 'F2F2F2' }
cover.addText('Q3 review', { placeholder: 'title' })
cover.addShape('roundRect', { x: 0.6, y: 1.8, w: 3, h: 0.6, fill: { color: '4472C4' } })
cover.addText('Read the report', {
	x: 0.6,
	y: 2.6,
	w: 4,
	h: 0.5,
	hyperlink: { url: 'https://example.com/report' },
})
cover.addImage({ data: PNG_1PX_URL, x: 6.4, y: 1.8, w: 2, h: 2 })
cover.addNotes('Open on the number, not the agenda.')

const numbers = pres.addSlide({ masterTitle: 'REVIEW', sectionTitle: 'Quarter' })
numbers.addText('Where it came from', { placeholder: 'title' })
numbers.addTable(REGION_ROWS, { x: 0.6, y: 1.6, w: 8.8, colW: [4, 2.4, 2.4], fontSize: 12 })

pres.addSlide({ masterTitle: 'REVIEW', sectionTitle: 'Quarter' }).addChart(BAR_DATA, {
	type: 'bar',
	barDir: 'col',
	x: 0.6,
	y: 1.6,
	w: 8.8,
	h: 4.6,
	showValue: true,
})
```

**pptxgenjs**

```js
const PNG_1PX_BARE = 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcS…'
const REGION_ROWS = [[{ text: 'Region' }, { text: 'Revenue' }, { text: 'Growth' }], [{ text: 'North America' }, { text: '24.9' }, { text: '14.2%' }], [{ text: 'EMEA' }, { text: '12.6' }, { text: '16.8%' }], [{ text: 'APAC' }, { text: '6.8' }, { text: '9.4%' }]]
const BAR_DATA = [{ name: 'Revenue', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [12, 19, 7, 24] }]

pres.defineSlideMaster({
	title: 'REVIEW',
	objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.5, w: 8.8, h: 1 } } }],
})
pres.addSection({ title: 'Quarter' })

const cover = pres.addSlide({ masterName: 'REVIEW', sectionTitle: 'Quarter' })
cover.background = { color: 'F2F2F2' }
cover.addText('Q3 review', { placeholder: 'title' })
cover.addShape('roundRect', { x: 0.6, y: 1.8, w: 3, h: 0.6, fill: { color: '4472C4' } })
cover.addText('Read the report', {
	x: 0.6,
	y: 2.6,
	w: 4,
	h: 0.5,
	hyperlink: { url: 'https://example.com/report' },
})
cover.addImage({ data: PNG_1PX_BARE, x: 6.4, y: 1.8, w: 2, h: 2 })
cover.addNotes('Open on the number, not the agenda.')

const numbers = pres.addSlide({ masterName: 'REVIEW', sectionTitle: 'Quarter' })
numbers.addText('Where it came from', { placeholder: 'title' })
numbers.addTable(REGION_ROWS, { x: 0.6, y: 1.6, w: 8.8, colW: [4, 2.4, 2.4], fontSize: 12 })

pres.addSlide({ masterName: 'REVIEW', sectionTitle: 'Quarter' }).addChart('bar', BAR_DATA, {
	barDir: 'col',
	x: 0.6,
	y: 1.6,
	w: 8.8,
	h: 4.6,
	showValue: true,
})
```

## The timing corpus

The [comparison](comparison.md) also reports how long each library takes to turn a deck
into bytes. The small end of that measurement is the bundle corpus above; the large end is
this, one deck shape built at 50, 200 and 500 slides. It is printed once because the slide
count is the only thing that changes between them, and it is deliberately dull: a timing
corpus is not hunting for the slowest construct, it is making the per-slide cost visible.

`slides` below is that count. Everything else is the same code at every size.

**ts-pptx**

```js
const rows = [
	[{ text: 'Region' }, { text: 'Revenue' }, { text: 'Growth' }],
	[{ text: 'North America' }, { text: '24.9' }, { text: '14.2%' }],
	[{ text: 'EMEA' }, { text: '12.6' }, { text: '16.8%' }],
	[{ text: 'APAC' }, { text: '6.8' }, { text: '9.4%' }],
]
const bars = [{ name: 'Revenue', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [12, 19, 7, 24] }]
for (let index = 0; index < slides; index++) {
	const slide = pres.addSlide()
	slide.addText('Slide ' + (index + 1), { x: 0.6, y: 0.4, w: 8.8, h: 0.8, fontSize: 24, bold: true })
	slide.addText(
		[
			{ text: 'What we saw', options: { bullet: true, bold: true } },
			{ text: 'What we changed', options: { bullet: true } },
			{ text: 'What it cost', options: { bullet: true } },
		],
		{ x: 0.6, y: 1.4, w: 4, h: 2.4, fontSize: 14 }
	)
	if (index % 2 === 0) slide.addTable(rows, { x: 5, y: 1.4, w: 4.4, colW: [2, 1.2, 1.2], fontSize: 10 })
	else slide.addChart(bars, { type: 'bar', barDir: 'col', x: 5, y: 1.4, w: 4.4, h: 3, showValue: true })
	slide.addNotes('Speaker notes for slide ' + (index + 1) + '.')
}
```

**pptxgenjs**

```js
const rows = [
	[{ text: 'Region' }, { text: 'Revenue' }, { text: 'Growth' }],
	[{ text: 'North America' }, { text: '24.9' }, { text: '14.2%' }],
	[{ text: 'EMEA' }, { text: '12.6' }, { text: '16.8%' }],
	[{ text: 'APAC' }, { text: '6.8' }, { text: '9.4%' }],
]
const bars = [{ name: 'Revenue', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [12, 19, 7, 24] }]
for (let index = 0; index < slides; index++) {
	const slide = pres.addSlide()
	slide.addText('Slide ' + (index + 1), { x: 0.6, y: 0.4, w: 8.8, h: 0.8, fontSize: 24, bold: true })
	slide.addText(
		[
			{ text: 'What we saw', options: { bullet: true, bold: true } },
			{ text: 'What we changed', options: { bullet: true } },
			{ text: 'What it cost', options: { bullet: true } },
		],
		{ x: 0.6, y: 1.4, w: 4, h: 2.4, fontSize: 14 }
	)
	if (index % 2 === 0) slide.addTable(rows, { x: 5, y: 1.4, w: 4.4, colW: [2, 1.2, 1.2], fontSize: 10 })
	else slide.addChart('bar', bars, { barDir: 'col', x: 5, y: 1.4, w: 4.4, h: 3, showValue: true })
	slide.addNotes('Speaker notes for slide ' + (index + 1) + '.')
}
```

## Adding one

The corpus is `scripts/comparison/probes.mjs`, one object per intent, and both arms of a
probe are ordinary code. A pull request that adds an intent is welcome, including one
ts-pptx fails. The harness reports the four outcomes it reads, and the comparison page
prints an intent upstream emits and we do not rather than dropping it.

The bundle corpus is `scripts/comparison/programs.mjs`, one object per program, on one
extra condition: both arms have to build the same deck. A program only one library can
build measures two different pieces of work and reports the difference as a size.
