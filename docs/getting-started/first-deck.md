---
doc-schema-version: 1
title: "Your first deck"
summary: "A tutorial that builds a four-slide quarterly summary from an array: a title slide, bullet points with speaker notes, a table and a chart, then saves the file."
read_when:
  - Building a first deck with ts-pptx
  - Looking for a complete program to copy and run
  - Getting a deck out as a file or as bytes
doc_type: "guide"
---

# Your first deck

This page builds a four-slide quarterly summary from an array of numbers: a title slide, a slide of
bullet points with speaker notes, a table and a bar chart. Each section adds a few lines, and the
[complete program](#the-complete-program) is at the end.

You need Node.js 24 or later and the package installed ([Installation](installation.md)). Save the
program as `deck.mts` and run it with `node deck.mts`. Node 24 runs TypeScript directly by stripping
the types, so there is no build step.

## The data

In a real report this array comes from a database or an API:

```ts
const quarters = [
  { quarter: "Q1", revenue: 1.2, customers: 310 },
  { quarter: "Q2", revenue: 1.5, customers: 355 },
  { quarter: "Q3", revenue: 1.9, customers: 420 },
  { quarter: "Q4", revenue: 2.3, customers: 480 },
]
```

## Create the presentation

```ts
import TsPptx, { type TableRow } from "pptx-ts"

const pptx = new TsPptx()
pptx.layout = "LAYOUT_WIDE"
```

`TsPptx` is the presentation. `layout` sets the slide size for the whole deck: `LAYOUT_WIDE` is
13.333 by 7.5 inches, the widescreen size PowerPoint uses for a new deck. Without it you get
`LAYOUT_16x9`, 10 by 5.625 inches. `TableRow` is a type the table further down uses.

## A title slide

```ts
const cover = pptx.addSlide()
cover.background = { color: "1F3A5F" }
cover.addText("Quarterly summary", {
  x: 0.8, y: 2.6, w: 11.7, h: 1.2,
  fontSize: 44, bold: true, color: "FFFFFF",
})
cover.addText("Revenue and customers, Q1 to Q4", {
  x: 0.8, y: 3.8, w: 11.7, h: 0.8,
  fontSize: 20, color: "C9D6E8",
})
```

`addSlide()` appends a slide and returns it. Every `add*` call on a slide takes the content first,
then one options object holding the position, the size and the formatting together.

- `x` and `y` place the top-left corner of the text box, and `w` and `h` size it. All four are in
  inches from the top-left corner of the slide.
- Colours are six-digit hex values. A leading `#` is accepted too.

In PowerPoint this is a dark blue slide with a large white title and a paler subtitle under it.

## Bullet points and speaker notes

```ts
const total = quarters.reduce((sum, q) => sum + q.revenue, 0)
const best = quarters.reduce((a, b) => (b.revenue > a.revenue ? b : a))
const peakCustomers = Math.max(...quarters.map((q) => q.customers))

const highlights = pptx.addSlide()
highlights.addText("Highlights", { x: 0.8, y: 0.5, w: 11.7, h: 0.9, fontSize: 32, bold: true })
highlights.addText(
  [
    { text: `Revenue for the year: $${total.toFixed(1)}M`, options: { bullet: true, breakLine: true } },
    { text: `Best quarter: ${best.quarter}, at $${best.revenue.toFixed(1)}M`, options: { bullet: true, breakLine: true } },
    { text: `Customers at the peak: ${peakCustomers}`, options: { bullet: true } },
  ],
  { x: 0.8, y: 1.6, w: 11.7, h: 3, fontSize: 24 },
)
highlights.addNotes("Q4 includes the holiday promotion, so expect Q1 to come in lower.")
```

Instead of a string, `addText` takes an array of runs here. Each run is a piece of text with its own
options: `bullet: true` makes it a bullet point, and `breakLine: true` ends the paragraph after it.
The options object after the array applies to the whole box.

`addNotes` writes the speaker notes, which PowerPoint shows under the slide and in Presenter View.

## A table from the data

```ts
const tableSlide = pptx.addSlide()
tableSlide.addText("By quarter", { x: 0.8, y: 0.5, w: 11.7, h: 0.9, fontSize: 32, bold: true })

const header = { bold: true, color: "FFFFFF", fill: { color: "1F3A5F" } }
const rows: TableRow[] = [
  [
    { text: "Quarter", options: header },
    { text: "Revenue ($M)", options: header },
    { text: "Customers", options: header },
  ],
  ...quarters.map((q): TableRow => [
    { text: q.quarter },
    { text: q.revenue.toFixed(1), options: { align: "right" } },
    { text: String(q.customers), options: { align: "right" } },
  ]),
]
tableSlide.addTable(rows, { x: 0.8, y: 1.6, w: 8, colW: [2, 3, 3], fontSize: 18 })
```

A table is a list of rows, and a row is a list of cells. Each cell has its `text` and, optionally,
its own `options`. `colW` sets the column widths in inches. A table with more rows than fit on one
slide can continue onto new slides by itself; [Tables](../tables.md) covers that and the rest of the
styling.

## A chart from the same data

```ts
const chartSlide = pptx.addSlide()
chartSlide.addText("Revenue by quarter", { x: 0.8, y: 0.5, w: 11.7, h: 0.9, fontSize: 32, bold: true })
chartSlide.addChart(
  [
    {
      name: "Revenue ($M)",
      labels: quarters.map((q) => q.quarter),
      values: quarters.map((q) => q.revenue),
    },
  ],
  {
    type: "bar",
    x: 0.8, y: 1.6, w: 11.7, h: 5.4,
    chartColors: ["1F3A5F"],
    showValue: true,
    dataLabelFormatCode: "0.0",
  },
)
```

The first argument is a list of series, each with a `name`, the category `labels` and one number per
label in `values`. `type` picks the chart. The chart carries a workbook with those numbers inside the
file, so right-clicking it in PowerPoint and choosing **Edit Data** opens them.

## Save the file

```ts
await pptx.writeFile({ fileName: "quarterly-summary.pptx" })
```

Under Node this writes the file into the current directory. In a browser, the same call downloads it.
To send the deck somewhere instead, as a server response or an upload, ask for the bytes:

```ts
const bytes = await pptx.toBytes()
```

[Where it runs](runtime.md#which-build-the-bare-import-gives-you) lists what
`writeFile` does in each runtime.

## The complete program

<!-- first-deck:program -->
```ts
import TsPptx, { type TableRow } from "pptx-ts"

const quarters = [
  { quarter: "Q1", revenue: 1.2, customers: 310 },
  { quarter: "Q2", revenue: 1.5, customers: 355 },
  { quarter: "Q3", revenue: 1.9, customers: 420 },
  { quarter: "Q4", revenue: 2.3, customers: 480 },
]

const pptx = new TsPptx()
pptx.layout = "LAYOUT_WIDE"

// Title slide
const cover = pptx.addSlide()
cover.background = { color: "1F3A5F" }
cover.addText("Quarterly summary", {
  x: 0.8, y: 2.6, w: 11.7, h: 1.2,
  fontSize: 44, bold: true, color: "FFFFFF",
})
cover.addText("Revenue and customers, Q1 to Q4", {
  x: 0.8, y: 3.8, w: 11.7, h: 0.8,
  fontSize: 20, color: "C9D6E8",
})

// Bullet points and speaker notes
const total = quarters.reduce((sum, q) => sum + q.revenue, 0)
const best = quarters.reduce((a, b) => (b.revenue > a.revenue ? b : a))
const peakCustomers = Math.max(...quarters.map((q) => q.customers))

const highlights = pptx.addSlide()
highlights.addText("Highlights", { x: 0.8, y: 0.5, w: 11.7, h: 0.9, fontSize: 32, bold: true })
highlights.addText(
  [
    { text: `Revenue for the year: $${total.toFixed(1)}M`, options: { bullet: true, breakLine: true } },
    { text: `Best quarter: ${best.quarter}, at $${best.revenue.toFixed(1)}M`, options: { bullet: true, breakLine: true } },
    { text: `Customers at the peak: ${peakCustomers}`, options: { bullet: true } },
  ],
  { x: 0.8, y: 1.6, w: 11.7, h: 3, fontSize: 24 },
)
highlights.addNotes("Q4 includes the holiday promotion, so expect Q1 to come in lower.")

// A table from the data
const tableSlide = pptx.addSlide()
tableSlide.addText("By quarter", { x: 0.8, y: 0.5, w: 11.7, h: 0.9, fontSize: 32, bold: true })

const header = { bold: true, color: "FFFFFF", fill: { color: "1F3A5F" } }
const rows: TableRow[] = [
  [
    { text: "Quarter", options: header },
    { text: "Revenue ($M)", options: header },
    { text: "Customers", options: header },
  ],
  ...quarters.map((q): TableRow => [
    { text: q.quarter },
    { text: q.revenue.toFixed(1), options: { align: "right" } },
    { text: String(q.customers), options: { align: "right" } },
  ]),
]
tableSlide.addTable(rows, { x: 0.8, y: 1.6, w: 8, colW: [2, 3, 3], fontSize: 18 })

// A chart from the same data
const chartSlide = pptx.addSlide()
chartSlide.addText("Revenue by quarter", { x: 0.8, y: 0.5, w: 11.7, h: 0.9, fontSize: 32, bold: true })
chartSlide.addChart(
  [
    {
      name: "Revenue ($M)",
      labels: quarters.map((q) => q.quarter),
      values: quarters.map((q) => q.revenue),
    },
  ],
  {
    type: "bar",
    x: 0.8, y: 1.6, w: 11.7, h: 5.4,
    chartColors: ["1F3A5F"],
    showValue: true,
    dataLabelFormatCode: "0.0",
  },
)

await pptx.writeFile({ fileName: "quarterly-summary.pptx" })
```

Run it with `node deck.mts` and open `quarterly-summary.pptx`.

## Where to go next

| To | Read |
| --- | --- |
| Understand the model behind these calls | [Core concepts](concepts.md) |
| Style a table, merge cells, or page a long one across slides | [Tables](../tables.md) |
| Clip a picture to a shape | [Image embedded in a shape](../image-in-shape.md) |
| Make text fit its box | [Measured text fit](../measured-text-fit.md) |
| Draw lines that stay attached to shapes | [Connectors](../connectors.md) |
| Open an existing deck and edit it | [PPTX read and round-trip](../reference/pptx-read.md) |
| Turn a deck into a script | [PPTX to script](../reference/pptx-to-script.md) |
| Ship a smaller browser bundle | [Smaller bundles](../bundle-size.md) |
| Look up every option | [API reference](../reference/api/index.md) |
