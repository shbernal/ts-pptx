---
doc-schema-version: 1
title: "HTML tables to slides"
summary: "Convert an existing HTML table into PowerPoint tables with tableToSlides, including how column widths are decided in and out of a browser."
read_when:
  - Converting an HTML table into slides
  - Deciding why column widths differ between a browser and Node
  - Changing the tableToSlides column-width basis
doc_type: "reference"
---

# HTML tables to slides

`pptx-ts/html` reproduces an existing HTML `<table>` as a PowerPoint table, paging
across as many slides as its rows need. One artifact serves both runtimes: it runs in
a browser, and under Node with any DOM implementation.

```ts
import { TsPptx } from 'pptx-ts'
import { tableToSlides } from 'pptx-ts/html'
import { Window } from 'happy-dom'

const win = new Window()
win.document.body.innerHTML = '<table id="report">…</table>'

const pptx = new TsPptx()
tableToSlides(pptx, win.document.getElementById('report'))
await pptx.writeFile({ fileName: 'report.pptx' })
```

Pass the element itself and no global DOM is consulted at all. To pass a string id
instead, say which document it belongs to:

```ts
tableToSlides(pptx, 'report', { document: win.document })
```

In a browser, `options.document` defaults to the global `document`, so
`tableToSlides(pptx, 'report')` is enough. The equivalent method form,
`pptx.tableToSlides('report', options)`, is on the browser build and delegates to the
same implementation.

Cell text (with `<br>` kept as a line break), `colspan`/`rowspan`, computed colors,
weight, alignment, padding, borders and auto-paging behave the same wherever the
conversion runs. Column widths are the one thing that depends on the runtime.

## Column widths need a layout engine

In a browser the columns are sized from each cell's rendered `offsetWidth`,
reproducing the table's real proportions. Nothing outside a browser lays a table out,
so `offsetWidth` is `0` there and the conversion falls back in two steps: it uses the
computed CSS `width`s when the stylesheet states them for every column in one unit
(all `px` or all `%`), and an equal split when it does not.

The first step is a *fallback*, not a graceful loss of precision. `offsetWidth` is the
border box and computed `width` the content box, so padding alone can put the two
bases in different proportions. The same table can therefore come out with different
column widths in a browser and outside one.

To pin widths regardless of runtime, annotate the `<thead>` header cells, which win
outright on every path:

```html
<thead>
  <tr>
    <th data-pptx-width="2.5">Name</th>
    <th data-pptx-min-width="1">Qty</th>
  </tr>
</thead>
```

Building a table from data you already hold is a different job: use
`addTable(rows, opts)` and skip the DOM entirely. See [Tables](tables.md).

## Where the line sits

Converting an HTML table is a supported, tested, portable path, covered end to end
against happy-dom in `test/regression/html/html-to-slides-node.test.js` and against a
real Chromium in `test/browser/table-widths.spec.mjs`. What a browser adds is
*measurement*, and measurement is the only thing lost without one.

Matching how a browser laid a page out is a separate question, and it stays outside
what this project actively develops: see
[Project Target](project-target.md#out-of-active-scope-contributions-welcome).
