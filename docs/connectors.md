---
doc-schema-version: 1
title: "Connectors"
summary: "Draw straight, elbow and curved connectors between two points with addConnector(), place their bends, and attach their ends to shapes by objectName."
read_when:
  - Drawing a line or arrow between two points on a slide
  - Making an elbow or curved connector, or setting where it bends
  - Attaching a connector end to a shape so PowerPoint keeps it attached
  - Choosing between addConnector() and addShape() for connector geometry
doc_type: "guide"
---

# Connectors

`slide.addConnector()` draws a line between two points as a PowerPoint connector, which can attach to a shape at either end.

```ts
import TsPptx from "pptx-ts"

const pptx = new TsPptx()
const slide = pptx.addSlide()
slide.addConnector({ x1: 1, y1: 1, x2: 5, y2: 3, endArrowType: "triangle" })
await pptx.writeFile({ fileName: "connector.pptx" })
```

The four endpoint coordinates are required and everything else is optional. The call returns the slide, so calls chain.

## Options at a glance

| Option | Type | Default | Effect |
| --- | --- | --- | --- |
| `x1`, `y1`, `x2`, `y2` | `Coord` | required | Start point and end point |
| `type` | `ConnectorType` | `"straight"` | Routing style: `"straight"`, `"elbow"` or `"curved"` |
| `bends` | `1`, `2` or `3` | the length of `adj`, else `1` | How many adjustable bends an elbow or curved connector has |
| `adj` | `number` or `number[]` | PowerPoint's 50% | Where each bend sits, as a percent of the connector's box |
| `startShape`, `endShape` | `string` | none | `objectName` of the shape that end attaches to |
| `startShapeIdx`, `endShapeIdx` | `number` | `0` | Connection site on that shape, counting from 0 |
| `color` | `HexColor` | `"333333"` | Line colour |
| `width` | `number` | `1` | Line width in points. `0` also gives 1 |
| `dashType` | dash style | `"solid"` | Dash pattern |
| `beginArrowType`, `endArrowType` | `LineEndType` | none | Arrowhead at the start or the end |
| `objectName` | `string` | `Connector 1`, `Connector 2`, ... | Selection Pane name |
| `altText` | `string` | none | Alt text |

## Pick a routing style

`type` picks the routing style and `bends` picks how many adjustable bends it has. Together they choose the preset PowerPoint draws:

| `type` | `bends: 1` | `bends: 2` | `bends: 3` |
| --- | --- | --- | --- |
| `"straight"` | `straightConnector1` | n/a | n/a |
| `"elbow"` | `bentConnector3` | `bentConnector4` | `bentConnector5` |
| `"curved"` | `curvedConnector3` | `curvedConnector4` | `curvedConnector5` |

```ts
slide
  .addConnector({ type: "elbow", x1: 1, y1: 1, x2: 5, y2: 3 })
  .addConnector({ type: "curved", x1: 1, y1: 4, x2: 5, y2: 6, bends: 2 })
```

A straight connector has no bends. Passing `bends` or `adj` with one warns and changes nothing.

## Place the bends

`adj` takes one percent per bend, measured across the connector's box. A single number sets the only bend of a one-bend connector. An array sets each bend in order.

```ts
slide.addConnector({ type: "elbow", x1: 1, y1: 1, x2: 5, y2: 3, adj: 25 })
slide.addConnector({ type: "elbow", x1: 1, y1: 1, x2: 5, y2: 3, bends: 2, adj: [30, 70] })
slide.addConnector({ type: "curved", x1: 1, y1: 4, x2: 5, y2: 6, adj: [10, 50, 90] }) // three bends
```

<svg role="img" aria-label="Two one-bend elbow connectors drawn in the same box. With adj 25 the vertical segment sits a quarter of the way across the box. With adj omitted it sits halfway." viewBox="0 0 520 165" width="100%" style="max-width: 520px" fill="none" stroke="currentColor" font-size="12">
  <rect x="40" y="30" width="180" height="80" stroke-dasharray="4 3" opacity="0.6"/>
  <polyline points="40,30 85,30 85,110 220,110" stroke-width="2"/>
  <line x1="40" y1="124" x2="85" y2="124"/>
  <rect x="300" y="30" width="180" height="80" stroke-dasharray="4 3" opacity="0.6"/>
  <polyline points="300,30 390,30 390,110 480,110" stroke-width="2"/>
  <line x1="300" y1="124" x2="390" y2="124"/>
  <g fill="currentColor" stroke="none" text-anchor="middle">
    <text x="62" y="140">25%</text>
    <text x="130" y="158">adj: 25</text>
    <text x="345" y="140">50%</text>
    <text x="390" y="158">adj omitted</text>
  </g>
</svg>

| `adj` | Outcome | Code |
| --- | --- | --- |
| omitted | accepted, and PowerPoint puts each bend at 50% | none |
| one number on a one-bend connector | accepted | none |
| an array as long as `bends` | accepted, one value per bend | none |
| an array, with `bends` omitted | accepted, and `bends` becomes the array's length | none |
| a value below 0 or above 100 | warns, and the bend sits outside the box | `connector/adj-out-of-range` |
| any `adj` or `bends` on a straight connector | warns, and both are ignored | `connector/bends-ignored-for-straight` |
| a count that differs from `bends` | throws | `connector/adj-count-mismatch` |
| `NaN`, `Infinity`, or a value that is not a number | throws | `connector/adj-non-finite` |
| more than three values, with `bends` omitted | throws | `connector/invalid-bends` |

## Place the endpoints

You give the two endpoints and ts-pptx derives the box PowerPoint stores. The box starts at the smaller `x` and the smaller `y` of the two points. When the end point is left of the start point the connector is flipped horizontally, and when it is above the start point it is flipped vertically. Any pair of points, in any direction, draws correctly.

<svg role="img" aria-label="Two straight connectors and the boxes derived from them. Left: the end point is right of and below the start point, so the box origin is the start point and nothing is flipped. Right: the end point is left of and above the start point, so the box origin is the end point and the connector is flipped both horizontally and vertically." viewBox="0 0 520 190" width="100%" style="max-width: 520px" fill="none" stroke="currentColor" font-size="12">
  <rect x="40" y="40" width="180" height="90" stroke-dasharray="4 3" opacity="0.6"/>
  <line x1="40" y1="40" x2="220" y2="130" stroke-width="2"/>
  <rect x="300" y="40" width="180" height="90" stroke-dasharray="4 3" opacity="0.6"/>
  <line x1="480" y1="130" x2="300" y2="40" stroke-width="2"/>
  <g fill="currentColor" stroke="none">
    <circle cx="40" cy="40" r="4"/>
    <circle cx="220" cy="130" r="4"/>
    <circle cx="300" cy="40" r="4"/>
    <circle cx="480" cy="130" r="4"/>
    <text x="40" y="28">start (x1, y1), box origin</text>
    <text x="220" y="148" text-anchor="end">end (x2, y2)</text>
    <text x="130" y="178" text-anchor="middle">end right of and below start: no flip</text>
    <text x="300" y="28">end (x2, y2), box origin</text>
    <text x="480" y="148" text-anchor="end">start (x1, y1)</text>
    <text x="390" y="178" text-anchor="middle">end left of and above start: flipH, flipV</text>
  </g>
</svg>

```ts
// The end is left of and above the start, so the connector is flipped both ways
slide.addConnector({ type: "elbow", x1: 5, y1: 3, x2: 2, y2: 1 })

// Percentages: x of the slide width, y of the slide height
slide.addConnector({ x1: "10%", y1: "50%", x2: "90%", y2: "50%" })
```

- Each coordinate takes any `Coord`: inches as a number, or a string such as `"50%"`, `"2in"` or `"72pt"`. [Core concepts](getting-started/concepts.md#positions-and-sizes) lists the units.
- The endpoints are the only source of direction. A connector has no `rotate`, `flipH` or `flipV` option.
- Other objects are placed by `x`, `y`, `w` and `h`, and a negative `w` or `h` gets the same origin and flip treatment. See [Positions and sizes](reference/layout-units.md#positions-and-sizes).

## Bind a connector to a shape

Name a shape on the same slide in `startShape` or `endShape`. PowerPoint then treats that end as attached to the shape, and reroutes the connector when the shape moves.

```ts
slide.addShape("rect", { x: 1, y: 1, w: 2, h: 1, objectName: "boxA" })
slide.addShape("rect", { x: 6, y: 4, w: 2, h: 1, objectName: "boxB" })
slide.addConnector({
  type: "elbow",
  x1: 3, y1: 1.5, x2: 6, y2: 4.5,
  startShape: "boxA", startShapeIdx: 3,
  endShape: "boxB", endShapeIdx: 1,
})
```

- Names resolve when the deck is written, so the shape can be added before or after the connector.
- `startShapeIdx` and `endShapeIdx` pick a connection site on the shape, counting from 0. The shape's preset decides which sites exist. A custom geometry shape lists its own in `connectionSites`, in the order the index counts.
- `x1`, `y1`, `x2` and `y2` are still written as the connector's geometry, and are all an end has when its name does not resolve.
- A shape inside a group resolves like a top-level one. The lookup checks top-level objects from the bottom of the stack up, then group children, and the first object with the name wins. Animations that target an `objectName` use the same lookup.

| Condition | Outcome | Code |
| --- | --- | --- |
| no object on the slide has the name | warns when the deck is written, and that end is written unattached | `connector/unresolved-binding` |
| the name belongs to an object on another slide | the same as no match | `connector/unresolved-binding` |
| two objects on the slide share the name | the first in lookup order wins, and the duplicate warns when the deck is written | `object-name/duplicate` |
| `startShape` or `endShape` is empty or only whitespace | throws | `connector/invalid-binding-name` |
| a site index is negative or not an integer | throws | `connector/invalid-connection-site` |
| a site index is past the shape's last site | not checked | none |

## Style the line

```ts
slide.addConnector({
  x1: 1, y1: 1, x2: 5, y2: 3,
  color: "C00000",
  width: 2,
  dashType: "dash",
  beginArrowType: "oval",
  endArrowType: "triangle",
  objectName: "Flow A to B",
  altText: "Flow from A to B",
})
```

A connector's line takes a colour, a width, a dash pattern and an arrowhead at each end. It has no transparency, gradient or pattern option.

## Choose between addConnector() and addShape()

The connector presets are also shape names, so `addShape()` accepts them too. It produces a different kind of object:

| | `addConnector()` | `addShape()` with a connector preset |
| --- | --- | --- |
| Written as | a connector (`<p:cxnSp>`) | an ordinary shape (`<p:sp>`) |
| Placed by | two endpoints | `x`, `y`, `w` and `h` |
| Attaches to shapes | yes, with `startShape` and `endShape` | no |
| Rerouted by PowerPoint when a shape moves | yes, when attached | no |
| Presets | `straightConnector1`, `bentConnector3` to `bentConnector5`, `curvedConnector3` to `curvedConnector5` | the same, plus `bentConnector2` and `curvedConnector2` |

```ts
slide.addShape("bentConnector2", { x: 1, y: 1, w: 3, h: 2, line: { color: "000000" } })
```

`bentConnector2` and `curvedConnector2` have one fixed corner and no adjustable bend. `addConnector()` starts at one bend, so only `addShape()` reaches them. Use `addConnector()` whenever the line should stay attached to shapes.

## Invalid input

Throws happen inside the `addConnector()` call, as `InvalidOptionError`. The rows marked "when written" are reported while the deck is exported.

| Condition | Warns or throws | Code |
| --- | --- | --- |
| any of `x1`, `y1`, `x2`, `y2` is missing | throws | `connector/missing-endpoints` |
| `type` is not `"straight"`, `"elbow"` or `"curved"` | throws | `connector/invalid-type` |
| `bends` is not 1, 2 or 3 | throws | `connector/invalid-bends` |
| the `adj` count differs from `bends` | throws | `connector/adj-count-mismatch` |
| an `adj` value is not a finite number | throws | `connector/adj-non-finite` |
| an `adj` value is outside 0 to 100 | warns | `connector/adj-out-of-range` |
| `bends` or `adj` is set on a straight connector | warns | `connector/bends-ignored-for-straight` |
| `startShape` or `endShape` is empty | throws | `connector/invalid-binding-name` |
| a site index is negative or fractional | throws | `connector/invalid-connection-site` |
| a bound name resolves to nothing (when written) | warns | `connector/unresolved-binding` |
| `objectName` is only whitespace, longer than 255 characters, or holds control characters | warns | `object-name/empty`, `object-name/too-long`, `object-name/control-characters` |
| an `objectName` appears twice on the slide (when written) | warns | `object-name/duplicate` |

[Errors and warnings](errors-and-warnings.md) covers the error classes and how to route warnings.

## Limits

- Direction comes from the endpoints only. There is no `rotate`, `flipH` or `flipV` option.
- A connector has at most three adjustable bends.
- Arrowhead size cannot be set.
- A binding reaches objects on the same slide only.
- ts-pptx checks neither the target's kind nor whether the site index exists on it.
- A site index without its shape name is ignored.
- `addGroup()` has no connector child. [`groupObjects()`](groups.md#group-objects-already-on-a-slide) can group a connector that is already on the slide.

## Reading it back

- Before export, [`slide.objects`](groups.md#list-what-a-slide-holds) lists a connector with `type: "connector"`.
- After loading a deck with `pptx-ts/read`, a connector is a `Connector`. `startConnection` and `endConnection` each return a `ConnectionSite` with `shapeId`, `siteIndex` and `boundShape`, or `null` for an unattached end. See [Connector endpoint binding](reference/pptx-read.md#connector-endpoint-binding).
- [PPTX inspection](reference/pptx-inspection.md) reports connectors with the rest of a slide's elements, each with a slide-absolute box.

## See also

- [Groups](groups.md)
- [Positions and sizes](reference/layout-units.md#positions-and-sizes)
- [Core concepts](getting-started/concepts.md)
- [Errors and warnings](errors-and-warnings.md)
- API reference: [`ConnectorProps`](reference/api/index/interfaces/ConnectorProps.md), [`ConnectorType`](reference/api/index/type-aliases/ConnectorType.md), [`Slide`](reference/api/index/interfaces/Slide.md), [`SHAPE_NAME`](reference/api/index/type-aliases/SHAPE_NAME.md)
