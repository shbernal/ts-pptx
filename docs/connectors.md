---
doc-schema-version: 1
title: "Connectors"
summary: "Draw straight, elbow, and curved connectors between two points (or bound to shapes) with addConnector(), emitted as a PowerPoint connector (<p:cxnSp>)."
read_when:
  - Drawing a line or connector between two points on a slide
  - Making an elbow or curved connector, or controlling its bend positions
  - Binding a connector's endpoint to a shape so it reroutes when the shape moves
  - Choosing between addConnector() and addShape() for a connector geometry
doc_type: "guide"
---

# Connectors

A connector is a line drawn between two points, emitted as a PowerPoint connector
(`<p:cxnSp>`) rather than a plain line shape. Because it is a real connector,
PowerPoint treats it as selectable and reroutable, and — when its endpoints are
bound to shapes — reroutes it automatically as those shapes move.

```js
const s = pptx.addSlide()
s.addConnector({ x1: 1, y1: 1, x2: 5, y2: 3, endArrowType: 'triangle' })
```

`slide.addConnector(options)` requires the four endpoint coordinates
(`x1, y1, x2, y2`) and returns the slide, so calls chain. Everything else is
optional.

## Routing style: `type` and `bends`

`type` selects the routing style; for `elbow` / `curved`, `bends` selects the
preset variant (how many adjustable jogs it has):

| `type` | `bends` | OOXML preset | Jogs |
| --- | --- | --- | --- |
| `'straight'` (default) | — | `straightConnector1` | none |
| `'elbow'` | `1` (default) / `2` / `3` | `bentConnector3` / `4` / `5` | 1 / 2 / 3 |
| `'curved'` | `1` (default) / `2` / `3` | `curvedConnector3` / `4` / `5` | 1 / 2 / 3 |

```js
s.addConnector({ type: 'elbow',  x1: 1, y1: 1, x2: 5, y2: 3 })  // one-jog elbow
s.addConnector({ type: 'curved', x1: 1, y1: 4, x2: 5, y2: 6 })  // one-jog curve
```

`bends` is ignored for `type: 'straight'` (a straight connector has no bends);
passing `bends`/`adj` with a straight connector **warns** rather than silently
doing nothing.

## Bend positions: `adj`

`adj` places each jog as a percent of the connector box (`0`–`100`), one value per
bend. A single number sets the sole jog of a one-bend connector; an array sets
each jog of a multi-bend connector and its length **must equal** `bends`:

```js
s.addConnector({ type: 'elbow',  x1: 1, y1: 1, x2: 5, y2: 3, adj: 25 })
s.addConnector({ type: 'elbow',  x1: 1, y1: 1, x2: 5, y2: 3, bends: 2, adj: [30, 70] })
s.addConnector({ type: 'curved', x1: 1, y1: 4, x2: 5, y2: 6, bends: 3, adj: [10, 50, 90] })
```

Values are emitted as `<a:gd name="adjN" fmla="val …"/>` guides (OOXML
1000ths-of-a-percent, so `25` → `25000`). When omitted, PowerPoint uses the preset
default (50%). Out-of-range values are **allowed with a warning** — they place a
jog beyond the endpoint box, as PowerPoint itself does when endpoints flip. A
non-finite `adj`, or an `adj` array whose length does not match `bends`,
**throws** rather than emitting a degenerate guide.

## Endpoints and the derived box

You give the two endpoints directly; the bounding box is derived. The connector's
origin is the **min corner** of the two points, and `flipH` / `flipV` are set when
the end point is left of / above the start point — so the connector draws
correctly from any pair of endpoints, in any direction:

```js
// end is left of and above start → the box flips both ways
s.addConnector({ type: 'elbow', x1: 5, y1: 3, x2: 2, y2: 1 })
```

Each of `x1/y1/x2/y2` accepts any `Coord`: a number of inches, a percent string
like `'50%'`, or a unit string like `'2in'`.

Every other object gets the same treatment from the other direction: a negative
`w` or `h` is normalized to the min-corner origin, the absolute extent, and the
matching flip, so a shape placed from a signed delta is safe too.

```js
// h is negative — the line runs upward. Emitted as y=1in, cy=2in, flipV="1".
s.addShape('line', { x: 1, y: 3, w: 1.5, h: -2 })
```

A flip derived this way XOR-composes with one you set yourself, so
`{ w: -2, flipH: true }` is mirrored twice and therefore not mirrored at all.
This matters because `<a:ext>` is `ST_PositiveCoordinate`: a negative extent is
out of range, and PowerPoint rejects the whole package (0x80070570) rather than
the offending shape.

## Binding to shapes

Bind an endpoint to a shape on the **same slide** by that shape's `objectName`.
The connector then attaches to the shape: PowerPoint reroutes it when the shape
moves, and the elbow auto-router can engage.

```js
s.addShape('rect', { x: 1, y: 1, w: 2, h: 1, objectName: 'boxA' })
s.addShape('rect', { x: 6, y: 4, w: 2, h: 1, objectName: 'boxB' })
s.addConnector({
  type: 'elbow',
  x1: 3, y1: 1.5, x2: 6, y2: 4.5,        // static fallback geometry
  startShape: 'boxA', startShapeIdx: 3,   // -> <a:stCxn id=… idx=3>
  endShape: 'boxB',                       // -> <a:endCxn id=… idx=0>
})
```

- `startShape` / `endShape` name the target; `startShapeIdx` / `endShapeIdx` pick
  the connection site on that shape (0-based; a shape's `<a:cxnLst>` enumerates its
  sites — the valid range is preset-dependent, default `0`).
- The target's `<p:cNvPr>` id is resolved at serialize time, so the shape may be
  added before or after the connector.
- `x1/y1/x2/y2` remain the static fallback geometry and are used if a name can't be
  resolved. An **unresolved** name **warns** and emits an empty `<p:cNvCxnSpPr/>`
  rather than a dangling id; a negative connection-site index **throws**.

A shape **inside a group** is a valid target — group children are named on the same
slide (see [Grouping objects](groups.md#cross-references-into-a-group)). The
connector itself cannot be a group child.

## Line styling

```js
s.addConnector({
  x1: 1, y1: 1, x2: 5, y2: 3,
  color: 'FF0000',          // 6-digit hex, no '#' (default '000000')
  width: 2,                  // points (default 1)
  dashType: 'dash',
  beginArrowType: 'oval',
  endArrowType: 'triangle',
  objectName: 'Flow A→B',    // Selection Pane name
  altText: 'Flow from A to B',
})
```

## `addConnector()` vs. `addShape()` for connector geometry

There are two ways to place a connector geometry, and they are not the same object:

- **`addConnector()`** emits a live `<p:cxnSp>` — a real connector that is
  selectable, reroutable, and can bind to shapes. This is almost always what you
  want. It reaches `straightConnector1`, `bentConnector3/4/5`, and
  `curvedConnector3/4/5`.
- **`addShape('<preset>', …)`** emits a static `<p:sp>` positioned by a box.
  The raw connector preset names are part of the public `SHAPE_NAME` union, so you
  can pass them as ordinary shape geometry — but the result is a fixed drawing, not
  a connector: it does not reroute and cannot bind to shapes.

The single-segment variants **`bentConnector2`** and **`curvedConnector2`** (a
connector with a single fixed corner and **no** adjustable jog) are reachable only
through `addShape` as static geometry; `addConnector`'s `bends` starts at `1`,
which maps to the `…Connector3` variants.

## See also

- API reference: [`ConnectorProps`](reference/api/index/interfaces/ConnectorProps.md),
  [`ConnectorType`](reference/api/index/type-aliases/ConnectorType.md).
- Where it lives in the pipeline: [Architecture](architecture.md) —
  `gen/define/connector.ts` `addConnectorDefinition` (add) →
  `gen/slide/object.ts` `slideObjectToXml` (emit `<p:cxnSp>`).
- Regression coverage: `test/regression/connector-shape.test.js`.
- Schema fixture: `connectors (straight/elbow/curved, flipped, arrowheads)` in
  `test/schema-cases.js`.
