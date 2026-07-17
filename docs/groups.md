---
doc-schema-version: 1
title: "Grouping Objects"
summary: "Wrap slide objects in a PowerPoint group (<p:grpSp>) with addGroup() or groupObjects(), including nesting, framing, and cross-references."
read_when:
  - Grouping shapes/text/images into one selectable PowerPoint group
  - Grouping objects that were already added to a slide (groupObjects)
  - Understanding why a group never moves or scales its children
  - Binding a connector or animation to a shape inside a group
doc_type: "guide"
---

# Grouping Objects

A group is a single selectable PowerPoint object (`<p:grpSp>`) that contains other
objects. PptxGenJS offers two entry points:

- `slide.addGroup(children, options?)` — build a group from child *descriptors*
  (the same shorthand `addShape`/`addText`/`addImage` accept), in one call.
- `slide.groupObjects(objectNames, options?)` — wrap objects that are **already on
  the slide**, addressed by their `objectName`. Use this when independent renderers
  each added their own objects and you want to group them after the fact without
  replaying their descriptors.

Both produce the same XML and share the same framing and naming rules.

## The identity child space (why a group never moves its children)

Every group PptxGenJS writes keeps an **identity child coordinate space**:
`chOff/chExt == off/ext` at every nesting depth. The practical consequence:

> A group's own frame only places the selection handle and the rotate pivot. It
> never moves or scales the children — they keep their slide-absolute `x/y/w/h`.

This makes grouping visually a no-op: the objects render exactly where they did
before, but PowerPoint now treats them as one unit (one Selection Pane entry, one
drag target, one rotate handle). It also means you never recompute child
coordinates to "put them inside" a group — you author each child at its final
slide position and let the group wrap it.

## `addGroup()` — build a group from descriptors

```js
slide.addGroup(
  [
    { rect: { x: 1, y: 1, w: 2, h: 1, fill: { color: 'CC0000' } } },
    { text: { text: 'Label', options: { x: 1.2, y: 1.2, w: 1.6, h: 0.6, color: 'FFFFFF' } } },
    { image: { path: 'logo.png', x: 3.5, y: 1, w: 1, h: 1 } },
  ],
  { objectName: 'Branding' }
)
```

Supported child descriptors: `rect`, `roundRect`, `line`, `shape` (any preset),
`text`, `image`, and `group` (nesting). Each child is authored in slide-absolute
inches, exactly as the top-level `add*` methods take them.

**Not supported as children yet:** `chart`, `table`, `media`, and `placeholder`.
Each is skipped with a warning (the relationship/id/transform plumbing to nest them
is pending). A group left with no renderable children — e.g. because every child
was an unsupported kind — warns and emits a degenerate zero-size group rather than
silently producing one.

### Framing: auto-bounds or an explicit frame (all-or-nothing)

The group frame is **all-or-nothing**:

- **Omit all of `x/y/w/h`** → the frame auto-computes to the bounding box of the
  children (recursing into nested auto-sized groups). This is the usual case.
- **Pass all four** → the frame is used verbatim. Because the child space stays
  identity, this only relocates the handle/pivot; the children do not move.

A **partial** frame (some axes set, others not) is ambiguous — `{ x: 5 }` reads
like a reposition but cannot be one without moving the children out from under the
box. So a partial frame **warns and falls back to auto-bounds** on every axis. Pass
all four or none.

### Rotate, flip, lock, and alt text

`GroupProps` accepts the same identity/formatting options as other objects; they
apply to the group as a whole:

```js
slide.addGroup([{ rect: { x: 1, y: 1, w: 2, h: 1 } }], {
  rotate: 45,          // whole group rotates about its pivot
  flipH: true,
  flipV: true,
  objectName: 'Badge',
  altText: 'Award badge',           // -> group cNvPr @descr
  objectLock: { noMove: true, noResize: true }, // -> a:grpSpLocks
})
```

Group locks are the group-valid subset (`noGrp`, `noSelect`, `noRot`,
`noChangeAspect`, `noMove`, `noResize`); a flag valid only on shapes/pictures
(e.g. `noCrop`) is dropped with a warning rather than silently coerced.

### Nesting

A `group` child nests. Names number **per slide** and **inside-out** (a nested
group built first takes the lower `Group N` index):

```js
slide.addGroup([
  { rect: { x: 1, y: 1, w: 1, h: 1 } },
  { group: { children: [
    { rect: { x: 3, y: 1, w: 1, h: 1 } },
    { text: { text: 'Nested', options: { x: 3, y: 1, w: 1, h: 1 } } },
  ], options: { objectName: 'Inner' } } },
], { objectName: 'Outer' })
```

## `groupObjects()` — group objects already on the slide

```js
const s = pptx.addSlide()
s.addShape('rect', { x: 1, y: 1, w: 2, h: 1, objectName: 'Header' })
s.addText('Caption', { x: 1.2, y: 2.2, w: 1.6, h: 0.6, objectName: 'Caption' })
s.addShape('rect', { x: 5, y: 1, w: 1, h: 1, objectName: 'Loose' })

s.groupObjects(['Header', 'Caption'], { objectName: 'Banner' })
// 'Header' and 'Caption' are now inside 'Banner'; 'Loose' stays top-level.
```

Two ordering rules keep the lift visually a no-op:

- **Children keep their existing slide z-order** — *not* the order you name them.
  Naming is a selection, not a restack; `['Top', 'Bottom']` never lifts `Top`
  above `Bottom`.
- **The wrapper takes the topmost member's former slot** — it sits above everything
  the selection sat above and below everything it sat below. A non-member that sat
  between two members surfaces above the group.

An existing group can be a member, so you can compose larger logical groups out of
groups you already made.

### Failure is a throw, not a warn

Unlike `addGroup` (which warns and skips), every `groupObjects` failure **throws**,
because each one would otherwise leave the intended object silently loose on the
slide — the footgun the group was meant to remove. Resolution runs fully *before*
anything moves, so a bad name leaves the slide untouched rather than half-grouped:

- a name no top-level object has (distinguished from one that is already inside
  another group),
- an ambiguous name shared by more than one object,
- an ungroupable kind (`chart`/`table`/`media`/`placeholder`),
- an empty or duplicate-laden `objectNames` array.

## Cross-references into a group

A connector or animation can target a shape **inside** a group by its `objectName`
— group children are `<p:cNvPr>`-named on the same slide and are valid targets:

```js
s.addGroup([{ rect: { x: 1, y: 1, w: 2, h: 1, objectName: 'boxInGroup' } }], { objectName: 'Grp' })
s.addConnector({ type: 'elbow', x1: 3, y1: 1.5, x2: 6, y2: 4.5, startShape: 'boxInGroup' })
s.addAnimation({ preset: 'fadeIn', objectName: 'boxInGroup' })
```

If two objects share a name, a top-level object wins the reference over a group
child (duplicate names are warned about separately). A reference naming nothing on
the slide warns rather than emitting a dangling target.

## Reading groups back

On the read side, a `<p:grpSp>` surfaces as a `GroupShape` whose `.shapes` are its
children; each child's `absoluteFrame` composes the enclosing group transforms
(scale, rotation, flips) to report its true slide-absolute geometry. `inspect()`
reports a `group` element with its children and applies the same composition. See
[PPTX read API](reference/pptx-read.md) and
[PPTX Inspection](reference/pptx-inspection.md).

Grouped text still participates in the export-time measured-fit pass — the fit pass
descends into groups, and because the child space is identity, a grouped text box's
authored `w/h` is its true rendered size (see `docs/measured-text-fit.md`).

## See also

- Runnable demo: `demos/modules/demo_group.mjs` (run with `node demos/node/demo.js Group`).
- Regression coverage: `test/regression/group-shapes.test.js`.
- Schema fixtures: `flat-group`, `nested-group`, `group-cross-references`,
  `group-existing-objects` in `test/schema.test.js`.
