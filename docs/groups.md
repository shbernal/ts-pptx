---
doc-schema-version: 1
title: "Grouping Objects"
summary: "Wrap slide objects in a PowerPoint group (<p:grpSp>) with addGroup() or groupObjects(), including nesting, framing, and cross-references."
read_when:
  - Grouping shapes/text/images into one selectable PowerPoint group
  - Grouping objects that were already added to a slide (groupObjects)
  - Listing what a slide already holds, to decide what to group (slide.objects)
  - Understanding why a group never moves or scales its children
  - Binding a connector or animation to a shape inside a group
doc_type: "guide"
---

# Grouping Objects

A group is a single selectable PowerPoint object (`<p:grpSp>`) that contains other
objects. ts-pptx offers two entry points:

- `slide.addGroup(children, options?)`:build a group from child *descriptors*
  (the same shorthand `addShape`/`addText`/`addImage` accept), in one call.
- `slide.groupObjects(objectNames, options?)`:wrap objects that are **already on
  the slide**, addressed by their `objectName`. Use this when independent renderers
  each added their own objects and you want to group them after the fact without
  replaying their descriptors.

Both produce the same XML and share the same framing and naming rules.
`slide.objects` reports what is already there, which is how a caller that did not
author the objects learns the names `groupObjects()` takes.

## The identity child space (why a group never moves its children)

Every group ts-pptx writes keeps an **identity child coordinate space**:
`chOff/chExt == off/ext` at every nesting depth. The practical consequence:

> A group's own frame only places the selection handle and the rotate pivot. It
> never moves or scales the children: they keep their slide-absolute `x/y/w/h`.

This makes grouping visually a no-op: the objects render exactly where they did
before, but PowerPoint now treats them as one unit (one Selection Pane entry, one
drag target, one rotate handle). It also means you never recompute child
coordinates to "put them inside" a group: you author each child at its final
slide position and let the group wrap it.

## `addGroup()`: build a group from descriptors

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
is pending). A group left with no renderable children: e.g. because every child
was an unsupported kind: warns and emits a degenerate zero-size group rather than
silently producing one.

### Framing: auto-bounds or an explicit frame (all-or-nothing)

The group frame is **all-or-nothing**:

- **Omit all of `x/y/w/h`** → the frame auto-computes to the bounding box of the
  children (recursing into nested auto-sized groups). This is the usual case.
- **Pass all four** → the frame is used verbatim. Because the child space stays
  identity, this only relocates the handle/pivot; the children do not move.

A **partial** frame (some axes set, others not) is ambiguous: `{ x: 5 }` reads
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

## `groupObjects()`: group objects already on the slide

```js
const s = pptx.addSlide()
s.addShape('rect', { x: 1, y: 1, w: 2, h: 1, objectName: 'Header' })
s.addText('Caption', { x: 1.2, y: 2.2, w: 1.6, h: 0.6, objectName: 'Caption' })
s.addShape('rect', { x: 5, y: 1, w: 1, h: 1, objectName: 'Loose' })

s.groupObjects(['Header', 'Caption'], { objectName: 'Banner' })
// 'Header' and 'Caption' are now inside 'Banner'; 'Loose' stays top-level.
```

Two ordering rules keep the lift visually a no-op:

- **Children keep their existing slide z-order**: *not* the order you name them.
  Naming is a selection, not a restack; `['Top', 'Bottom']` never lifts `Top`
  above `Bottom`.
- **The wrapper takes the topmost member's former slot**: it sits above everything
  the selection sat above and below everything it sat below. A non-member that sat
  between two members surfaces above the group.

An existing group can be a member, so you can compose larger logical groups out of
groups you already made.

### Failure is a throw, not a warn

Unlike `addGroup` (which warns and skips), every `groupObjects` failure **throws**,
because each one would otherwise leave the intended object silently loose on the
slide: the footgun the group was meant to remove. Resolution runs fully *before*
anything moves, so a bad name leaves the slide untouched rather than half-grouped:

- a name no top-level object has (distinguished from one that is already inside
  another group),
- an ambiguous name shared by more than one object,
- an ungroupable kind (`chart`/`table`/`media`/`placeholder`),
- an empty or duplicate-laden `objectNames` array.

## `slide.objects`: what is on the slide, and what can be grouped

`groupObjects()` addresses objects by name, so something has to know the names.
When one caller authored everything, that is easy. When a slide is assembled by
independent renderers, it is not: nobody kept the descriptors, and a parallel
ledger of what was added is wrong the moment a renderer adds an object it did not
announce. `slide.objects` is the read-back half.

```js
const s = pptx.addSlide()
renderHeader(s) // adds 'card:header', 'card:header_icon'
renderBody(s) // adds 'card:body', 'card:body_label'

for (const o of s.objects) {
  console.log(o.type, o.objectName, o.canGroup)
}
// text  card:header       true
// image card:header_icon  true
// text  card:body         true
// text  card:body_label   true

const card = s.objects.filter((o) => o.canGroup && o.objectName.startsWith('card:'))
s.groupObjects(
  card.map((o) => o.objectName),
  { objectName: 'Card' }
)
```

Each entry is a `SlideObjectInfo`:

| field           | meaning                                                                |
| --------------- | ---------------------------------------------------------------------- |
| `type`          | the kind the object was authored as (a shape is a `text` object)        |
| `objectName`    | the Selection Pane name, in the spelling that resolves (see below)      |
| `isPlaceholder` | occupies a layout placeholder, which grouping refuses on top of kind    |
| `canGroup`      | whether `groupObjects()` accepts this object's kind                     |
| `children`      | a group's members, same shape, nested to any depth; empty for a leaf    |

Three things are worth knowing:

- **The order is z-order**, bottom to top, the same order the objects were added
  and the same order `groupObjects()` uses when it stacks a group's children.
- **`objectName` is always a string.** An object authored without one still gets
  the generated `Shape 3` / `Text 1` / `Group 2` identity PowerPoint shows in the
  Selection Pane, and that name addresses it just as well. Nothing records which
  is which, so if the distinction matters, make it with your own naming
  convention.
- **`canGroup` answers about the object, not the call.** It is the same predicate
  `groupObjects()` throws on, so it will not drift from the groupable-kinds list.
  The remaining failures are about the *selection* (a name that resolves to
  nothing, or to two objects), which no single object can speak for.

It is a snapshot, not a live handle: a fresh array on every access, describing the
slide as it was, and writing to it does nothing. To change the slide, call the
authoring API with the names it gave you.

### Why the name comes back decoded

`objectName` is stored attribute-escaped, and `groupObjects()` escapes the
caller's spelling before it compares. So `slide.objects` reports the *caller's*
spelling: a shape named `Q&A` reads back as `Q&A`, not the stored `Q&amp;A`,
which would escape a second time on the way in and resolve to nothing.

The guarantee is the round trip, not invertibility: a name that comes out of
`slide.objects` goes back into `groupObjects()` and finds its object. That holds
even for a name that is itself an entity spelling (`&amp;` authored literally).

## Cross-references into a group

A connector or animation can target a shape **inside** a group by its `objectName`:
group children are `<p:cNvPr>`-named on the same slide and are valid targets:

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

Grouped text still participates in the export-time measured-fit pass: the fit pass
descends into groups, and because the child space is identity, a grouped text box's
authored `w/h` is its true rendered size (see `docs/measured-text-fit.md`).

## See also

- Runnable demo: the KPI cards in `demos/showcases/quarterly-review/index.mjs`
  (run with `pnpm demos:build quarterly-review`).
- Regression coverage: `test/regression/shape/group-shapes.test.js`.
- Schema fixtures: `flat-group`, `nested-group`, `group-cross-references`,
  `group-existing-objects` in `test/schema-cases.js`.
