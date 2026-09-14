---
doc-schema-version: 1
title: "Groups"
summary: "Combine slide objects into one PowerPoint group with addGroup() or groupObjects(), list a slide's objects with slide.objects, and control a group's frame, nesting and stacking order."
read_when:
  - Grouping shapes, text and images into one selectable PowerPoint group
  - Grouping objects already on a slide by their objectName (groupObjects)
  - Listing a slide's objects to decide what to group (slide.objects)
  - Setting a group's frame, or finding out why a group never moves its children
  - Pointing a connector or animation at a shape inside a group
doc_type: "guide"
---

# Groups

A group turns several slide objects into one PowerPoint object, with one Selection Pane entry, one drag target and one rotation handle.

```ts
import TsPptx from "pptx-ts"

const pptx = new TsPptx()
const slide = pptx.addSlide()
slide.addGroup(
  [
    { rect: { x: 1, y: 1, w: 2, h: 1, fill: { color: "CC0000" } } },
    { text: { text: "Label", options: { x: 1.2, y: 1.2, w: 1.6, h: 0.6, color: "FFFFFF" } } },
  ],
  { objectName: "Badge" },
)
await pptx.writeFile({ fileName: "group.pptx" })
```

`slide.addGroup(children, options?)` builds a group from child descriptors. `slide.groupObjects(objectNames, options?)` groups objects that are already on the slide. Both take the same options and return the slide.

## Options at a glance

| Option | Type | Default | Effect |
| --- | --- | --- | --- |
| `x`, `y`, `w`, `h` | `Coord` | the children's bounding box | The group frame. Pass all four or none |
| `rotate` | `number` | none | Rotates the whole group, in degrees |
| `flipH`, `flipV` | `boolean` | `false` | Mirrors the whole group |
| `objectName` | `string` | `Group 1`, `Group 2`, ... | Selection Pane name |
| `altText` | `string` | none | Alt text for the group |
| `objectLock` | `ObjectLockProps` | none | Locks. A group takes `noGrp`, `noSelect`, `noRot`, `noChangeAspect`, `noMove` and `noResize` |

## The identity child space

A group never moves or scales its children. Each child stays at the slide position you gave it, and the group frame only sets the selection box and the centre the group rotates about.

<svg role="img" aria-label="Left: a group with no frame given. Its dashed frame is the bounding box of a rectangle and a circle. Right: the same rectangle and circle at the same positions, with an explicit frame drawn larger around them. The children have not moved." viewBox="0 0 520 200" width="100%" style="max-width: 520px" fill="none" stroke="currentColor" font-size="12">
  <rect x="50" y="50" width="90" height="40" stroke-width="2"/>
  <circle cx="175" cy="95" r="25" stroke-width="2"/>
  <rect x="50" y="50" width="150" height="70" stroke-dasharray="4 3" opacity="0.6"/>
  <rect x="320" y="50" width="90" height="40" stroke-width="2"/>
  <circle cx="445" cy="95" r="25" stroke-width="2"/>
  <rect x="295" y="30" width="200" height="125" stroke-dasharray="4 3" opacity="0.6"/>
  <g fill="currentColor" stroke="none" text-anchor="middle">
    <text x="125" y="175">no frame given:</text>
    <text x="125" y="191">the frame is the bounding box</text>
    <text x="395" y="175">all four of x, y, w, h given:</text>
    <text x="395" y="191">the children stay where they are</text>
  </g>
</svg>

So you place every child at its final slide position, and grouping leaves the slide looking exactly as it did. [Text that fits](text-fit.md) measures text inside a group at its authored size, like any other text box.

## Build a group from descriptors

```ts
slide.addGroup(
  [
    { roundRect: { x: 1, y: 1, w: 3, h: 1.5, fill: { color: "1F3A5F" } } },
    { shape: { type: "ellipse", options: { x: 1.2, y: 1.2, w: 1, h: 1, fill: { color: "FFFFFF" } } } },
    { image: { path: "logo.png", x: 2.5, y: 1.25, w: 1, h: 1 } },
    { line: { x: 1, y: 2.7, w: 3, h: 0, line: { color: "1F3A5F" } } },
  ],
  { objectName: "Logo lockup" },
)
```

| Descriptor | Adds |
| --- | --- |
| `{ rect: ShapeProps }`, `{ roundRect: ShapeProps }`, `{ line: ShapeProps }` | a shape of that preset |
| `{ shape: { type, options } }` | a shape of any preset in `SHAPE_NAME` |
| `{ text: { text, options } }` | a text box |
| `{ image: ImageProps }` | an image |
| `{ group: { children, options } }` | a nested group |

- Each child takes slide coordinates, the same options its `add*` method takes.
- The array order is the stacking order. The first child is at the bottom.
- Charts, tables, media and placeholders cannot be group children. The type rejects them, and at run time each one warns and is skipped.

## Set the group frame

| You pass | Frame written |
| --- | --- |
| none of `x`, `y`, `w`, `h` | the bounding box of the children as they are drawn, nested groups included |
| all four | that box, exactly |
| one, two or three of them | a `group/partial-frame` warning, then the bounding box |

```ts
// The frame is larger than the rectangle, and the rectangle stays at x 1, y 1
slide.addGroup([{ rect: { x: 1, y: 1, w: 2, h: 1 } }], { x: 0.5, y: 0.5, w: 3, h: 2 })
```

- A partial frame such as `{ x: 5 }` reads like a move, and a group frame cannot move its children, so ts-pptx falls back to the bounding box rather than draw the frame away from its content.
- The bounding box uses the size each child is drawn at. A text box with no `w` counts at its default width, and an image with a `sizing` box counts at that box.
- A negative `w` or `h` on an explicit frame becomes a positive size plus a flip. See [Positions and sizes](reference/layout-units.md#positions-and-sizes).

## Rotate, flip and lock a group

```ts
slide.addGroup([{ rect: { x: 1, y: 1, w: 2, h: 1 } }], {
  rotate: 45,
  flipH: true,
  objectName: "Badge",
  altText: "Award badge",
  objectLock: { noMove: true, noResize: true },
})
```

- `rotate`, `flipH` and `flipV` apply to the whole group, about the centre of its frame.
- A lock flag that groups do not support, such as `noCrop`, warns and is left out.

## Nest groups

```ts
slide.addGroup(
  [
    { rect: { x: 1, y: 1, w: 1, h: 1 } },
    {
      group: {
        children: [
          { rect: { x: 3, y: 1, w: 1, h: 1 } },
          { text: { text: "Nested", options: { x: 3, y: 1, w: 1, h: 1 } } },
        ],
        options: { objectName: "Inner" },
      },
    },
  ],
  { objectName: "Outer" },
)
```

- A `group` child nests to any depth, and every level keeps its children where they were placed.
- Default names count per slide, and a nested group takes its number before the group around it. An unnamed group holding an unnamed group, on a slide with no other groups, writes `Group 2` outside and `Group 1` inside.
- `groupObjects()` also accepts an existing group as a member.

## Group objects already on a slide

`groupObjects()` takes the `objectName` of each top-level object to group. It suits a slide built by separate functions, where replaying each object's descriptor just to group it is not an option.

```ts
slide.addShape("rect", { x: 1, y: 1, w: 2, h: 1, objectName: "Header" })
slide.addText("Caption", { x: 1.2, y: 2.2, w: 1.6, h: 0.6, objectName: "Caption" })
slide.addShape("rect", { x: 5, y: 1, w: 1, h: 1, objectName: "Loose" })

slide.groupObjects(["Header", "Caption"], { objectName: "Banner" })
// "Header" and "Caption" are now inside "Banner". "Loose" stays top-level.
```

| Object | Groupable |
| --- | --- |
| shapes, text boxes, images, connectors and groups | yes |
| charts, tables, media, 3D models, OLE objects and zooms | no, the call throws |
| any object in a layout placeholder | no, the call throws |

The call keeps the slide looking the same. Here `A` and `C` are grouped with `slide.groupObjects(["C", "A"])`, and each column lists the stack from the top down:

<svg role="img" aria-label="Stacking order before and after grouping A and C. Before, from the top: Over, C, B, A, Under. After, from the top: Over, then a group holding C above A, then B, then Under." viewBox="0 0 480 200" width="100%" style="max-width: 480px" fill="none" stroke="currentColor" font-size="12">
  <rect x="40" y="36" width="140" height="24"/>
  <rect x="40" y="66" width="140" height="24" fill="currentColor" fill-opacity="0.15"/>
  <rect x="40" y="96" width="140" height="24"/>
  <rect x="40" y="126" width="140" height="24" fill="currentColor" fill-opacity="0.15"/>
  <rect x="40" y="156" width="140" height="24"/>
  <rect x="280" y="36" width="140" height="24"/>
  <rect x="280" y="66" width="140" height="60" stroke-dasharray="4 3"/>
  <rect x="292" y="73" width="116" height="20" fill="currentColor" fill-opacity="0.15"/>
  <rect x="292" y="99" width="116" height="20" fill="currentColor" fill-opacity="0.15"/>
  <rect x="280" y="132" width="140" height="24"/>
  <rect x="280" y="162" width="140" height="24"/>
  <g fill="currentColor" stroke="none" text-anchor="middle">
    <text x="110" y="22">Before</text>
    <text x="110" y="52">Over</text>
    <text x="110" y="82">C</text>
    <text x="110" y="112">B</text>
    <text x="110" y="142">A</text>
    <text x="110" y="172">Under</text>
    <text x="350" y="22">After</text>
    <text x="350" y="52">Over</text>
    <text x="350" y="87">C</text>
    <text x="350" y="113">A</text>
    <text x="350" y="148">B</text>
    <text x="350" y="178">Under</text>
    <text x="448" y="100">group</text>
  </g>
</svg>

1. The members keep their stacking order among themselves. The order of the names in the array does not matter.
2. The group takes the place of the topmost member.
3. An object that sat between two members ends up above the group.
4. Every name is checked before anything moves, so a call that throws leaves the slide unchanged.

The frame and naming rules are the same as for `addGroup()`.

## List what a slide holds

`slide.objects` returns the objects on the slide, bottom of the stack first, with the names `groupObjects()` accepts. It is how code that did not add the objects finds out what to group.

```ts
slide.addShape("rect", { x: 1, y: 1, w: 3, h: 2, objectName: "card:frame" })
slide.addText("Revenue", { x: 1.2, y: 1.2, w: 2.6, h: 0.5, objectName: "card:title" })
slide.addText("Draft", { x: 5, y: 1, w: 1, h: 0.5 })

for (const o of slide.objects) {
  console.log(o.type, o.objectName, o.canGroup)
}
// text card:frame true
// text card:title true
// text Text 3 true

const card = slide.objects.filter((o) => o.canGroup && o.objectName.startsWith("card:"))
slide.groupObjects(card.map((o) => o.objectName), { objectName: "Card" })
```

Each entry is a `SlideObjectInfo`:

| Field | Meaning |
| --- | --- |
| `type` | the kind the object was added as. Shapes and text boxes both report `"text"` |
| `objectName` | the name you gave, spelled as you gave it, or the generated name PowerPoint shows, such as `Shape 3` |
| `isPlaceholder` | the object fills a layout placeholder, which `groupObjects()` refuses |
| `canGroup` | `groupObjects()` accepts this kind of object |
| `children` | a group's members, bottom first, nested to any depth. Empty for anything else |

- A name read from `slide.objects` resolves when you pass it back, including names that contain `&`, `<` or quotes.
- `canGroup` speaks for the object alone. A call can still throw for a name that matches nothing, or more than one object.
- Each read returns a fresh snapshot. Changing it does nothing to the slide.

## Target a shape inside a group

Group children are named on the slide like top-level objects, so a connector end or an animation can point at one.

```ts
slide.addGroup([{ rect: { x: 1, y: 1, w: 2, h: 1, objectName: "boxInGroup" } }], { objectName: "Grp" })
slide.addConnector({ type: "elbow", x1: 3, y1: 1.5, x2: 6, y2: 4.5, startShape: "boxInGroup", startShapeIdx: 3 })
slide.addAnimation({ preset: "fadeIn", objectName: "boxInGroup" })
```

[Bind a connector to a shape](connectors.md#bind-a-connector-to-a-shape) has the lookup order and what happens when a name does not resolve. An animation whose name resolves to nothing warns `animation/target-not-found` and is left out.

## Invalid input

`addGroup()` warns and carries on. `groupObjects()` throws instead, because a failed lookup would otherwise leave an object outside the group with no sign of it. Rows marked "when written" are reported while the deck is exported.

| Condition | Warns or throws | Code |
| --- | --- | --- |
| an `addGroup()` child is a chart, table, media or placeholder | warns, child skipped | `group/unsupported-child` |
| an `addGroup()` child has a key no descriptor uses | warns, child skipped | `group/unrecognized-child` |
| an `addGroup()` child needs a construct family the presentation was composed without | warns, child skipped | `family/child-not-composed` |
| `addGroup()` is left with no child to draw | warns, writes a zero-size group | `group/no-children` |
| one to three of `x`, `y`, `w`, `h` are set (when written) | warns, uses the bounding box | `group/partial-frame` |
| a lock flag groups do not support (when written) | warns, flag left out | `object-lock/unsupported-on-shape` |
| an `objectName` appears twice on the slide, group children included (when written) | warns | `object-name/duplicate` |
| `groupObjects()` gets an empty list, or no list | throws `InvalidOptionError` | `group/missing-object-names` |
| a name is not a non-empty string | throws `InvalidOptionError` | `group/invalid-object-name` |
| a name appears twice in the list | throws `InvalidOptionError` | `group/duplicate-object-name` |
| no top-level object has the name. The message says when it is already inside a group | throws `InvalidOptionError` | `group/unresolved-object-name` |
| more than one top-level object has the name | throws `InvalidOptionError` | `group/ambiguous-object-name` |
| the object's kind cannot be grouped, or it is a placeholder | throws `UnsupportedFeatureError` | `group/kind-not-groupable` |

Composing a presentation from construct families is covered in [Smaller bundles](bundle-size.md). [Errors and warnings](errors-and-warnings.md) covers the error classes and how to route warnings.

## Limits

- `addGroup()` takes only the descriptors in the table above.
- `groupObjects()` takes shapes, text boxes, images, connectors and groups only.
- `groupObjects()` reaches top-level objects only, so an object belongs to one group at a time.
- A written group never scales its children.
- There is no call to ungroup.
- A group takes no fill or outline of its own.

## Reading it back

- Before export, `slide.objects` reports groups and their `children`, as described above.
- After loading a deck with `pptx-ts/read`, a group is a `GroupShape` whose `shapes` are its children. [PPTX read API](reference/pptx-read.md) covers each child's slide-absolute geometry through the groups around it.
- [PPTX inspection](reference/pptx-inspection.md#geometry-groups-and-z-order) reports groups as `kind: 'group'` elements, and gives each child a slide-absolute box.

## See also

- [Connectors](connectors.md)
- [Positions and sizes](reference/layout-units.md#positions-and-sizes)
- [Text that fits](text-fit.md)
- [Demos](demos.md), whose showcase deck groups its content
- [Errors and warnings](errors-and-warnings.md)
- API reference: [`Slide`](reference/api/index/interfaces/Slide.md), [`GroupProps`](reference/api/index/interfaces/GroupProps.md), [`GroupChildProps`](reference/api/index/type-aliases/GroupChildProps.md), [`SlideObjectInfo`](reference/api/index/interfaces/SlideObjectInfo.md), [`ObjectLockProps`](reference/api/index/interfaces/ObjectLockProps.md)
