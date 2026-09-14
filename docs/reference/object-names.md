---
doc-schema-version: 1
title: "Object names and alt text"
summary: "The Selection Pane name and alt text of every slide object: the default objectName of each object kind and how its number counts, the default descr, where names are used, and the warnings a name raises."
read_when:
  - Finding an object by its Selection Pane name, in slide.objects or in a deck read back
  - Working out why a default name skips a number, such as Shape 2 after Text 1
  - Checking what alt text an object carries when altText is omitted
  - Debugging an object-name/empty, object-name/too-long, object-name/control-characters or object-name/duplicate warning
doc_type: "reference"
---

# Object names and alt text

Every object ts-pptx places on a slide has a name and alt text. They are written to the `name` and `descr` attributes of the object's `p:cNvPr` element. PowerPoint shows the name in the Selection Pane and the alt text under Edit Alt Text.

Set them with the `objectName` and `altText` options. Zooms take `objectName` only.

```ts
import TsPptx from "pptx-ts"

const pptx = new TsPptx()
const slide = pptx.addSlide()
slide.addText("Revenue", { x: 1, y: 1, w: 3, h: 0.5 })
slide.addShape("rect", { x: 1, y: 2, w: 3, h: 1, objectName: "card" })
slide.addShape("rect", { x: 5, y: 2, w: 3, h: 1, altText: "Empty card" })

for (const o of slide.objects) console.log(o.type, o.objectName)
// text Text 1
// text card
// text Shape 3
```

The text box and the shapes count together, and `card` used number 2, so the last shape is `Shape 3`.

## Default names

An object without `objectName` gets its kind's label and a number. Numbers start at 1. Each slide counts on its own, and so does each layout `defineSlideMaster()` makes.

| Object | Default `objectName` | How the number counts | `descr` without `altText` |
| --- | --- | --- | --- |
| `addText()` | `Text N` | together with `addShape()` | empty |
| `addShape()` | `Shape N` | together with `addText()` | empty |
| `addImage()` | `Image N` | images | the `path` as you passed it, or `preencoded.png` for an image given as `data` |
| `addChart()` | `Chart N` | charts, classic and 2016 types together | empty |
| `addTable()` | `Table N` | tables | empty |
| `addConnector()` | `Connector N` | connectors | empty |
| `addMedia()` | `Media N` | audio, video and online video together | empty |
| `addOleObject()` | `Object N` | OLE objects | empty |
| `addModel3d()` | `3D Model N` | 3D models | empty |
| `addGroup()`, `groupObjects()` | `Group N` | groups. A nested group takes its number before the group around it | empty |
| `addSlideZoom()`, `addSectionZoom()`, `addSummaryZoom()` | `Slide Zoom N`, `Section Zoom N`, `Summary Zoom N` | the three zoom kinds together | always empty |
| a placeholder in `defineSlideMaster()` | its `name` | no number | empty |
| `addText()` with a `placeholder` that names a layout placeholder | the layout placeholder's name | uses up a number from the `Text` count | the layout placeholder's `altText`, or empty |
| a layout placeholder the slide leaves empty | the layout placeholder's name | no number | the layout placeholder's `altText`, or empty |
| the slide number | `Slide Number Placeholder 0` | fixed | no `descr` attribute |

- Every object takes a number when it is added, including an object with its own `objectName` and an object inside a group.
- `addText()` with a `placeholder` that matches no layout placeholder gets `Text N`.
- Give images `altText`. Without it the file path, or `preencoded.png`, becomes the text a screen reader announces.

## Where names are used

- `slide.objects` lists each object's `objectName`, bottom of the stack first, with a group's members in `children`. It reports the name as you spelled it, or the default name.
- `addConnector()` attaches an end to the object named in `startShape` or `endShape`. See [Bind a connector to a shape](../connectors.md#bind-a-connector-to-a-shape).
- `groupObjects()` and `addAnimation({ objectName })` find their objects by name. See [Group objects already on a slide](../groups.md#group-objects-already-on-a-slide).
- `pptx-ts/read` reports the written name as `shape.name`.

## Name validation

ts-pptx stores `objectName` as you wrote it and XML-encodes it once, when the deck is written. A name you supply can raise these warnings. None of them throws.

- `object-name/empty`: the name is only whitespace. An empty string is not checked: `objectName: ""` gets the default name.
- `object-name/control-characters`: the name holds a control character, U+FFFE or U+FFFF. The file carries the name with those characters removed, while `slide.objects` keeps them.
- `object-name/too-long`: the name is longer than 255 characters. PowerPoint may not keep it.
- `object-name/duplicate`: two objects on one slide or layout share a name, group members included. It is raised each time the deck is written.
