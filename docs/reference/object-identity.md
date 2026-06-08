---
doc-schema-version: 1
title: "Object Identity"
summary: "Public objectName and altText contracts for generated PPTX objects."
read_when:
  - Using stable Selection Pane names
  - Testing generated object names or alt text
  - Updating objectName or altText serialization
doc_type: "reference"
---

# Object Identity

PptxGenJS exposes `objectName` on generated slide objects that have a PowerPoint
Selection Pane identity. When set, `objectName` is serialized to the generated
object's `p:cNvPr` `name` attribute.

The explicit `objectName` contract applies to:

- text boxes and text-backed shapes;
- shapes;
- images, including SVG-backed images;
- charts;
- tables;
- media objects;
- slide master placeholders.

When `objectName` is omitted, PptxGenJS emits an internal default such as
`Text 0`, `Shape 0`, `Image 0`, `Chart 0`, `Table 0`, or `Media 0`. Consumers
that need stable semantic identity should set `objectName` explicitly instead
of depending on generated default names.

## Alt text

`altText` is a universal contract across every object kind listed above: text
boxes and text-backed shapes, shapes, images (including SVG-backed images),
charts, tables, and media. When set, `altText` is serialized to the generated
object's `p:cNvPr` `descr` attribute. Images additionally fall back to the image
filename for `descr` when `altText` is omitted; all other kinds emit an empty
`descr` when `altText` is omitted.

## Name validation

`objectName` values are XML-encoded before serialization. PptxGenJS also warns
(via `console.warn`, without throwing) at generation time when a name cannot
provide a stable Selection Pane identity, so identity bugs surface early instead
of producing a silently broken deck:

- empty or whitespace-only names;
- names containing control characters (these are stripped during XML encoding,
  silently changing the stored name);
- names longer than 255 characters (may not be preserved by PowerPoint);
- duplicate `objectName` values emitted on a single slide (consumers that rely
  on unique names, such as semantic manifests, cannot disambiguate them).

These are warnings, not errors: existing decks with loose names continue to
build unchanged.
