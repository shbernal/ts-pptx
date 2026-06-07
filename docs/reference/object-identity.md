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

Images and charts also expose `altText`. For those object types, explicit
`altText` is serialized to the generated object's `p:cNvPr` `descr` attribute.
Other object types do not currently have a public alt-text contract.
