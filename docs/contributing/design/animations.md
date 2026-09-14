---
doc-schema-version: 1
title: "Animations and transitions: design"
summary: "How transitions and build animations are written, read and kept coherent: the p:transition target, the p:timing tree, why the two constructs are modelled differently, and the spid operations the import paths rely on."
read_when:
  - Changing how slide transitions are written or parsed
  - Changing build animation emission or the preset templates
  - Touching importShape or importSlide where animation spid references are affected
  - Adding a preset or a transition feature and its PowerPoint-authored fixture
doc_type: "architecture"
---

# Animations and transitions: design

The user guide is [Animations and transitions](../../animations-and-transitions.md). This page is
how the code behind it is shaped.

## Two constructs, modelled differently

A transition is self-contained. An animation refers to shapes by id, through `p:spTgt/@spid` and
`p:bldP/@spid`, and that coupling decides the design:

| | Reading a deck | Writing a deck |
| --- | --- | --- |
| Transition | typed model, get and set | typed model |
| Animation | preserved as it is, with spid-aware operations | fixed preset templates |

The transition schema is small and bounded, so a typed model is cheap. The `p:timing` tree
(`CT_TimeNodeList`) is deep and recursive, and two trivial effects already come to about ninety lines,
so there is no semantic model of it. The read side keeps the DOM and tracks only the `spid`
references; the write side emits templates captured from PowerPoint.

## The transition target

`p:transition` sits in `p:sld` after `p:clrMapOvr` and before `p:timing`.
`src/gen/slide/slide.ts` writes the children in that order, calling `slideTransitionToXml`
(`src/gen/anim/transition.ts`) between them.

| Part of the target | What is written |
| --- | --- |
| `spd` | the `speed` given, otherwise the bucket for `durationMs` from `src/ooxml/transition-speed.ts` (up to 500 ms `fast`, up to 1000 ms `med`, else `slow`), otherwise nothing |
| `advClick` | `0` only when `advanceOnClick` is `false` |
| `advTm` | `advanceAfterMs`, rounded |
| type element | one of the 21 base types in `TRANSITION_TYPES` (`src/ooxml/st-enums.ts`); an unknown type warns and no transition is written |
| type attributes | `variant` keys checked against `TRANSITION_VARIANT_ATTRIBUTES`; values are not checked |
| `p:sndAc` | `p:stSnd`/`p:snd` for an embedded start sound, `p:endSnd` for stop-previous |
| exact duration | when `durationMs` is set, an `mc:AlternateContent` whose `p14` Choice carries `p14:dur` and whose Fallback is the same transition without it |

The read side (`src/read/api/transition.ts`) accepts both the bare `p:transition` and the
`mc:AlternateContent` form, preferring the Choice so `durationMs` survives, and reports a transition in
the `p14`, `p15` or `p159` namespace through `namespace`. Its setter validates times
(`transition/invalid-time`) and refuses a sound it would have to embed (`transition/sound-unsupported`).

### Transition sounds

A start sound needs an audio media part and a relationship. `registerTransitionSounds`
(`src/package/assemble.ts`) registers them while the package is assembled, not when `slide.transition`
is assigned: a slide hands out relationship ids and media part names in registration order, and a
transition is often set before the slide's pictures, so registering early would renumber everything
added after it. A slide re-exported with the same sound reuses its relationship; `stopPrevious` needs no
part; a transition the emitter will not write registers nothing. The part's extension comes from the
data URI's media type through `audioExtensionForSubtype`, because PowerPoint writes `audio/x-wav`, which
is not itself an extension.

## The timing tree

```mermaid
flowchart TD
  timing["p:timing"] --> tnLst["p:tnLst / p:par"]
  tnLst --> root["p:cTn nodeType=tmRoot"]
  root --> seq["p:seq / p:cTn nodeType=mainSeq"]
  seq --> click["p:par / p:cTn: a click step, delay indefinite"]
  click --> sub["p:par / p:cTn: a sub-step, with its delay"]
  sub --> effect["p:par / p:cTn presetID, presetClass, presetSubtype, nodeType"]
  effect --> behaviour["behaviours targeting p:spTgt spid"]
  timing --> bldLst["p:bldLst / p:bldP spid grpId"]
```

`buildAnimationSeq` (`src/gen/anim/animation.ts`) groups effects the way PowerPoint does: `onClick`
opens a click step, `afterPrevious` opens a sub-step delayed by the previous effect's duration, and
`withPrevious` joins the current sub-step. `buildBldList` writes one `p:bldP` per animated shape.
Each preset in `ANIM_PRESETS` is a template keyed by preset name, carrying its `presetID`, class,
subtype and default duration, and parameterized only by `spid` and duration. Adding a preset means
authoring a PowerPoint fixture and capturing its template, not a new code path.

`addAnimation` (`src/families/animations.ts`) only records the options; targets resolve at write time
by `objectName` or by `shapeIndex` among the slide's drawn top-level objects.

## spid-aware operations

| Operation | What it does | Where it runs |
| --- | --- | --- |
| enumerate | collects every `spid` from `p:spTgt` and `p:bldP` | `Slide.animationSpids()` (internal) |
| remap | rewrites `spid`s through an old-to-new map | `carryShapeAnimations`; `Slide.remapAnimationSpids` (internal) |
| prune | removes a shape's `p:bldP` and the effects targeting it, then empty wrappers | `carryShapeAnimations`; `Slide.pruneAnimationSpids` (internal) |
| flatten | removes the whole `p:timing` when it holds build animations, keeping a media-only timing | `Slide.flattenAnimations()` |
| carry | copies a lifted shape's click steps and `p:bldP` into the destination timing, creating it if absent, pruning effects on shapes not carried, remapping ids and renumbering `p:cTn` ids past the destination's highest | `importShape`/`importShapes` with `carryAnimation` (`src/read/api/presentation-imports.ts`) |

`hasAnimations` is true for a `p:bldP` or a `presetID`-bearing `p:cTn`, which is what separates build
animation timing from the timing a looping media object needs.

When `importSlide` carries master or layout decorations onto a slide, `src/read/api/ops/import-slide.ts`
renumbers them past every id already on the slide, so a slide animation's `spid` never ends up naming
two shapes.

## Fixtures

Every emitter and parser here was built against PowerPoint-authored decks, recorded with their oracles
in `test/read/fixtures/README.md`. `test/regression/slide-content/animations-transitions.test.js`
compares the emitted transition forms and timing trees with those oracles byte for byte, and the
`p:sndAc` with relationship ids normalized, since the ids are this library's own numbering.

| Fixture | What it pins |
| --- | --- |
| `slide-transition` | transition types, speed buckets, `p14:dur`, click and timed advance, and the effect table the write side uses |
| `slide-transition-sound` | embedded, looped and stop-previous sounds and their media relationship |
| `slide-animation-basic` | one entrance effect: the `p:timing` tree, `p:bldLst` and its `spid` |
| `slide-animation-rich` | entrance, emphasis and exit across all three triggers; the source of spid enumerate, remap and prune tests |
| `slide-animation-presets` | the templates for every preset |
| `import-animation-merge` | how PowerPoint merges a pasted shape's animation into the destination slide |
