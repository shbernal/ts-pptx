---
doc-schema-version: 1
title: "Animations & Transitions"
summary: "Scope and design for PptxGenJS slide-transition and per-shape build-animation support: a full typed model for transitions (read + write), opaque spid-aware preservation for the p:timing animation tree on read, and preset-template authoring on write. Not yet implemented — this is the foundation decision record and the fixture gate."
read_when:
  - Implementing or changing slide transition emit/parse (p:transition)
  - Implementing or changing animation handling (p:timing / p:bldLst)
  - Touching importShape / importSlide where slide-scoped timing or spid references are affected
  - Authoring or interpreting the animation/transition fixtures + oracles
doc_type: "decision"
---

# Animations & Transitions

## Status

**Scoping locked; fixtures authored; not yet implemented.** This doc records the
agreed read/write scope so the fixtures + oracles could be authored against a
fixed target. Per `AGENTS.md` ("OOXML And PowerPoint Work" → fixture-gated work),
implementation was **blocked on the PowerPoint-authored fixtures + oracles**
enumerated in §7; those are the precondition, not synthetic XML.

Authorability was proven 2026-06-26: desktop PowerPoint COM bakes both
`p:transition` and `p:timing` into slide XML non-interactively (probe
`.tmp/author-anim-probe.ps1`). See the memory
`animations-transitions-fixture-authorability`.

**The §7 fixture gate is now satisfied (2026-06-26).** All three decks + oracles
are authored, verified clean via COM, and documented in
`test/read/fixtures/README.md`; the backlog entry `gitbrent/PptxGenJS#1431`
records the satisfaction. The implementation is unblocked but remains ON STAND-BY
until scheduled:

- `slide-transition.pptx` / `.oracle.json` — fade/push(dir)/wipe(dir)/cut/dissolve
  across fast/med/slow buckets, bare vs `mc:AlternateContent` (`p14:dur`) forms,
  `advClick` vs `advTm`; the oracle embeds the full **158-row probed
  `PpEntryEffect → {element, ns, variant, modernOnly}` table** (the write-side
  preset table — 159 valid ints incl. `0`; e.g. `1537 → dissolve`, and 95 modern
  `p14`/`p15`/`p159`-only effects with a `<p:fade/>` fallback).
- `slide-animation-basic.pptx` / `.oracle.json` — one entrance Fade-on-click;
  pins the `p:timing`/`p:bldLst` tree, the `(presetID 10, entr, subtype 0)`
  triple, and `spid 2`.
- `slide-animation-rich.pptx` / `.oracle.json` — four shapes spanning
  `entr`/`emph`/`exit` × `clickEffect`/`afterEffect`/`withEffect` (emphasis
  Grow/Shrink `presetID 6` via probed `MsoAnimEffect=59`); the source for the
  write-side preset templates and the spid enumerate/remap/prune tests.

## Decisions (locked)

Two constructs, two subsystems. The constructs differ in one decisive way:
**a transition is self-contained; an animation references shapes by `spid`**
(`<p:spTgt spid="N"/>` and `<p:bldP spid="N"/>`). That coupling is the crux of
the whole design.

|              | Read / round-trip                         | Write / authoring                          |
| ------------ | ----------------------------------------- | ------------------------------------------ |
| **Transition** | Full typed model (`slide.transition`)   | Full typed model (slide `transition` prop) |
| **Animation**  | **Opaque, spid-aware** preservation     | **Preset-template** effects (fixed set)    |

- **Direction:** both *preserve* (read→write keeps these intact, including across
  the import paths) and *author* (a from-scratch API adds them to generated
  slides).
- **Transitions get a full typed model** both ways — the schema is small and
  bounded (§3), so semantic modeling is cheap and high-value.
- **Animations are modeled opaquely.** The `p:timing` tree is `CT_TimeNodeList`
  — dozens of node types, deeply recursive; even two trivial effects produce
  ~90 lines (§4). We do **not** build a semantic AST. On read we preserve the
  DOM and track only the `spid` references; on write we emit known-good
  templates captured from PowerPoint for a fixed preset set. Authoring without a
  semantic builder is *only* possible via preset templates — that is the agreed
  authoring mechanism, consistent with the opaque internal model.

## Architecture context (why the two sides differ)

- **Read (`src/read/`)** is a lazy-parse, DOM-preserving model. `Part` keeps its
  original bytes and `serialize()` returns them byte-identically unless
  `markDirty()` was called (`src/read/opc/part.ts`). Typed accessors parse on
  demand over the live DOM; edits go through `getOrAddChild`/`setAttr` +
  `markDirty()` (pattern: `Shape.resolvedFill`, `TextFrame.text`). **Consequence:
  an unmodified slide already round-trips its `p:transition`/`p:timing`
  byte-identically for free.** The work is (a) a typed transition accessor and
  (b) keeping `spid` references coherent when ids change.
- **Write (`src/gen-xml.ts`)** is string concatenation from the object model.
  `makeXmlSlide` emits `spTree → clrMapOvr → timing`; a `p:timing` builder
  already exists for media looping (`slideTimingToXml`, ~L2310) and is the exact
  structural template for animation emit. Shape ids are deterministic
  (`idx + 2`), so `spid` targeting from the authoring API is tractable.

## OOXML target

### Transition (`CT_SlideTransition`, ECMA-376 + p14 extension)

Position in `CT_Slide`: `cSld → clrMapOvr → **transition** → timing → extLst`.
So `p:transition` is emitted **between `p:clrMapOvr` and `p:timing`** — today
`makeXmlSlide` puts nothing there.

Base ECMA-376 `p:transition`:

- Attributes: `spd` (`ST_TransitionSpeed` = `slow`|`med`|`fast`, default `fast`),
  `advClick` (bool, default `true`), `advTm` (`unsignedInt`, ms — advance after
  time).
- Child: exactly one transition-type element (choice of 21: `blinds`, `checker`,
  `circle`, `dissolve`, `comb`, `cover`, `cut`, `diamond`, `fade`, `newsflash`,
  `plus`, `pull`, `push`, `random`, `randomBar`, `split`, `strips`, `wedge`,
  `wheel`, `zoom`), then optional `sndAc` (sound), then `extLst`. Each type
  element carries its own variant attrs (e.g. `<p:push dir="u"/>`,
  `<p:wipe dir="d"/>`).

**Exact duration is not in base ECMA** — `spd` is only a coarse bucket. PowerPoint
2010+ writes the precise milliseconds as `p14:dur` inside an
`mc:AlternateContent` wrapper, and emits a base-only `mc:Fallback`. Probe output:

```xml
<mc:AlternateContent xmlns:mc="…/markup-compatibility/2006">
  <mc:Choice xmlns:p14="…/office/powerpoint/2010/main" Requires="p14">
    <p:transition spd="slow" p14:dur="1500"><p:dissolve/></p:transition>
  </mc:Choice>
  <mc:Fallback>
    <p:transition spd="slow"><p:dissolve/></p:transition>
  </mc:Fallback>
</mc:AlternateContent>
```

Newer transition types (Morph, etc.) live entirely in `p14`/`p15` and only appear
inside `mc:Choice`. **Write target: match PowerPoint** — emit the
`mc:AlternateContent` form when a precise duration is requested (Choice with
`p14:dur` + base Fallback); a plain `p:transition` is acceptable when only a
speed bucket is given. **Read target:** handle both the bare `p:transition` and
the `mc:AlternateContent`-wrapped form, preferring the `p14` Choice for duration.

`spd` ↔ duration mapping and the `PpEntryEffect`-int ↔ transition-element mapping
are PowerPoint specifics, not in the XSD. There are 159 valid `PpEntryEffect`
ints (probed); the `family<<8` intuition is wrong (`1537` → `dissolve`, not
fade). **A required build artifact is a probed `PpEntryEffect → {element, variant
attrs, p14 variant?}` table** (author once by iterating accepted ints and
dumping XML); this becomes the write-side preset table and is captured in the
transition fixture's oracle.

### Animation (`p:timing` / `p:bldLst`) — opaque

Position: last child before `extLst`. Structure (from the probe):
`timing → tnLst → par → cTn(nodeType="tmRoot") → childTnLst → seq(mainSeq) → …`
with one nested `par/cTn` per effect carrying `presetID` / `presetClass` (e.g.
`entr`) / `presetSubtype` / `nodeType` (`clickEffect`|`afterEffect`|`withEffect`)
and `cBhvr/tgtEl/spTgt @spid`. A sibling `<p:bldLst>` holds one `<p:bldP
spid= grpId=>` per animated shape.

We treat this subtree as **opaque XML**. The only structured data we extract is
the set of referenced `spid`s (the "spid-aware" contract, §5). The known
preset effects we *author* (write side) are stored as verbatim templates keyed
by `(presetID, presetClass, presetSubtype, nodeType)`, parameterized only by
`spid`, `delay`, and `dur`.

## Read model design

1. **`slide.transition`** — typed accessor (`src/read/api/slide.ts`):
   - getter: parse `p:transition` (bare or `mc:AlternateContent`) → `{ type,
     speed, durationMs?, advanceOnClick, advanceAfterMs?, variant? }` or `null`.
   - setter: write/replace via DOM helpers + `markDirty()`, emitting the
     `mc:AlternateContent` form when `durationMs` is set, inserted at the
     schema-correct slot (before `p:timing`).
2. **Animation preservation** — no semantic accessor. Unmodified slides keep the
   `p:timing`/`p:bldLst` DOM untouched (free byte-identical round-trip). Add an
   internal **spid index** helper that, given the slide DOM, enumerates every
   `p:spTgt/@spid` and `p:bldP/@spid`. Expose minimally: `slide.hasAnimations`
   (bool) and an internal `animationSpids()` for the import paths. The raw tree
   is never reconstructed from a model — it is carried as-is or pruned in place.

## Write model design

1. **Transition** — slide-level `transition?: TransitionProps` on `PresSlide`
   (`src/core-interfaces.ts`); a `slideTransitionToXml(slide)` emitter
   (`src/gen-xml.ts`) inserted in `makeXmlSlide` **between `p:clrMapOvr` and
   `slideTimingToXml(...)`**. Requires declaring `xmlns:mc` on the slide root (or
   locally) for the AlternateContent form. Reuse the probed preset table for
   element + variant.
2. **Animation (preset templates)** — authoring API targets shapes by their
   generated id (`idx + 2`). For each requested effect emit the verbatim
   template (entrance/exit/emphasis from the fixed set) with `spid`/`delay`/`dur`
   filled in, assembling them under one `mainSeq` and appending matching
   `<p:bldP>` entries. This **extends `slideTimingToXml`** so a slide can carry
   *both* looping media and build animations in the single allowed `p:timing`.
   The supported preset set == the set we author fixtures + oracles for; adding a
   preset later means adding a fixture + template, not a new code path.

## spid-awareness (the crux)

Animations dangle if a referenced `spid` disappears or is renumbered — PowerPoint
then shows a repair prompt. The opaque model still must keep references coherent.
Three operations, all purely structural (no semantic parse):

- **enumerate** — collect `spid`s from `p:spTgt` and `p:bldP` (read side helper).
- **remap** — given an `oldSpid → newSpid` map, rewrite all matching `@spid`.
  Needed by any id-reassigning op.
- **prune** — when a shape is removed, drop its `p:bldP` and the effect nodes
  whose `spTgt` targets it, so no dangling reference survives.

### Interaction with import paths (`src/read/api/presentation.ts`)

- **`importSlide` copy mode** — whole slide part copied byte-identically;
  transition + timing already survive. No change needed.
- **`importSlide` preserve/restyle modes** — rebind to destination master; shape
  ids are not generally renumbered, so timing usually survives, but this must be
  **verified by fixture** and remap applied if ids do change.
- **`importShape`** — lifts a single `p:sp` and **reassigns its id**; slide-scoped
  timing is not part of the shape subtree (documented limitation at
  `presentation.ts:~1372`). Phase 2 option: carry that shape's `p:bldP` + effect
  nodes into the destination timing with a remap. Phase 1: keep the current
  drop, but ensure we never leave a dangling reference on either side.

## Public API surface (sketch — to be finalized in implementation)

```ts
// Read (src/read): typed transition; animation stays opaque.
slide.transition // -> TransitionInfo | null   (get/set)
slide.hasAnimations // -> boolean

// Write (core): slide-level transition + preset animations.
interface TransitionProps {
  type: TransitionType        // 'fade' | 'push' | 'wipe' | 'cut' | 'dissolve' | …
  durationMs?: number         // exact ms (emits p14:dur + AlternateContent)
  speed?: 'slow' | 'med' | 'fast'
  advanceOnClick?: boolean    // default true
  advanceAfterMs?: number     // advTm (auto-advance)
  variant?: Record<string, string> // e.g. { dir: 'u' } for push/wipe
}
pptx.addSlide({ /* … */ }).transition = { type: 'fade', durationMs: 1500 }

// Preset build animations target shapes added to the slide.
interface AnimationProps {
  preset: PresetEffect        // fixed set, e.g. 'fadeIn' | 'flyIn' | 'appear' | 'fadeOut'
  trigger?: 'onClick' | 'withPrevious' | 'afterPrevious'
  delayMs?: number
  durationMs?: number
}
```

(Per-shape vs slide-level animation list — and the exact preset enum — are
finalized during implementation against the authored fixtures.)

## Fixtures + oracles (the gate)

Author with desktop PowerPoint COM (skill `powerpoint-fixture-authoring`; clear
`Resiliency\DocumentRecovery` first). Each `.pptx` pairs with an oracle JSON
(verbatim XML + extracted fields), like `embedded-fonts.oracle.json`.

1. **`slide-transition.pptx`** — several slides, distinct transitions (at least
   `fade`, `push` (with `dir`), `wipe`, `cut`, `dissolve`) covering: speed
   buckets, exact `p14:dur`, advance-on-click vs `advTm` timed. Oracle pins each
   slide's full transition XML (Choice + Fallback) and the decoded fields, **plus
   the probed `PpEntryEffect → element` table** used by the write side.
2. **`slide-animation-basic.pptx`** — one slide, single entrance (fade-on-click).
   Oracle pins the `p:timing` tree + `p:bldLst` + the referenced `spid` and the
   `(presetID, presetClass, presetSubtype)` triple.
3. **`slide-animation-rich.pptx`** — multiple shapes/effects/triggers (click +
   after-previous + with-previous; an exit and an emphasis). Exercises spid
   enumerate/remap/prune and is the source of the write-side preset templates.

Update `test/read/fixtures/README.md` (provenance, hash, purpose, PowerPoint
check date). Record the fixtures as the blocking precondition in
`docs/backlog.yml` before implementing.

## Out of scope / phasing

- **No semantic animation AST.** No general `p:timing` builder, no editing of
  arbitrary effects, no animation paths/triggers beyond the preset set.
- **Phase 1:** transitions (full, both ways); animation opaque preserve + spid
  enumerate/remap/prune; preset-template authoring for the fixtured effect set.
- **Phase 2 (candidate):** carry a shape's build animation through `importShape`
  (remap into destination timing); expand the preset set; transition sounds
  (`sndAc`, which pulls in audio rels).
- **Open questions:** (a) preset animations as a per-shape option vs a
  slide-level ordered list (ordering/sequence semantics favor a slide-level
  list); (b) whether to always emit the `mc:AlternateContent` form or only when
  `durationMs` is set; (c) exact `spd`↔`durationMs` bucketing when both are
  given.
```
