# Grouping — audit findings & remediation plan

Status: **Phase 1 landed (D1, D2, D3); Phase 2 landed (D5, D6); Phase 3 landed
(D4); Phase 4 landed (§3 name resolution); Phases 5–6 outstanding.** Every defect
below was verified by generating or re-reading a package, not by reading source
alone; the measured evidence is quoted inline so each item can be re-checked
independently.

Phase 1 notes: the mapping was extracted to `src/group-transform.ts`
(`composeGroupFrame`) and is now the single implementation behind both
`Shape.absoluteFrame` and `inspect()` — the "second implementation" risk in §2/D1
is closed rather than duplicated. `inspect()` also gained `rotation`/`flipH`/`flipV`,
which D1 needs to report a rotated group child honestly. D3 required switching
`inspect.ts` to a `preserveOrder` parse: fast-xml-parser's default tag-keyed output
cannot express sibling order across different tag names, so document order was not
recoverable without it.

Phase 2 notes: both defects were one model, so they were replaced by one — a
monotonic per-slide, per-kind counter (`slide._objectNameCounts`, taken via
`nextObjectNameIdx`) instead of counting `_slideObjects` at add time. Counting
was what the splice defeated (D5); a module-global was what leaked across
presentations (D6). An index is consumed by every object of a kind, explicitly
named or not, which is what keeps the pre-existing non-group numbering intact
(524 unit tests unchanged). §4's `groupObjects()` is unblocked. Not folded in:
`_chartCounter` — it is presentation-scoped (chart part filenames), and slides
hold no back-reference to their presentation, so it needs threading rather than
the same one-line move; it remains a real "same input, different bytes" defect
and is recorded below.

Phase 3 notes: **(b) was chosen** — a partial frame warns and falls back to
auto-bounds on every axis. (a) per-axis fallback was rejected on the grounds
recorded under D4: with an identity child space the frame never moves or scales
the children, so `{ x: 5 }` would read like a reposition, leave the children
where they were, and put the group's box somewhere they are not. Falling back
whole keeps the box around its content and says so out loud. The predicate is
shared (`hasCompleteGroupFrame`, `src/gen-xml.ts`) between `resolveObjBounds` and
the group renderer, so a partial frame resolves identically where a parent group
sizes around it; the warning is emitted at the render site only, which runs once
per group object, so a nested partial frame warns exactly once.

Phase 4 notes: both drops were one lookup (`_slideObjects.findIndex` by
`objectName`), so both were replaced by one — `collectSlideShapeIds` (`src/gen-xml.ts`)
maps every object on the slide, group children included, to the `<p:cNvPr>` id it is
rendered with, and `resolveObjectNameToId` searches it. Top-level objects are inserted
first, so they still win a duplicate name and no existing deck's bindings move. The map
**mirrors** the id allocation the render walk performs (it must: a reference can be
emitted before the walk reaches its target), so the new tests parse the emitted `cNvPr`
ids back out and assert each reference points at the shape it names — a hardcoded id
would let the two drift while still passing. PowerPoint (COM) confirms it resolves the
references, not merely that the package opens: a connector bound to a grouped shape at
both ends reports `begin=GroupedBox end=DeepBox`, and all three animations land on their
named shapes. Not folded in: connectors as group *children* (§3) — that is an
`addGroup` scope question (a `<p:cxnSp>` child kind), not a name-resolution one.

Scope: `<p:grpSp>` across the three surfaces that touch it — the **write** path
(`addGroup`), the **read** path (`Presentation` / `Shape.absoluteFrame`), and the
**inspect/measure** path (`inspectPptx`, `applyMeasuredFit`).

Contains no downstream identifiers; safe to commit.

---

## 1. Summary

The three surfaces are at very different maturity levels, and that asymmetry is
the story:

| Surface | State | Assessment |
|---|---|---|
| Read (`src/read/api/shapes.ts`, `presentation.ts`) | Composes nested scale + rot + flips; guards degenerate `chExt`; PowerPoint-verified fixture | **Solid.** Reference implementation |
| Write (`addGroup`) | Documented MVP: identity child space, 4 child kinds rejected | **Intentionally narrow**, with sharp edges |
| Inspect (`src/inspect.ts`) | No group awareness at all | **Silently wrong** on real decks |

The single most important fact: **`inspect.ts` is correct today only by
coincidence.** It reads each shape's raw `a:xfrm` and never applies the group
mapping. That happens to be right for packages this library authored, because the
writer hardcodes an identity child space (`chOff/chExt == off/ext`). It is wrong
for any deck PowerPoint has touched — a user resizing a group is enough to make
`chExt` non-identity. A correct implementation already exists ~200 lines away in
`Shape.absoluteFrame` (`src/read/api/shapes.ts:615`) and is not reused.

`measure-fit.ts` shares that latent dependency (`computeBox`, `:182-196`, reads
`opts.w/h` off the child with no ancestor walk) but is **not** currently wrong:
`addGroup` cannot author a scaled group, and measured-fit only runs over objects
this library built. It needs a comment tying it to the invariant, not a fix.

---

## 2. Confirmed defects

### D1 — `inspect()` reports group children in child space (P1)

`src/inspect.ts` contains zero occurrences of `grpSp`, `chOff`, or `chExt`.
`normalizeElement` (`:280-282`) emits each shape's own `a:xfrm` as `box` with no
ancestor traversal, so the `off + (p − chOff) × (ext / chExt)` mapping and group
`@rot`/`@flipH`/`@flipV` are never applied.

Measured against `test/read/fixtures/group-transform.pptx` (PowerPoint-authored),
child `scale-rot child blue rect`:

| Source | x | y | w | h | rot |
|---|---|---|---|---|---|
| `inspect()` `box` (EMU) | 889000 | 4381500 | **1219200** | 609600 | — |
| `absoluteFrame` (truth) | 1191370 | 4121483 | **1609344** | 499872 | 25° |

Width is off by ~32% (the group's scale ratio); position is wrong; rotation is
absent. Silent — no warning.

**Fix:** have `inspect.ts` compose enclosing group transforms. Prefer delegating
to the existing `absoluteFrame` logic over a second implementation; if the
inspect surface must stay DOM-proxy-free, extract the mapping into a shared
helper and use it from both. Note `absoluteFrame` returns `null` for degenerate
`chExt` — inspect needs a defined answer for that case (omit + warn).

### D2 — `inspect()` never reports the group itself (P1, same file)

`'p:grpSp'` is absent from the key list at `src/inspect.ts:269`, so groups are
harvested only as a side effect of the generic walker (`:442-451`) recursing into
every object value. Consequences: the group's `cNvPr` id/name and `grpSpPr` fill
are dropped, `PptxSlideElementKind` (`:62`) has no `'group'` member, and
`PptxSlideElement` (`:87-111`) has no parent/child field — the returned list is
flat with no indication which elements were grouped.

**Fix:** add a `'group'` kind and a parent/children relationship. This is an API
shape change to the inspect model; per the API Evolution Policy that is
acceptable, but it must land in `CHANGELOG.md` with migration guidance.

### D3 — `inspect()` `zIndex` is not document order (P2, same file)

`.map((el, zIndex) => …)` at `src/inspect.ts:224` assigns z from `collectElements`
order, which is "all `p:sp` of a node, then all `p:pic`, then all `p:cxnSp`, then
recurse". Verified on the same fixture: every `text` element is listed before
every `shape`, and grouped children always sort after all top-level shapes
regardless of true z. This is wrong for mixed-type slides **even without groups**,
so it is strictly broader than grouping — but grouping makes it worse and it
should be fixed alongside D1/D2 while the file is open.

**Fix:** walk children in document order rather than per-key.

### D4 — Partial group options emit a degenerate group (P1) — **FIXED** (option (b))

`hasExplicit` (`src/gen-xml.ts:1243-1247`) is an **OR** over `x/y/w/h`. When any
one is set, the unset axes fall back to the shared defaults (`src/gen-xml.ts:534`:
`x=0`, `y=0`, `cx=75% of layout width`, **`cy=0`**) instead of the child bbox.

Verified — `addGroup([rect], { x: 5, y: 2 })` emits:

```xml
<a:xfrm><a:off x="4572000" y="1828800"/><a:ext cx="6858000" cy="0"/>
        <a:chOff x="4572000" y="1828800"/><a:chExt cx="6858000" cy="0"/></a:xfrm>
```

A zero-height group whose `cx` is silently a slide-width fraction. On re-read
every child of it resolves to `null` via the degenerate-`chExt` guard
(`src/read/api/shapes.ts:639`).

This is precisely the footgun AGENTS.md names: *"prefer warning or failing on
`NaN` / `undefined` / out-of-range values over emitting a degenerate result (e.g.
a zero-size object)."*

**Fix (decide one, then document it):**
- (a) per-axis fallback — each unset axis takes the child bbox value; or
- (b) all-or-nothing — a partial frame `warn()`s and falls back to full auto-bounds.

(a) is more useful (`{ x }` to reposition a group reads naturally) but note it
does **not** move children: `chOff` tracks `off`, so the mapping stays identity
and only the group handle and rotate pivot move. If that is surprising, (b) is
the honest choice. Either way the current behaviour is indefensible.

**Decided: (b).** Precisely because `{ x }` reads as a reposition and is not one —
(a) would satisfy the reading silently and put the group's box where its children
are not, which is the same class of footgun as the zero-extent group it replaces.
A complete frame (all four axes) is still honoured verbatim, unchanged.

### D5 — Default object names collide across the group boundary (P1) — **FIXED**

Default names count the current `_slideObjects` (e.g. `src/gen-objects.ts:1170`,
`:705`), but `buildGroupObject` **splices children back out** of that array
(`:168`), so the count never advances.

Verified — `addGroup([{ rect }])` then `addShape('rect', …)` on one slide:

```
id=2 name="Group 1"
id=4 name="Shape 0"    <- grouped child
id=3 name="Shape 0"    <- later top-level shape
```

`cNvPr` **ids** are correctly unique (`childIdxAlloc`, `src/gen-xml.ts:496`, is
seeded past the last top-level id and is regression-tested) — so the package is
valid and this is **not** a repeat of the media-id bug `3842f57f`. But Selection
Pane identity collides, and the duplicate-name check at `src/gen-xml.ts:467-468`
maps only top-level objects so it cannot catch it.

**This blocks the tracked feature in §4:** `groupObjects(objectNames)` is
name-keyed, and default names are not unique. Fix D5 first.

**Fix:** count group children too (thread a slide-wide per-kind counter through
`buildGroupObject`), and extend the duplicate check to recurse `_groupObjects`.

### D6 — Group names are not deterministic across presentations (P2) — **FIXED** (group names only; see `_chartCounter` below)

`_groupNameCounter` (`src/gen-objects.ts:126`) is module-global and never reset.
Verified — three **identical, independent** presentations in one process:

```
run1: [ 'name="Group 1"' ]
run2: [ 'name="Group 2"' ]
run3: [ 'name="Group 3"' ]
```

Same input, different bytes. `_chartCounter` (`:80`) has the same flaw but only
feeds internal part filenames; group names are user-visible in the Selection Pane.

**Fix:** reset per presentation (or derive per slide). Fold into D5 — both are
the group-naming model. Consider `_chartCounter` in the same pass.

**Carried forward:** `_chartCounter` is *not* fixed. Group names moved to a
per-slide counter, which `_chartCounter` cannot copy — chart part filenames
(`ppt/charts/chartN.xml`) must be unique per *presentation*, and
`addChartDefinition` only receives the slide, which holds no back-reference to
its presentation. So two identical decks built in one process still differ in
their chart part names. Internal-only (not user-visible like a Selection Pane
name), but it is the same "same input, different bytes" defect. Fixing it means
threading a presentation-scoped counter through `Slide` — record it separately
rather than absorbing it here.

---

## 3. Silent drops (write path)

Each fails quietly. Per AGENTS.md, none should be silent — minimum bar is a
`warn()`; the parenthesised note is the real fix.

- ~~**Animations on group children are dropped.**~~ **FIXED** (Phase 4).
  `resolveAnimationSpid` resolved `objectName` only against `slide._slideObjects` —
  group children were spliced out — and returned `null`, which the caller filtered
  with no warning. Now resolves through `collectSlideShapeIds` (every object on the
  slide, any depth); an unresolvable target warns that its effect was dropped.
- ~~**Connector shape-binding to a grouped shape fails.**~~ **FIXED** (Phase 4).
  Same lookup, same fix; the misleading "no shape with that objectName on the
  slide" now reads "no object with that objectName" and is emitted only when the
  name truly resolves to nothing.
- **Connectors cannot be group children.** `GroupChildProps` has no `connector`
  key and `addChildDefinition` (`src/gen-objects.ts:106-123`) has no branch, so
  `{ connector }` hits the generic "unrecognized child descriptor" warn+skip.
  (Note `{ line }` is `ShapeType.line` → a `<p:sp>`, not a `<p:cxnSp>`.)
- **`chart` / `table` / `media` / `placeholder` children** are rejected with a
  `warn` + skip (`src/gen-objects.ts:156-159`, "rels/ID/transform work pending").
  This is a **deliberate MVP scope**, documented in TSDoc and `CHANGELOG.md:940`.
  Listed for completeness — not proposed work unless a need appears.
- **Group fill / line / effects are not authorable.** `GroupProps` is
  `PositionProps + ObjectNameProps + rotate/flipH/flipV`, and `<p:grpSpPr>` emits
  only `<a:xfrm>`, though `CT_GroupShapeProperties` allows fill/effects — and the
  **read** API can already write a group fill. Write/read asymmetry.

### Adjacent risk (not group-specific)

`resolveAnimationSpid` validates `shapeIndex` only as `>= 0`, so an index past the
last top-level object emits `spid = shapeIndex + 2` targeting no shape — the
dangling-`spid` class the desktop-smoke skill names as a 0x80070570 source. Phase 4
un-silenced the *name* path around it but deliberately left this: it is a distinct
input-validation defect (and the fix, checking the index against the resolved shape
ids, is now one line away). Record separately rather than absorbing it.

`src/gen-xml.ts:1293` hardcodes `id="25"` for the slide-number placeholder. Since
group children allocate ids past `_slideObjects.length`, a grouped slide reaches
id 25 sooner than a flat one. Pre-existing (24 top-level objects suffices) and
strictly out of scope here — record separately rather than absorbing it.

---

## 4. Tracked feature — `dn-group-existing-slide-objects`

`docs/backlog.yml:1466` — status `target`, priority `p2`, `applies_to_current_project: yes`.
Proposes `slide.groupObjects(objectNames, options?)` to group already-authored
objects without replaying their descriptors.
next_action (`:1506`): *"Add a fixture-backed groupObjects API and document z-order semantics."*

**Dependency:** ~~blocked on **D5**~~ — **unblocked**: default names are now
unique slide-wide (Phase 2), so a name-keyed API can address every object.

Design questions to settle in the entry before coding: wrapper z-order (proposal:
the highest selected object's former slot); auto-bounds when no frame is given
(reuse `resolveObjBounds`, `src/gen-xml.ts:502`); behaviour for missing/duplicate
names (fail, per the footgun rule); and which kinds are accepted — existing groups
must be, for nested logical groups.

---

## 5. Test & doc gaps

**Untested but implemented:**
- **Group `rotate`/`flipH`/`flipV` on the write path.** Declared
  (`src/core-interfaces.ts:3094-3099`), emitted via the shared `locationAttr`
  (`src/gen-xml.ts:1266`), and verified by hand here —
  `<a:xfrm flipH="1" rot="2700000">` — but **zero tests** set them on a group.
- `objectLock` on groups (`GROUP_SHAPE_LOCK_ATTRS`, `src/gen-xml.ts:231`); group
  `altText`; empty group. (~~default `Group N` naming~~ and ~~no write→read
  round-trip test~~ are covered as of Phase 2; ~~explicit `x/y/w/h` on a group~~
  as of Phase 3.)

**Existing coverage (do not duplicate):** `test/regression/group-shapes.test.js`
(18 cases: identity xfrm, auto-bounds, unique ids, nested, unsupported-child warn,
plus Phase 2's cross-boundary name uniqueness, group-aware duplicate warning,
per-process group-name determinism, per-slide/inside-out group numbering, and the
write→read round-trip; plus Phase 3's partial-frame warn+fallback, complete-frame
verbatim, nested partial frame warning once with its parent sizing around the
fallback, and a partial-frame write→read round-trip; plus Phase 4's connector bound to
a grouped shape, animation targeting a nested group child, unresolvable-animation-target
warning, and top-level-wins-a-duplicate-name resolution — the first two assert against
the `cNvPr` ids parsed out of the emitted XML, which is what guards the id-allocation
mirror described in the Phase 4 notes),
schema fixtures `flat-group` / `nested-group` / `group-cross-references`
(`test/schema.test.js:2734`, `:2760`, `:2782`),
grouped + nested measured-fit (`test/regression/measured-fit-dist.test.mjs:158`,
`:173`), and rich read-side coverage against `group-transform.pptx`.

**Docs:** there is no groups page in `docs/` and no `addGroup` demo in `demos/`.
Write-side limitations live only in TSDoc; read-side ones are documented in
`docs/reference/pptx-read.md`.

**Ledger hygiene:** `group-shapes.test.js:7` and `schema.test.js:2732`,`:2758`
cite `upstream-issue-307`, which **does not exist** in `docs/backlog.yml`
(upstream tracking is retired). Dangling reference — retarget or drop the cite.

> Correction to a premise worth recording: the `group-scale` / `group-rot-flip`
> keys near `docs/backlog.yml:135` are the **`vocabulary.constructs` tagging
> list**, not a coverage-gap list. Being listed implies nothing about status, and
> both are currently tagged by zero items. `group-scale` is not applicable to the
> write path by design; `group-rot-flip` is implemented-but-untested there.

---

## 6. Sequencing

Ordered by (silently wrong × reachable), and by unblocking.

**Phase 1 — stop reporting wrong numbers.** D1 + D2 + D3. One file, one
CHANGELOG entry. Highest value: `inspect()` currently returns confidently wrong
geometry for any PowerPoint-touched deck, and the correct logic already exists to
delegate to. Land D1 first with a failing test asserting inspect ≈ `absoluteFrame`
on `group-transform.pptx`.

**Phase 2 — fix the naming model.** D5 + D6. Cheap, self-contained, and unblocks
§4. Add the missing write→read round-trip test here. **Done** — see the Phase 2
notes at the top; the round-trip test (`Presentation.load()` over an `addGroup()`
deck, asserting unique ids + names across the whole tree) landed in
`test/regression/group-shapes.test.js`, so §5's "no write→read round-trip test"
gap is closed.

**Phase 3 — close the degenerate-bounds trap.** D4. Needs the (a)/(b) decision
above; it is a behaviour change, so CHANGELOG it. **Done** — (b) was chosen; see
the Phase 3 notes at the top. Landed with four tests, a CHANGELOG entry marked
BREAKING, TSDoc on `GroupProps`/`addGroup` stating the all-or-nothing rule, and a
PowerPoint desktop smoke pass.

**Phase 4 — un-silence the drops.** §3 animation + connector name resolution
(shared fix), and correct the misleading connector warning. **Done** — see the
Phase 4 notes at the top. Landed with four regression tests, a schema fixture
(`group-cross-references`), a CHANGELOG entry, TSDoc on `startShape` /
`AnimationProps`, and a PowerPoint desktop smoke pass that checked the references
resolve, not just that the deck opens. The remaining §3 items are scope questions
for `addGroup` (connector children, group fill/line/effects), not defects of this
shape.

**Phase 5 — `groupObjects()`.** §4, fixture-backed, only after Phase 2.

**Phase 6 — cover and document.** Write-side rotate/flip/lock tests, a groups
docs page, an `addGroup` demo, and the `upstream-issue-307` cleanup. Fold the
`measure-fit.ts` invariant comment (§1) in here.

Phases 1–4 are independent of each other and can land in any order; only Phase 5
has a hard dependency (on Phase 2). Phases 1–4 have landed; Phase 5 is unblocked
and is next, with Phase 6 (cover and document) after it.

---

## 7. Verification (per AGENTS.md)

- Source changes: `pnpm run build` and `pnpm run typecheck`.
- Behaviour changes: `pnpm run test:unit`.
- Emitted-OOXML changes: add/update a fixture in `test/schema.test.js`, then
  `pnpm run test:schema` (needs `./tools/ooxml-validator/install.sh`).
- Backlog edits: `pnpm run backlog:validate` — must be clean before commit.
- Any change altering emitted group XML should also get a PowerPoint desktop
  smoke pass (the `powerpoint-desktop-smoke` skill) — group geometry defects of
  the D4 shape (zero extent) are exactly what schema validation accepts and
  PowerPoint rejects.
- Read-side behaviour must be asserted against a PowerPoint-authored fixture, not
  round-tripped XML (AGENTS.md, "Fixture-Gated Work"). `group-transform.pptx`
  already exists and covers scale/rot/flip/nesting — no new fixture needed for
  Phase 1.

## 8. Backlog reconciliation

Per AGENTS.md, a not-yet-implemented candidate belongs in `docs/backlog.yml`, and
work implemented immediately is recorded by commits + `CHANGELOG.md` instead —
not both. This document is a working plan, not a substitute for either.

- **D1–D6 are defects, not candidates.** If they are fixed now, they need no
  ledger entry — commit + CHANGELOG suffices. If any is deferred, add a
  `dn-<slug>` entry tagging `constructs: [group-scale]` (D1) or
  `[group-rot-flip]` (write-side tests) — which would make those two vocabulary
  keys used for the first time.
- **`dn-group-existing-slide-objects`**: on implementing, set `status:
  implemented`, `last_reviewed` to that date, `next_action: none`, and update
  `current_project_notes` / `evidence.local_files`.
