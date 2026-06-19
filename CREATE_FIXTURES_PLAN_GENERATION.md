# Generation-Feature Fixtures Plan — ON STAND-BY

**Status: on stand-by.** This plan is not active work. It captures the fixture
strategy for the two `interesting-with-tweaks` backlog items so it is not lost; do
not build any of it until the corresponding backlog item is scheduled. Track:
`upstream-pr-1447` (native PowerPoint comments, p2) and `upstream-pr-1431`
(animation engine, p3) in `docs/backlog.yml`.

Companion to `CREATE_FIXTURES_PLAN.md`, which covers the **read** fixtures in
`test/read/fixtures/`. The items here are different in kind: they are
**write/serialization** features, so they follow the generation-construct testing
pattern, not the read-fixture pattern.

## Testing pattern for these items

- **Primary evidence is a serialization schema fixture** in `test/schema.test.js`
  asserting the exact OOXML the writer emits ("option X → attribute/part Y"), run
  with `pnpm run test:schema` (validator-backed). This is the regression guard.
- **A PowerPoint-authored `.pptx` is an authoring oracle, not a read-harness
  fixture.** There is no read accessor for comments or animation timing, so a
  desktop-PowerPoint deck is used only to *discover and pin the correct OOXML shape*
  (preset IDs, part wiring, namespaces) that the schema fixture then asserts. Author
  it with the `powerpoint-fixture-authoring` skill and record provenance/SHA-256 as
  usual, but wire it into a comparison/inspection check rather than `test:read`.
- Both PRs' `interesting-with-tweaks` status means we take the OOXML intent and
  re-implement against the fork's ESM/TS boundary, discarding the upstream
  `dist`/CJS/bundle changes.

## A. Native comments (`upstream-pr-1447`) — comment parts + package wiring

PowerPoint has two comment formats; PowerPoint 365 writes **modern** comments but
still reads legacy. Decide which the fork emits (modern is the realistic target;
legacy is the simpler ISO/IEC 29500 baseline) and fixture accordingly.

- **Legacy comments (ISO/IEC 29500 §13):**
  - Per-slide Comments part `/ppt/comments/comment{N}.xml`, root `<p:cmLst>`,
    content type
    `application/vnd.openxmlformats-officedocument.presentationml.comments+xml`,
    implicit relationship from each authoring Slide part. Body:
    `<p:cm authorId dt idx><p:pos x y/><p:text>…</p:text></p:cm>`.
  - Presentation-level Comment Authors part `/ppt/commentAuthors.xml`, root
    `<p:cmAuthorLst>`, content type `…presentationml.commentAuthors+xml`, implicit
    relationship from the Presentation part. Body:
    `<p:cmAuthor id name initials lastIdx clrIdx/>`.
- **Modern comments (PowerPoint 2021+, MS-PPTX 2.16):** per-slide modern-comment
  part whose `<cm>` carries `id`/`authorId`/`created`/`status`/etc. and a `<txBody>`
  + `<replyLst>`, anchored by content monikers (`pc:sldMkLst`) rather than `pos`,
  plus a presentation-level `/ppt/authors.xml` (Authors part) listing `<author>`s.
- **Fixtures to add (with the feature):**
  - **Schema fixture(s)** in `test/schema.test.js`: assert the emitted comment part
    XML, the `commentAuthors`/`authors` part, the `Content_Types.xml` Overrides, and
    the slide→comments and presentation→authors relationships — for a deck with one
    author + one comment, and ideally a two-author / multi-comment case to pin `idx`
    per-author numbering.
  - **`comments-reference.pptx` oracle:** a desktop-PowerPoint deck with a couple of
    comments by two authors, used to confirm the exact part layout/attributes the
    schema fixture asserts. Read-harness wiring is **not** required (no read getter).

## B. Animations (`upstream-pr-1431`) — slide `p:timing` tree

Animation OOXML lives in the slide's `<p:timing>` element (child of `CT_Slide`),
which holds `<p:tnLst>` (the time-node tree of `par`/`seq`/`cTn` with behaviors
`anim`/`animEffect`/`animMotion`/`animRotation`/`animScale`/`set`/`cmd`) and
`<p:bldLst>` (per-shape build entries). Entrance / emphasis / exit / motion-path map
to Microsoft-proprietary `presetClass` / `presetID` / `presetSubtype` values
(MS-PPTX), so the correct integers must be captured from a real deck, not guessed.

- **Fixtures to add (with the feature):**
  - **One PowerPoint-authored oracle deck per animation class** (or one deck with a
    shape per class): a simple entrance (e.g. Fade/Appear), an emphasis, an exit, and
    a motion-path animation, each on a named shape. These capture the authentic
    `presetClass`/`presetID`/`presetSubtype` integers and the `tnLst`/`bldLst`
    structure — the magnitudes-from-real-PowerPoint principle already used for
    rotation angles in `CREATE_FIXTURES_PLAN.md` Remaining Work C.
  - **Schema fixture(s)** in `test/schema.test.js` asserting the `p:timing` tree the
    writer emits for a minimal "fade-in on one shape" model first, expanding per
    class as the engine grows. Start from the minimal timing/animation part model the
    backlog note calls for, validated against the oracle decks; do not port the
    upstream engine wholesale.
  - Keep this `p3`: per the backlog it is a low-priority research signal, so the
    oracle decks can be authored opportunistically to unblock the model design.

## Workflow reminder

For any new fixture or fixture addition: author with the
`powerpoint-fixture-authoring` skill, run `scripts/verify-powerpoint-fixture.ps1`,
confirm the target OOXML construct is present, update
`test/read/fixtures/README.md` (provenance, SHA-256, desktop-PowerPoint check date)
if it lands there, and pin the writer output with a `test/schema.test.js` fixture
(`pnpm run test:schema`).
