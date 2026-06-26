# Phase 2 Fixture Authoring Plan — Animations & Transitions

Scope: author the PowerPoint-authored `.pptx` fixtures + oracle JSONs that gate
Phase 2 of `docs/animations-and-transitions.md`. Per `AGENTS.md` ("OOXML And
PowerPoint Work" → fixture-gated work), each Phase 2 capability is **blocked**
until its fixture/oracle exists; this document is the authoring spec so the
fixtures can be produced against a fixed target, then the code written against
them.

Status as of 2026-06-26: Phase 1 verified green (typecheck/build OK; 19
anim/transition tests + 110 schema-validation tests pass) but **staged,
uncommitted**. Commit Phase 1 as a clean baseline before starting Phase 2.

---

## Phase 2 capabilities and their fixture gates

| # | Capability (from doc §"Out of scope / phasing") | Code groundwork present | New fixture needed |
| - | ----------------------------------------------- | ----------------------- | ------------------ |
| A | Carry a shape's build animation through `importShape` (remap into destination timing) | `remapAnimationSpids` + enumerate/prune helpers (Phase 1) | `import-animation-merge.pptx` — cross-slide copy oracle |
| B | Expand the preset effect set | `slideTimingToXml` template emit (Phase 1) | `slide-animation-presets.pptx` — one shape per new preset |
| C | Transition sounds (`sndAc` + audio rels) | none | `slide-transition-sound.pptx` — built-in + embedded sound |

These are independent; author and land them one at a time. Companion backlog:
`upstream-pr-1431` (the umbrella, Phase 1 `implemented`) and `dn-1863`
(`importShape` drops animation — `deferred`, p3, low value: only 3/640 reference
slides carry animation). Prioritize B or C over A unless downstream needs A.

---

## Shared workflow (every fixture)

Use the **`powerpoint-fixture-authoring`** skill
(`.agents/skills/powerpoint-fixture-authoring/`). Requires Windows desktop
PowerPoint driven over COM via PowerShell 7. Per-fixture steps:

1. From repo root, `git status --short`; clear `Resiliency\DocumentRecovery`
   first (per doc §"Fixtures + oracles").
2. Write the COM authoring script to `.tmp/author-<name>.ps1` and run it with the
   call operator `& '.tmp/author-<name>.ps1'` (never
   `powershell.exe -ExecutionPolicy Bypass …` — trips the sandbox classifier).
   Snapshot pre-existing `POWERPNT` PIDs so the reap only kills the spawned
   server.
3. Author minimal & deterministic: explicit 16:9 slide size, stable shape names,
   fixed coords/colors, no external assets unless the construct requires them.
   Save via `Presentation.SaveAs()`, set `Saved = $true`, close, quit, release
   COM, confirm no `POWERPNT` remains.
4. **Verify**: reopen once via COM with no repair prompt
   (`scripts/verify-powerpoint-fixture.ps1`); confirm `docProps/app.xml` shows
   `Microsoft Office PowerPoint` + `AppVersion`; dump the target slide XML with
   `scripts/dump-slide-xml.ps1`; compute SHA-256.
5. **Author the oracle JSON** next to the `.pptx` (`<name>.oracle.json`) matching
   the existing schema — see field maps below. Oracles embed *verbatim* XML
   (whitespace-exact, since Phase 1 asserts byte-for-byte) plus decoded fields.
6. Update `test/read/fixtures/README.md`: provenance, SHA-256, purpose,
   PowerPoint check date, and tick the verification checklist.
7. Update `docs/backlog.yml` (`upstream-pr-1431` notes + the relevant companion
   item) and run `pnpm run backlog:validate`.
8. Write the tests, then `pnpm run test:read`, `pnpm run test:unit`,
   `pnpm run test:schema`.

### Oracle JSON field reference (from Phase 1)

- Animation oracles (`slide-animation-basic/rich.oracle.json`):
  `deck, schema, application, appVersion, sha256, notes, shapeIds, effects,
  animationSpids, bldList, timingXml`.
- Transition oracle (`slide-transition.oracle.json`):
  `deck, schema, application, appVersion, sha256, notes, slides, entryEffectTable`.

New oracles must reuse these key names; add new keys only where a capability
introduces genuinely new structure (noted per fixture below).

---

## Fixture A — `import-animation-merge.pptx`

**Unblocks:** carry a build animation through `importShape` (remap into the
destination `p:timing`); the not-yet-wired `importSlide` preserve/restyle remap.

**Why a PowerPoint fixture:** the *merged end-state* PptxGenJS would emit is our
own output, but we need PowerPoint's ground truth for **how spids are renumbered
and how two build sequences merge** when a shape is copied into a slide that
already has animation. That observed behavior is the oracle the remap/merge is
verified against.

**Deck (2 slides):**
- Slide 1: one named shape ("Source") with an entrance build (Fade on click) —
  mirrors `slide-animation-basic`.
- Slide 2: one named shape ("HostExisting") with its own entrance build, then —
  in PowerPoint — copy "Source" from slide 1 onto slide 2 and keep its
  animation. Save. Slide 2 now has two builds with PowerPoint's renumbered spids.

**Oracle (`import-animation-merge.oracle.json`):** animation-oracle shape; pin
slide 2's merged `timingXml`, `bldList`, `animationSpids`, `shapeIds`, and the
per-effect `effects` triples. Add a key `mergeMap` documenting source-shape
spid → renumbered spid on slide 2, and the relative ordering of the two builds.

**Tests:** extend `test/read/animations-transitions.test.js` to drive
`importShape` from slide 1 → a destination, assert `remapAnimationSpids` +
timing merge match the oracle; assert no dangling `spTgt`/`bldP` on either side.

---

## Fixture B — `slide-animation-presets.pptx`

**Unblocks:** expanding the supported preset set (write-side templates).

**Deck (1 slide, one shape per new preset):** decide the target preset list
first (see Open Questions), then author one labeled shape per preset using the
`MsoAnimEffect` COM code for each, e.g. entrance Fly-In, Wipe, Appear; emphasis
Spin / Pulse; exit Fly-Out — extending the set already proven in
`slide-animation-rich` (Fade entr 10, Grow/Shrink emph 6 via `MsoAnimEffect=59`,
exit Fade 10). Use one trigger family per shape so each template is isolated.

**Oracle (`slide-animation-presets.oracle.json`):** animation-oracle shape. The
critical artifact is, per preset, the verbatim effect node keyed by
`(presetID, presetClass, presetSubtype, nodeType)` — these become the new
write-side templates (parameterized only by `spid`/`delay`/`dur`). Pin
`timingXml`, `bldList`, `effects`, `animationSpids`, `shapeIds`, plus a
`presetTemplates` map: preset name → `{ key, nodeXml, bldPXml }`.

**Probing note:** the `MsoAnimEffect → (presetID, presetClass, subtype)` mapping
is PowerPoint-specific and not in the XSD. Probe by iterating candidate
`MsoAnimEffect` codes, applying each, dumping the resulting node, and recording
the triple — mirrors how the 158-row `PpEntryEffect` table was built in Phase 1.

**Tests:** add a `test/schema.test.js` fixture emitting each new preset (validate
clean), and a `test/regression/animations-transitions.test.js` assertion that
each emitted node equals the oracle template byte-for-byte.

---

## Fixture C — `slide-transition-sound.pptx`

**Unblocks:** transition sounds (`sndAc` child of `p:transition` + audio rels).

**OOXML target:** `sndAc` follows the transition-type element inside
`p:transition` and references a sound via relationship; an embedded sound also
pulls in an audio part + content-type, like the existing `av-media` fixture's
audio rel graph (`…/2006/relationships/audio`, `audio/mpeg`). Confirm exact
`sndAc`/`stSnd`/`p:snd` structure and the `loop`/`builtIn` attributes against the
`ooxml` MCP (`CT_TransitionStartSoundAction` / `CT_EmbeddedWAVAudioFile`) before
authoring.

**Deck (≥2 slides):**
- Slide 1: transition with a **built-in** PowerPoint sound (e.g. Applause) —
  observe whether PowerPoint emits a rel or an inline reference.
- Slide 2: transition with an **embedded custom** sound — observe the audio part,
  the `r:embed`/`r:link` on `sndAc`, the `.rels` entry, and `[Content_Types].xml`.
- Optionally a slide with `sndAc` + the "stop previous sound" form to pin
  `stSnd`/`endSnd` if we intend to support it.

**Oracle (`slide-transition-sound.oracle.json`):** transition-oracle shape
(`slides` + per-slide `transitionXml`). Add a `soundRels` key capturing, per
slide, the `sndAc` XML, the relationship id/target/type, and the content-type
override; note built-in vs embedded handling.

**Tests:** read-side decode of `sndAc` into the transition model; write-side emit
that reproduces both the `sndAc` XML and the rel/content-type graph; a
`test/schema.test.js` fixture validating the embedded-sound package.

---

## Open questions to settle before authoring

1. **Preset list for Fixture B** — which entrance/emphasis/exit presets are in
   scope? (Drives shape count and probing effort.) Recommend the common slide-
   factory set: entrance fadeIn/flyIn/appear/wipe, emphasis pulse/spin, exit
   fadeOut/flyOut.
2. **Transition-sound surface (Fixture C)** — support embedded sounds (needs rel
   plumbing in the writer) or built-in only first? Built-in is cheaper; embedded
   is the higher-value, downstream-relevant case.
3. **`importShape` carry (Fixture A) priority** — `dn-1863` is parked at p3.
   Confirm a real deck needs it before investing, or defer A and ship B/C.
4. **Oracle whitespace** — Phase 1 asserts byte-for-byte; confirm the new oracles
   capture the exact serialization PowerPoint emits (indentation, self-closing
   form) so the regression assertions hold.

---

## Definition of done (per fixture)

- `.pptx` + `.oracle.json` in `test/read/fixtures/`, reopened clean via COM
  (no repair prompt), SHA-256 recorded.
- `test/read/fixtures/README.md` updated (provenance/hash/purpose/check date +
  checklist tick).
- `docs/backlog.yml` updated and `pnpm run backlog:validate` clean.
- Read + regression + schema tests added and green
  (`test:read`, `test:unit`, `test:schema`).
- `docs/animations-and-transitions.md` §"Out of scope / phasing" updated to move
  the capability from Phase 2 candidate → implemented.
