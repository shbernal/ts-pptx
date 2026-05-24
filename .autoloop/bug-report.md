# PptxGenJS — bug fix queue

This is the source of truth for the autofix loop. Each `## B<n>:` section
under **Open queue** below is one bug. Process top to bottom.
**Commit each fix directly to `master`.** Do **not** create per-bug
branches. Do **not** open PRs.

Fixes that have already shipped are listed under **Already fixed** for
traceability. The autoloop should not re-process those — they each
have a corresponding commit on `master`.

Project conventions:

- `README.md` — install / build commands and library overview.
- `CHANGELOG.md` — [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
  format. Append a `### Fixed` entry under an `## [Unreleased]` section
  when closing a fix.
- `TESTING.md` — manual cross-platform release validation. Out of scope
  per bug; run before cutting a release, not per fix.

For each open bug:

1. Read `README.md` and `package.json` at the repo root to confirm the
   project's install/build/test commands. The npm scripts are
   `npm run build` (rollup) and `npm test` (custom runner in `test/run.js`).
2. Reproduce against current `master` using the snippet under "Reproduction."
   If you cannot reproduce, mark the bug `cannot-reproduce` in
   `.autoloop/progress.md` and move on — do not fabricate a fix.
3. Locate the emit site. The TS source lives in `src/`. Most XML emission is
   in `src/gen-xml.ts`; chart emission in `src/gen-charts.ts`; object intake /
   mutation in `src/gen-objects.ts`. Validate via grep, not memory.
4. Implement the **minimal** fix. No refactoring beyond what the bug requires.
5. Add a regression test under `test/` following the pattern in
   `test/bug-*.test.js` (custom runner; see `test/run.js` and
   `test/helpers.js`). The test must fail before the fix and pass after.
6. `npm install && npm run build && npm test`. All existing tests must stay
   green. If the bug touches OOXML structure, also run `npm run schema-test`
   (requires `./tools/ooxml-validator/install.sh` first).
7. Update `CHANGELOG.md`: add a bullet under `### Fixed` in the
   `## [Unreleased]` section (create the section at the top of the file if
   it does not exist) referencing the bug ID and any upstream issue
   number(s).
8. **Commit the fix directly to `master`** with a Conventional Commits subject:
   `fix(<area>): <summary>`. Body should reference the bug ID and any
   upstream issue number(s).
9. Mark the bug as `closed` in `.autoloop/progress.md` and append a one-line
   entry to `.autoloop/fix-log.md` (`Bxx: <commit sha> <one-line summary>`).
10. **Continue to the next unprocessed bug.** Do not emit `task.complete`
    until every bug in the **Open queue** is marked `closed` or
    `cannot-reproduce` in `.autoloop/progress.md`.

If a bug already has an open upstream PR that fixes it correctly, study the
PR for guidance, but write the fix as a fresh commit on `master` of this
fork (no PR review, no cherry-pick — we are not pushing to upstream).

Priority tiers (open queue only):

- **P1 (file-corruption / "needs repair"):** B20
- **P2 (silent data loss / wrong output):** (none currently open)
- **P3 (cosmetic / non-breaking):** (none currently open)

Process P1 first, then P2, then P3.

Termination contract:

- Emit `task.complete` **only** when every `## B<n>:` heading in the
  **Open queue** has a corresponding `closed` or `cannot-reproduce`
  entry in `.autoloop/progress.md`.
- Until then, after each `bug.closed`, hand back to the diagnoser to pick
  the next unprocessed bug.

---

# Open queue

## B20: `<p:presentation>` child elements emitted out of order

**Priority:** P1  
**Source:** Schema validation — every fixture in `test/schema.test.js` trips this same error.

### Symptom

`ppt/presentation.xml` emits `<p:notesMasterIdLst>` after `<p:notesSz>`.
The OOXML CT_Presentation type requires children in a specific order:

```
sldMasterIdLst, notesMasterIdLst, handoutMasterIdLst,
sldIdLst, sldSz, notesSz, smartTags?, embeddedFontLst?, ...
```

The current emitter places `notesMasterIdLst` AFTER `notesSz`, which is
invalid OOXML. Microsoft's `OpenXmlValidator` flags this on every deck
PptxGenJS produces today.

### Reproduction

Every deck reproduces this — validate any output via the configured
schema-test target:

```javascript
var pptxgen = require('pptxgenjs');
var pres = new pptxgen();
pres.addSlide();
pres.stream().then(buf => {
  require('fs').writeFileSync('repro.pptx', buf);
});
// Then: ./tools/ooxml-validator/bin/OOXMLValidatorCLI repro.pptx
// Or:   npm run schema-test  (all 8 fixtures fail with this single error)
```

Validator output (Microsoft `OpenXmlValidator` via `OOXMLValidatorCLI`):

```
[Schema] The element has unexpected child element 'notesMasterIdLst'.
List of possible elements expected: <notesSz>.
Path: /ppt/presentation.xml   Id: Sch_UnexpectedElementContentExpectingComplex
```

### Root-cause hypothesis

`makeXmlPresentation` in `src/gen-xml.ts` constructs `<p:presentation>`
by appending children in source-code order rather than schema-canonical
order. Reorder the emission so `<p:notesMasterIdLst>` (and any other
out-of-position children) appear in their CT_Presentation positions.

### Acceptance criteria

- `npm run schema-test` reports zero schema errors against the existing
  fixture set (the only currently-flagged error is this one; fixing it
  should clear all 8 fixtures).
- Existing `npm test` (58 regex tests) continues to pass.
- Every other element in `<p:presentation>` (e.g. `sldMasterIdLst`,
  `sldIdLst`, `sldSz`) is also confirmed to be in the canonical order
  while you're in there.

### References

- OpenXmlValidator error id: `Sch_UnexpectedElementContentExpectingComplex`
  at XPath `/p:presentation[1]`, part `/ppt/presentation.xml`.
- OOXML CT_Presentation child sequence: ECMA-376 Part 1 §19.2.1.26.

---

# Already fixed

For traceability — each entry has a corresponding commit on `master`.
The autoloop should not re-process these. Once verified clean by the
schema validator and any future release tag, this section is a candidate
for archival to a separate file (e.g. `.autoloop/closed.md`).

| Bug | Summary | Commit |
|-----|---------|--------|
| B1  | Duplicate `<a:pPr>` per `<a:p>` from mixed-formatting runs | `1a557524` |
| B2  | Phantom slideMaster Override entries in `[Content_Types].xml` | `47f54555` |
| B4  | `writeFile()` mutated option objects (pt → EMU double-conversion on second call) | `780dddf1` |
| B5  | Reusing one shadow object across `addShape` calls (subsumed by B4 fix; regression test) | `142743ee` |
| B6  | Combo charts emitted too few axis defs vs IDs referenced | `2f7759e9` |
| B7  | Shape `shadow.color` with `#` prefix produced corrupt OOXML | `aadc3b8d` |
| B8  | 8-character hex (RGBA) silently corrupted the file in fill/text/line/shadow paths | `e72c859a` |
| B9  | Unicode bullet glyphs duplicated when `bullet: true` | `86413b21` |
| B10 | `addShape("oval", ...)` wrote invalid OOXML preset; alias map fix | `791eab35` |
| B11 | `addShape("roundedRectangle", ...)` same class as B10; resolved by the same alias map | `791eab35` |
| B12 | Solid-color slide background missing `<a:effectLst/>` triggered repair dialog | `1ec0bb09` |
| B13 | Textless shapes emitted `<p:sp>` missing `<p:txBody>` | `a67baf21` |
| B14 | Non-numeric table cell margin produced `NaN` XML attributes | `2bfab5ae` |
| B15 | notesMaster placeholder shapes/theme rels malformed (theme2.xml emit) | `a1b64cc5` |
| B16 | Unused Default extension types emitted in `[Content_Types].xml` | `f63f6492` |
| B17 | Empty `ppt/charts/` and `ppt/embeddings/` directories created unconditionally | `07064047` |
| B18 | Placeholder objects created as TEXT instead of PLACEHOLDER | `91da6a68` |
| B19 | `bullet: { type: "bullet" }` produced no bullet | `d6a1de4b` |

---

# Not reproduced

## B3: Dangling `<Relationship>` entries in `*.rels`

The original report acknowledged this could be `cannot-reproduce` on
current `master`. No deterministic repro was found and no fix was
attempted. If new evidence appears (e.g. a deck where some relationship
target is missing from the archive), re-open as a new bug entry with a
concrete reproducer.

---

## Out of scope for this loop

Listed for future work, not for autofix to pick up:

- **#1396 Charts not showing in Apple Numbers** — has a dedicated
  `issue-1396` upstream branch and is Keynote/Numbers-rendering specific;
  larger scope than a focused autofix.
- **Slide transitions, multi-column text, shape grouping** — feature
  gaps not currently tracked as defects.
- **#1349 RTL/LTR mixed text rendering, #1262 hyperlink WPS, #1306
  graphs in Mac Keynote, etc.** — third-party renderer quirks, not
  bugs in PptxGenJS XML emission.
- Any issue older than two years with no recent activity — likely
  resolved in v3.13/v4 era; let triage flag separately if still
  reproducible.
