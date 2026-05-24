# Task brief — automate validation pipeline

This is a feature-/infra-shaped task, not a bug fix. Two goals, in order.
Goal 1 must be complete before Goal 2 is valuable. Goal 3 is independent
of Goals 1–2 and may proceed in parallel.

## Context

- `npm test` runs a custom Node runner (`test/run.js`) over
  `test/bug-*.test.js` (58 cases, regex assertions on emitted XML).
- `npm run schema-test` runs `test/run-schema.js` over
  `test/schema.test.js` (8 fixtures), validating each generated `.pptx`
  with Microsoft's `OpenXmlValidator` via `tools/ooxml-validator/bin/OOXMLValidatorCLI`.
- `.autoloop/bug-report.md` is the bug-fix queue. The single open bug is
  **B20**: `<p:presentation>` child elements are emitted in the wrong order
  per CT_Presentation; every schema-test fixture trips the same error.
- `TESTING.md` is the manual cross-platform release-validation
  procedure (browser desktop + iOS, Node CLI, Node stream, Vite/TS,
  Web Worker, Microsoft 365 viewer). Today it is entirely human-driven.

## Goal 1 — Close B20 and merge schema-test into `npm test`

### 1a. Fix B20

Read `.autoloop/bug-report.md` for the full report. Summary:

- File: `src/gen-xml.ts`, function `makeXmlPresentation` (or wherever
  `<p:presentation>` is constructed; validate via grep, not memory).
- Currently emits `<p:notesMasterIdLst>` after `<p:notesSz>`.
- OOXML CT_Presentation requires order:
  `sldMasterIdLst, notesMasterIdLst, handoutMasterIdLst, sldIdLst,
  sldSz, notesSz, smartTags?, embeddedFontLst?, ...`
- Reorder the emission so every child appears in CT_Presentation
  position. Confirm `sldMasterIdLst`, `sldIdLst`, `sldSz` are also in
  the right slots while you are in there.

Acceptance:

- `npm run schema-test` reports 0 errors across all 8 fixtures.
- `npm test` (regex suite) continues to pass — 58/58.
- Add a regression case to `test/bug-*.test.js` that asserts ordering
  in `ppt/presentation.xml` directly (so the regression survives even
  if the schema validator is unavailable).
- `.autoloop/bug-report.md` is updated: B20 moves out of "Open queue"
  into "Already fixed" with the new commit SHA. P1 tier list updated.

### 1b. Fold schema-test into `npm test`

Once 1a is green:

- Change `package.json` `test` script from `node test/run.js` to
  `node test/run.js && node test/run-schema.js`.
- The `pretest` script already builds, so no further changes there.
- README.md gets one new line under the install / usage section
  explaining that `npm test` requires `./tools/ooxml-validator/install.sh`
  to have been run once. Preserve the existing demo guidance.

Acceptance:

- `npm test` runs both stages; full output shows
  `Passed: 58  Failed: 0` then `Passed: 8  Failed: 0`.
- A fresh clone followed by `./tools/ooxml-validator/install.sh && npm test`
  works end-to-end on macOS arm64 (the autofix loop's host).

## Goal 2 — Automate `TESTING.md` artefact generation

`TESTING.md` describes per-platform manual validation. Most of it is
either runnable demos (browser HTML, Node CLI, Vite SPA) or capture
steps (open a browser, look at the page). Wherever the validation
amounts to "run this demo, then check the output", it is automatable.
Wherever it is "render in PowerPoint on iOS", it is not — leave those
sections manual.

Scope for this goal:

### 2a. Headless browser harness

- Add `playwright` (or `puppeteer`) as a `devDependency`.
- New `test/release/browser.test.js`:
  - Launch the existing demo at `demos/browser/index.html` via the
    bundled `node browser_server.mjs` (already in `demos/`).
  - Click each top-level demo button, intercept the resulting
    `.pptx` download into a temp dir.
  - For each downloaded file, invoke `OOXMLValidatorCLI` and assert
    zero schema errors.
  - Also exercise `worker_test.html` (Web Worker path) the same way.

### 2b. Node demo harness

- New `test/release/node.test.js`:
  - Run `npm --prefix demos/node run demo` and `demo-all` and `demo-stream`
    as child processes, capturing exit code and any `.pptx` files written.
  - Validate each output via `OOXMLValidatorCLI`.
  - For `demo-stream`, read the streamed bytes back, write to disk,
    validate.

### 2c. Vite demo build harness

- New `test/release/vite.test.js`:
  - Run `npm --prefix demos/vite-demo run build`. Assert clean exit
    and that `dist/` contains the expected entry chunk.
  - (Skip running `npm run dev`; that is interactive and not part of
    a release gate.)

### 2d. Wiring

- New npm script `release-test` that runs **all three** above plus
  the regular `npm test`. This is what `TESTING.md` step 1
  ("Run `npm run ship`") effectively becomes once automation is in.
- Update `TESTING.md`:
  - Replace the manual "click each demo" step with `npm run release-test`.
  - Keep manual sections only for things that genuinely require human
    eyes: M365 web viewer upload, iOS rendering, PowerPoint visual
    inspection. These remain a release-time manual checklist.
  - Keep the existing checklist table; mark which rows are now
    automated vs which remain manual.

Acceptance:

- `npm run release-test` runs on macOS arm64 cleanly given a fresh
  install plus `./tools/ooxml-validator/install.sh`.
- Every automated row in the `TESTING.md` checklist table has a
  corresponding `test/release/*.test.js`.
- Manual rows are clearly marked and explain why they cannot be
  automated yet.

### 2e. Out of scope here

- **PowerPoint screenshot regression** (open generated `.pptx` in
  PowerPoint, capture per-slide PNGs, diff against baselines). This
  needs a self-hosted Windows or macOS runner with Office installed
  and a baseline-blessing workflow, plus a license decision. Treat as
  a separate, later task brief.
- **LibreOffice headless conversion** as a fidelity proxy. The user
  declined this approach earlier; do not add it.

## Goal 3 — Update-checker integration (small, parallel)

- Add `npm run check-validator` script that runs
  `node tools/ooxml-validator/check-updates.js`.
- The script already exists and exits 0 (current), 1 (newer), 2 (error).
- No automated invocation yet — leave it as a manually-run developer
  command. CI integration can come later.

## Conventions

- Conventional Commits everywhere (`fix(...)`, `test(...)`, `chore(...)`,
  `feat(...)`, `build(...)`, `docs(...)`).
- Commit each goal's work as a small set of focused commits, not one
  mega-commit. Goal 1 is naturally two commits (1a, 1b). Goal 2 is
  naturally one commit per harness plus one for `release-test` wiring.
- Update `CHANGELOG.md` `## [Unreleased]` for any user-visible change
  (Goal 1a is user-visible — a previously-corrupt OOXML structure is
  now correct).
- Do not introduce new external services, secrets, or non-MIT-compatible
  dependencies. Playwright (Apache 2.0) and Puppeteer (Apache 2.0) are
  acceptable; bias to Playwright for cross-browser support.
- All new test code lives under `test/`. Release-time tests under
  `test/release/` so they can be excluded from the inner-loop
  `npm test` if desired (the runner regex would need extending, but
  keep `npm test` fast — release-test is the slow lane).

## Done when

- B20 is closed (Goal 1a) and `npm test` enforces schema validity
  (Goal 1b).
- `npm run release-test` exists and runs the automatable subset of
  `TESTING.md` end-to-end without human input (Goal 2).
- `TESTING.md` reflects what is automated vs what remains manual.
- All new commits are clean (no autoloop-internal IDs leaking;
  no AWS/Amazon/internal references; deterministic on a fresh clone).
