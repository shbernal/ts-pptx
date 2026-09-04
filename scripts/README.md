# `scripts/`

Build, gate and maintenance tooling. Every file here is plain ESM run on the repo's own
Node — no transpile step — and is typechecked (`tsconfig.scripts.json`, `checkJs`),
linted and formatted alongside `src/`.

**Annotate what you add.** That project runs with `noImplicitAny`, the same as `src/`, so
a parameter without a JSDoc `@param` fails `typecheck:scripts`. Types go in JSDoc rather
than in `.ts` deliberately: annotations buy the checking, and the file extension buys
nothing on top of it while costing every citation of these paths in `docs/` and every glob
in the lint, format and tsconfig lists. Reach for a type from the library itself
(`import('../dist/read.js').Slide`) rather than restating its shape — a hand-written
duplicate is what goes stale when the API moves.

The table exists because the single most useful fact about a script here is not what it
does but **whether anything runs it**. A gate nobody invokes reads as coverage while
proving nothing, which is how `script-roundtrip.mjs` sat runnable-but-unrun. Keep the
"Runs in" column true when you add or wire a script.

## Kinds

- **Gate** — asserts, exits non-zero on regression. Belongs in `verify`, `verify:full` or CI.
- **Generator** — writes artifacts (fixtures, docs, decks). Asserts nothing; re-run when
  its inputs move and commit the output.
- **Diagnostic** — measures and reports. Exits 0 whatever it finds, on purpose: the number
  is for a human to read, not for a build to fail on.
- **Library** — imported by other scripts, no CLI of its own.

## The scripts

Every `.mjs` under this directory has a row, subdirectories included, and `docs-check.mjs`
fails when one does not: the
premise above -- that the most useful fact about a script here is whether anything runs it --
is worth nothing if a script can be added without answering the question. Four were missing
when the check was written, including `path-refs.mjs` and `run-steps.mjs`, which between them
are in every aggregate the repo has.

| Script | Kind | What it does | Runs in |
|---|---|---|---|
| `alias-package.mjs` | Library | Repacks the tarball under the alias package name for the dual publish | `publish.yml` |
| `append-ceiling.mjs` | Diagnostic | What survives appending an authored slide to a template deck | manual (`read:append-ceiling`) |
| `browser-harness-server.mjs` | Library | Static server for the Playwright harness | started by `playwright.config.ts` |
| `bundle-size-ratchet.mjs` | Gate | Gzipped size of the browser entry's closure vs `bundle-size-budget.json` | `verify:full`, `check:package` |
| `byte-identity.mjs` | Diagnostic | Freeze/compare emitted bytes across a refactor; `prove-whitespace` discharges one recorded reformat | manual — see note below |
| `coverage-gate.mjs` | Gate | Per-area coverage thresholds from `coverage-gates.json` | CI (`coverage` job) |
| `coverage-merge.mjs` | Library | Merges Node + browser coverage into one report | `coverage:gate` |
| `coverage-project.mjs` | Library | Re-expresses the browser lane's hits in the Node report's own instrumentation maps, so merging can move counts but never the denominator | `coverage-merge.mjs` |
| `docs-api.mjs` | Generator | TypeDoc → markdown API pages | `docs:check`, `docs:prepare` |
| `docs-check.mjs` | Gate | Frontmatter, nav and link validation; with `--dist`, that every generated `llms.txt` URL names a built page | `docs:check`, so `verify` and `check:static`; twice more inside `docs:build` (source tree, then build), which is in `verify:full` and `docs.yml` |
| `docs-frontmatter.mjs` | Library | Frontmatter parsing shared by the `docs:*` scripts | — |
| `docs-index.mjs` | Generator | Rebuilds `docs/doc-index.md` | `docs:prepare` |
| `docs-list.mjs` | Diagnostic | Lists docs with their `read_when` hints | manual (`docs:list`) |
| `docs-new.mjs` | Generator | Creates a new doc page from the template | manual (`docs:new`) |
| `ensure-dist.mjs` | Gate | `dist/` freshness guard; builds, or `--check` fails | almost every `pnpm run` |
| `font-oracle-probe.mjs` | Gate | Which faces the measurement oracles resolve, and from where; fails when `FONT_ORACLES_GENUINE` names a family this machine does not have installed | `test:oracles`, CI (`font-oracles` job) |
| `gen-inspect-snapshot.mjs` | Generator | The inspect-surface snapshot | manual; asserted by a regression test |
| `generate-llms-docs.mjs` | Generator | `docs/public/llms*.txt` | `docs:prepare` |
| `install-hooks.mjs` | Library | Installs lefthook, skipping where it cannot | `prepare` |
| `libreoffice-render-smoke.mjs` | Gate | Renders decks in LibreOffice, the one renderer here with no SmartArt layout engine, and reads the painted text back | manual, needs LibreOffice + `pdftotext` (`test:lo`) |
| `note-census.mjs` | Diagnostic | How many fixtures raise each declared fidelity note, per tier | manual (`script:census`) |
| `ooxml-version-probe.mjs` | Diagnostic | Validator error counts across Office versions | manual (`schema:versions`) |
| `pack-utils.mjs` | Library | `pnpm pack` helpers for the two package gates | — |
| `package-lint.mjs` | Gate | `publint` + `attw` on the packed tarball | `verify:full`, `check:package` |
| `package-smoke.mjs` | Gate | Installs the tarball and exercises every subpath | `verify:full`, `check:package` |
| `path-refs.mjs` | Gate | Every backticked repo path in the tree must name a file that exists | `verify`, `check:static` |
| `png-utils.mjs` | Library | Minimal PNG encode/decode, for the gates that read pixels | `powerpoint-com-smoke.mjs`; unit-tested |
| `powerpoint-com-smoke.mjs` | Gate | Opens decks in desktop PowerPoint over COM | manual, Windows only (`test:com`) |
| `pptx-parts.mjs` | Library | Explode/diff `.pptx` packages | — |
| `raw-xml-ratchet.mjs` | Gate | Hand-built XML per file vs `raw-xml-budget.json` | `verify`, `check:static` |
| `read-blindness-census.mjs` | Diagnostic | Which OOXML the read model never looks at | manual (`read:census`) |
| `read-emit-edits.mjs` | Generator | Edited decks for the manual PowerPoint check | manual |
| `read-emit-roundtrip.mjs` | Generator | `load()`→`save()` decks for the manual PowerPoint check | manual |
| `run-steps.mjs` | Library | Runs a list of package scripts as one sequence; assembles all four aggregates | `verify`, `verify:full`, `check:static`, `check:package` |
| `script-roundtrip.mjs` | Gate | Generated script must rebuild the deck it came from | `verify:full`, CI |
| `script-utils.mjs` | Library | `ROOT`, `run()`, and the shared CLI front end | — |
| `sync-version.mjs` | Generator | Rewrites the `VERSION` constant in `src/presentation.ts` from `package.json` | the `version` lifecycle script (`pnpm version …`); `--check` manual (`version:check`) |
| `xml-equivalence.mjs` | Library | Proves two XML parts differ only in inert inter-element whitespace | `byte-identity.mjs prove-whitespace`; unit-tested |
| `com/contract.mjs` | Library | The shape names, `ProgID`s and `PpActionType` values the COM decks, VBScripts and verifiers all have to agree on | `powerpoint-com-smoke.mjs` |
| `com/decks.mjs` | Library | Builds the four decks the COM smoke drives, from the current `dist/` | `powerpoint-com-smoke.mjs` |
| `com/vbs.mjs` | Library | The VBScript sources that drive desktop PowerPoint, one per deck | `powerpoint-com-smoke.mjs` |
| `comparison/health.mjs` | Library | Activity, adoption and source size for both projects, from the GitHub and npm APIs and a shallow clone | `comparison/measure.mjs` |
| `comparison/hygiene.mjs` | Library | What each library costs to install and to ship, from clean per-library installs | `comparison/measure.mjs` |
| `comparison/measure.mjs` | Generator | Builds every probe with ts-pptx and with upstream pptxgenjs, reads the emitted parts, measures the other three families, and writes `comparison/snapshot.json` | manual |
| `comparison/probes.mjs` | Library | The construct-coverage corpus: one deck intent per probe, expressed in each library's own idiom | `comparison/measure.mjs` |
| `comparison/render.mjs` | Generator + Gate | Renders `comparison/snapshot.json` into `docs/comparison.md` and the generated region of `README.md`; `--check` re-renders in memory and fails on drift | manual |
| `comparison/unavailable.mjs` | Library | The one shape a measurement takes when a fetch failed, and the walk that finds them in a finished snapshot | `comparison/measure.mjs` |
| `comparison/validity.mjs` | Library | Runs the decks the corpus built through the `test:schema` oracle, per library | `comparison/measure.mjs` |
| `gate-decks/chart-matrix.mjs` | Library | Gate deck reaching the chart emitters the showcase corpus never does | `byte-identity.mjs` |
| `gate-decks/index.mjs` | Library | The gate-deck registry — one list, so a deck cannot go undiffed | `byte-identity.mjs` |
| `gate-decks/shape-matrix.mjs` | Library | Gate deck reaching the slide-object constructs a presentation deck has no reason to carry | `byte-identity.mjs` |

### Why four of these are manual on purpose

Being unwired is a bug for a gate and a design decision for everything else. These four
are the second kind, and the reasons are worth not re-litigating:

- **`byte-identity.mjs`** compares against a baseline you freeze *before* a refactor. As a
  CI gate it would fail every feature PR that legitimately changes emitted bytes — which
  is the opposite of its purpose. It is an instrument for a refactor in progress.
- **`ooxml-version-probe.mjs`** spawns the validator seven times per fixture and asserts
  nothing `test:schema` does not already assert at Microsoft365. Its own header says so.
- **`read-emit-roundtrip.mjs`** and **`read-emit-edits.mjs`** write decks for a human to
  open in desktop PowerPoint and confirm no repair prompt. CI has no PowerPoint; running
  them there would produce files nobody opens.

### Where fixture tooling lives, and why not here

**No script that writes a committed fixture or fixture sidecar belongs in this directory.**
That whole chain lives in `test/read/fixtures/authoring/`, beside the fixtures it produces,
whatever it is written in and whatever platform it needs — the `author-*.ps1` COM recipes,
`gen-cases.mjs` and `measure-lo.py` which feed them, and the Node builders that derive the
committed `*.oracle.json` / `autofit-calibration.json` sidecars from the finished decks.
This directory is for build, gate and maintenance tooling for the repo itself.

Getting that line right took two passes, and the wrong one is instructive. `gen-cases.mjs`
and `measure-lo.py` used to sit here *and* in the authoring directory, one pair byte-identical
and one near-identical, with nothing keeping them in step. Deduplicating them onto the
authoring side produced a rule stated as *feeds the authoring of a fixture → there, derives
from an already-committed fixture → here* — which sounded principled and was false on the
day it was written. `build-oracles.mjs` and the three `build-*-oracle.mjs` scripts derive
from already-committed fixtures, are pure cross-platform Node, and had always lived in the
authoring directory. The rule described five files that were on the other side of it.

`extract-autofit-calibration.mjs` was the one file that rule fit, and only because it was the
one that happened to be here. It is indistinguishable in kind from `build-oracles.mjs` — read
a committed `.pptx` with `fflate`, write a committed JSON sidecar next to it — so it moved
too, and the line above is the one that every file now actually falls on. "Runs without a
desktop app" turned out to describe a property of individual recipes, not a directory
boundary; the authoring README marks which of its scripts need Windows.

The stale-copy mechanism was worth understanding rather than just deleting: `author-all.ps1`
loaded its engine and measure script from `.tmp/` rather than from `$PSScriptRoot`, so the
file that actually ran was whatever had last been staged into scratch — neither tracked copy.
That indirection is gone too; every recipe there now resolves from its own location.

## Conventions

**A shebang marks an entry point.** Anything invoked as a command — by `package.json`,
`lefthook.yml`, `playwright.config.ts` or a human — starts with `#!/usr/bin/env node`.
The four library modules (`docs-frontmatter`, `pack-utils`, `pptx-parts`, `script-utils`)
have none, so the first line tells you which kind of file you opened. The **exec bit is
deliberately not part of this**: every script is invoked as `node scripts/x.mjs`, never
`./scripts/x.mjs`, so all files are tracked `100644` and `core.filemode` is `false` on
Windows checkouts anyway. Do not `chmod +x` a script — it only adds a mode change to the
diff and makes the tree inconsistent.

**Every script takes `--help`.** Flags are parsed with `parseCli`/`parseCliOrExit` from
`script-utils.mjs`, which wraps `node:util`'s `parseArgs` and turns an unknown flag into
one line plus usage rather than a stack trace. Do not hand-roll `argv.indexOf('--flag') + 1`:
it silently takes the *next flag* as the value when the argument is omitted.

**A gate with parsing logic exports it and guards its CLI** behind `isMain(import.meta.url)`,
so `test/scripts/` can exercise the logic without the script measuring `dist/`, rewriting a
budget file, or calling `process.exit` inside the test runner. `bundle-size-ratchet.mjs`
and `raw-xml-ratchet.mjs` are the worked examples.

**Prefer `process.exitCode` over `process.exit()`** in anything with buffered output —
`runCli()` handles this — so a final `console.log` is not truncated on the way out.

**A ratchet's failure mode is silence.** Over-counting fails loudly and gets fixed;
under-counting reports an improvement and passes. When you touch scanning or parsing in
`bundle-size-ratchet.mjs` or `raw-xml-ratchet.mjs`, add the case to `test/scripts/` and
check the test actually fails against the old code before you trust it.
