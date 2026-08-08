# `scripts/`

Build, gate and maintenance tooling. Every file here is plain ESM run on the repo's own
Node — no transpile step — and is typechecked (`tsconfig.scripts.json`, `checkJs`),
linted and formatted alongside `src/`.

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

| Script | Kind | What it does | Runs in |
|---|---|---|---|
| `append-ceiling.mjs` | Diagnostic | What survives appending an authored slide to a template deck | manual (`read:append-ceiling`) |
| `backlog-ledger.mjs` | Gate | Validates and queries `docs/backlog.yml` | `verify`, `check:static` |
| `browser-harness-server.mjs` | Library | Static server for the Playwright harness | started by `playwright.config.ts` |
| `bundle-size-ratchet.mjs` | Gate | Gzipped size of the browser entry's closure vs `bundle-size-budget.json` | `verify:full`, `check:package` |
| `byte-identity.mjs` | Diagnostic | Freeze/compare emitted bytes across a refactor | manual — see note below |
| `check-commit-msg.mjs` | Gate | Rejects a shell-mangled commit message | `commit-msg` hook |
| `coverage-gate.mjs` | Gate | Per-area coverage thresholds from `coverage-gates.json` | CI (`coverage` job) |
| `coverage-merge.mjs` | Library | Merges Node + browser coverage into one report | `coverage:gate` |
| `docs-api.mjs` | Generator | TypeDoc → markdown API pages | `docs:check`, `docs:prepare` |
| `docs-check.mjs` | Gate | Frontmatter, nav and link validation | `docs:build` (Docs workflow) |
| `docs-frontmatter.mjs` | Library | Frontmatter parsing shared by the `docs:*` scripts | — |
| `docs-index.mjs` | Generator | Rebuilds `docs/doc-index.md` | `docs:prepare` |
| `docs-init.mjs` | Generator | Scaffolds the docs kit; inert in this repo | manual (`docs:init`) |
| `docs-list.mjs` | Diagnostic | Lists docs with their `read_when` hints | manual (`docs:list`) |
| `docs-new.mjs` | Generator | Creates a new doc page from the template | manual (`docs:new`) |
| `ensure-dist.mjs` | Gate | `dist/` freshness guard; builds, or `--check` fails | almost every `pnpm run` |
| `extract-autofit-calibration.mjs` | Generator | Fixture decks → `autofit-calibration.json` | manual, on fixture change |
| `gen-autofit-cases.mjs` | Generator | The `*.cases.json` autofit manifests — **duplicated**, see below | manual, on fixture change |
| `gen-inspect-snapshot.mjs` | Generator | The inspect-surface snapshot | manual; asserted by a regression test |
| `generate-llms-docs.mjs` | Generator | `docs/public/llms*.txt` | `docs:prepare` |
| `install-hooks.mjs` | Library | Installs lefthook, skipping where it cannot | `prepare` |
| `measure-autofit-lo.py` | Generator | LibreOffice cross-measure (Windows, needs LO) — **duplicated**, see below | manual — the one non-Node script |
| `ooxml-version-probe.mjs` | Diagnostic | Validator error counts across Office versions | manual (`schema:versions`) |
| `pack-utils.mjs` | Library | `pnpm pack` helpers for the two package gates | — |
| `package-lint.mjs` | Gate | `publint` + `attw` on the packed tarball | `verify:full`, `check:package` |
| `package-smoke.mjs` | Gate | Installs the tarball and exercises every subpath | `verify:full`, `check:package` |
| `pptx-parts.mjs` | Library | Explode/diff `.pptx` packages | — |
| `powerpoint-com-smoke.mjs` | Gate | Opens decks in desktop PowerPoint over COM | manual, Windows only (`test:com`) |
| `raw-xml-ratchet.mjs` | Gate | Hand-built XML per file vs `raw-xml-budget.json` | `verify`, `check:static` |
| `read-blindness-census.mjs` | Diagnostic | Which OOXML the read model never looks at | manual (`read:census`) |
| `read-emit-edits.mjs` | Generator | Edited decks for the manual PowerPoint check | manual |
| `read-emit-roundtrip.mjs` | Generator | `load()`→`save()` decks for the manual PowerPoint check | manual |
| `script-roundtrip.mjs` | Gate | Generated script must rebuild the deck it came from | `verify:full`, CI |
| `script-utils.mjs` | Library | `ROOT`, `run()`, and the shared CLI front end | — |

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

### Two generators exist twice — unresolved

Found while writing this table, recorded here rather than fixed because which directory
owns fixture authoring is a call worth making deliberately:

- `scripts/measure-autofit-lo.py` and `test/read/fixtures/authoring/measure-lo.py` are
  **byte-identical**.
- `scripts/gen-autofit-cases.mjs` and `test/read/fixtures/authoring/gen-cases.mjs` are the
  same ~380-line generator, differing only in how each resolves the fixtures directory.
  Both write the same four `*.cases.json` files.

Nothing keeps either pair in step, so editing one copy silently leaves the other stale.
`test/read/fixtures/authoring/README.md` documents its copies as part of the authoring
workflow, alongside the `author-*.ps1` scripts; `docs/measured-text-fit.md`,
`test/read/fixtures/README.md` and `extract-autofit-calibration.mjs` point at the
`scripts/` copies — and that last file cites *both* paths in different comments. Pick one
home, delete the other, and update the four references.

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
