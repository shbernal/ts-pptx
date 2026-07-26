---
doc-schema-version: 1
title: "Testing Guide"
summary: "Regression, schema, package, demo, and manual verification commands."
read_when:
  - Choosing verification commands
  - Updating test scripts or package smoke checks
  - Changing emitted OOXML or package exports
  - Deciding whether an uncovered branch is worth a test
doc_type: "guide"
---

# Testing Guide

Use `pnpm` for repository scripts. The package declares Node.js `>=24`.

## Standard Validation

For source changes, run one of the two aggregates rather than composing a set by
hand:

```bash
pnpm run verify       # per-change loop: typechecks, backlog validation, all suites
pnpm run verify:full  # before pushing / at the package boundary: the above plus package + demos
```

They are defined in `package.json` and described in
[development](development.md#common-commands). Neither runs `lint` or
`format:check` — the git hooks own those.

For documentation-only changes, no automated test is required unless the docs
change package, build, or testing claims.

## Fast inner loop (edit → test)

The suite imports from `dist/`, **not** `src/`. A full `pnpm run build` (8 entry
bundles + `.d.ts`) is far too slow to sit in a one-assertion edit loop. For the
inner loop, run a `tsdown` watcher and a Vitest watcher in **two terminals**:

```bash
# terminal 1 — rebuild dist/ on every src/ edit (fast: no .d.ts)
pnpm run watch:dev

# terminal 2 — rerun tests on every dist/ (or test) change, no rebuild of its own
pnpm run test:watch
```

`watch:dev` uses `tsdown.dev.config.ts`, which drops the most expensive build
step (`.d.ts` emit) while still emitting every Node-side entry the suites import.
`test:watch` is a bare `vitest --watch`: it does no build of its own, because in
this loop `watch:dev` is what keeps `dist/` current.

> **If you run only the Vitest watcher**, `src/**` edits appear to have no effect
> — you are testing the last-built `dist/`. Start `watch:dev` alongside it. The
> one-shot scripts do not have this problem: they all begin with
> `scripts/ensure-dist.mjs`, which rebuilds a stale `dist/` before running.

### Running a single test

Once `dist/` is current (a `watch:dev` running, or after `pnpm run build`), drive
Vitest directly — these skip the build, so mind the stale-`dist/` caveat:

```bash
pnpm exec vitest run test/regression/object-identity.test.js   # one file
pnpm exec vitest run test/regression -t "content type default"  # by test name
```

You can also add `.only` to a `test(...)`/`describe(...)` while iterating. Bare
`pnpm exec vitest` does **not** rebuild — if you edited `src/**` without a running
watcher, rebuild first or you are testing stale code. Going through a package
script (`pnpm run test:unit`, etc.) avoids this entirely.

## Regression Suite Layout

Regression tests live in `test/regression/` and are organized by behavior, not
by historical bug number. File names should describe the contract being tested,
such as `object-identity.test.js`, `content-type-defaults.test.js`, or
`slide-master-placeholders.test.js`.

Each regression file calls `defineRegressionSuite()` from `test/helpers.js`.
The optional second argument records legacy provenance, for example
`legacy bug-21`, so old issue references remain traceable without making the
suite name opaque.

Prefer public API deck generation plus focused package/XML assertions:

- Use `build()` to create a presentation and inspect the generated package.
- Use `readEntry()` for specific package parts such as `ppt/slides/slide1.xml`.
- Use helper assertions such as `assertContentTypeDefault()`,
  `assertContentTypeOverride()`, `assertXmlOrder()`, and
  `assertNonVisualDrawingProperty()` when they match the behavior under test.
- Keep raw XML substring or regex assertions local and narrowly targeted when a
  helper would hide the OOXML detail being tested.

Add a regression test when a public API call must keep producing a specific
package part, relationship, OOXML element, attribute, or absence of generated
parts. Name the file after the behavior, and include bug or upstream issue
context in the suite metadata or test name only when it helps future triage.

### `.test.js` vs `.test.mjs`

Both extensions coexist under `test/regression/` (the bulk are `.test.js`; a
handful are `.test.mjs`). The package is `"type": "module"`, so `.js` is
already ESM and Vitest resolves and runs both identically — the suffix has **no
functional effect** here (no build, transform, or resolution difference), and
both can import from `dist/` or `src/`. The `.mjs` files are a historical subset
that made the ESM boundary explicit for tests exercising built entry points or
the measurement/runtime subsystems directly with Vitest's `describe`/`test`
API, rather than the shared `defineRegressionSuite()` harness. Prefer
`.test.js` for new regression files to match the majority; use `.mjs` only if
you have a specific reason to signal the ESM boundary.

## Coverage Gate

Coverage is enforced by thresholds in `vitest.config.ts`:

```bash
pnpm run test:coverage
```

The suite executes the **built** package (tests import from `dist/`), so v8
collects coverage on the bundled `dist/**` output and remaps line/branch data
back to `src/` via the sourcemaps `tsdown` emits. Instrumenting `src/**` instead
would report only ~8%, because almost nothing under `src/` is executed directly.

Thresholds (statements / branches / functions / lines) are pinned a notch below
the current measured numbers so an accidental regression fails CI without the
gate being flaky. **Ratchet them upward as coverage improves; never loosen them
to make a red build pass.**

Reading the report has one trap: a line shown **red in the dist report may
already be covered** by a `src/`-importing unit test. Some helpers (for example
the HTML-table `htmlBorderToProps` / `resolveHtmlColWidth`) are only reached in
the bundle through the browser-only path, which is fenced with `v8 ignore`, so
the dist bundle never executes them even though `src/`-level unit tests do.
Before adding a case for a red line, check whether an existing `src/`-importing
test already exercises it — otherwise the gate cannot credit the redundant test.

### Branches that are not worth covering

The read model parses OOXML defensively: almost every element lookup is written
`const x = firstChild(parent, 'a:foo'); return x ? … : null`, whether or not the
schema lets `a:foo` be absent. That style is right — a reader should not throw on
a deck PowerPoint accepts — but it means a chunk of the branch count is guards
against input no valid package can contain, and those branches are **deliberately
left uncovered**.

Coverage exists to show that behaviour is pinned. A test that hand-builds a
`p:sp` with no `p:nvSpPr`, or a theme part with no root element, pins nothing:
such a file is not a deck, no user can produce one, and the assertion would only
restate the guard. It moves the number without adding a guarantee, and leaves
behind a fixture that has to be maintained. Prefer honest coverage plus a written
reason over a green metric.

So before writing a test for a red branch, ask which of these it is:

- **Schema-impossible** — the child, attribute, or root is `minOccurs="1"` (check
  with `ooxml_children` / `ooxml_attributes`), or the relationship is required for
  the package to resolve. Leave it. Do not add a `v8 ignore` fence either: the
  fence is for code the bundle genuinely cannot reach, not for code that a valid
  input merely never reaches.
- **Unreachable by construction** — a re-check of something the caller already
  established (a `p:ph` looked up again on a shape that was filtered *for* having
  one). Leave it, and consider whether the guard is telling you the type could be
  narrower.
- **Schema-legal but unrepresented in the fixtures** — the input is a deck
  PowerPoint could write, and no committed fixture happens to be shaped that way.
  **Cover this one.** Either promote a fixture from `pptx-bank/` or synthesize the
  XML, following the approach in `test/read/slide-background-edges.test.js`
  (splice the variant into an authored deck, so the rest of the package is real).

`src/read/api/chrome.ts` is the worked example: it sits near 64% branches while
its statements, functions, and lines are at or near 100%. The header comment in
`test/read/chrome-read-edges.test.js` enumerates every remaining branch and the
content model that makes it impossible. The two that turned out to be legal input
(`p:sldLayoutIdLst` and `p:txBody`, both `minOccurs="0"`) are asserted at the foot
of that file — each patched into an authored master and run past the schema
validator, so "the input is legal" is checked rather than asserted. Extend that
note rather than re-deriving it if the number ever comes up again.

## OOXML Schema Validation

Install the validator once:

```bash
./tools/ooxml-validator/install.sh
```

Run schema fixtures:

```bash
pnpm run test:schema
```

Use this path for emitted OOXML changes. Add or update focused fixtures in
`test/schema-cases.js` (a flat fixture data module — not a Vitest suite despite
living under `test/`; the runner `test/schema-validation.test.mjs` consumes it).

The fixtures run **concurrently** (`describe.concurrent`), which took the suite
from ~50s to ~10s and is what lets `verify` include it. Two consequences worth
knowing:

- Each concurrent fixture spawns its own `OOXMLValidatorCLI` process, and
  `test/read` spawns validators too, so the real process ceiling is
  workers × `maxConcurrency`. `maxConcurrency` is pinned in `vitest.config.ts`
  rather than left at the default, and is the first knob to lower if CI turns
  flaky or runs out of memory — re-serializing the suite is not the fix.
- `testTimeout` is raised well above Vitest's 5s default for the same reason:
  under concurrency a fixture's wall-clock time mostly measures how long it
  queued for CPU, not how long it worked.

A `beforeAll` validates one minimal deck serially before the concurrent fixtures
start. `OOXMLValidatorCLI` is a .NET single-file app that self-extracts on first
run, and firing many processes at a cold extract directory can race.

## Read/Round-Trip Suite (`ts-pptx/read`)

The lossless read/edit subsystem (`src/read/`) has its own harness:

```bash
pnpm run test:read
```

It runs `test/read/roundtrip.test.js` against real, PowerPoint-authored decks
in `test/read/fixtures/` (provenance in that directory's README): part-set
stability, per-part byte-identity for untouched parts, lazy-parse guarantees,
save idempotence, content-type/relationship resolution, the dirty
(mutate-and-reserialize) path, and schema validation of saved output. The
schema cases require the OOXML validator above. When it is missing they are
skipped locally — with an unmissable notice on stderr, because a green run that
skipped a few hundred schema assertions must not read as a complete one — and
they **fail hard under `CI`**, where installing the validator is part of the job.
The gate is `validatorAvailable()` in `test/validator.js`.

Changes under `src/read/` should run this suite; new read/edit capabilities
should extend it (and grow the fixture set) alongside the code.

### `pptx-bank/` — real-world deck corpus (uncommitted)

`pptx-bank/` at the repo root is an **uncommitted** bank of real-world
PowerPoint files for ad-hoc testing and verification: probing OOXML structures,
reproducing read/round-trip behaviour against decks far messier and larger than
the curated fixtures, and finding candidates worth promoting to a committed
fixture. It is gitignored (`/pptx-bank/*` with a `!README.md` negation), so you
can drop in arbitrary decks — including large, copyrighted, or client files —
with no risk of them entering Git history. See `pptx-bank/README.md`.

How it relates to the other deck locations:

| Location | Committed? | Role |
|---|---|---|
| `test/read/fixtures/` | yes — hash-pinned, provenance-tracked, license-clean | curated, minimal **inputs** the harness depends on (CI runs these) |
| `.tmp/` | no | generated **output** scratch (e.g. `pnpm run test:read:emit`) |
| `pptx-bank/` | no | free-form **input** corpus for exploration and manual verification |

The automated harness must only point at `test/read/fixtures/` — bank files are
not tracked, so other checkouts and CI will not have them. When a bank deck
proves a good minimal, license-clean regression case, **promote** it: copy it
into `test/read/fixtures/`, add it to that directory's provenance table +
SHA-256 list and purpose notes, and wire it into the harness
(`FIXTURES` in `test/read/roundtrip.test.js`). The `mixed.pptx` fixture was
promoted from the bank this way to cover connectors, nested groups, charts, and
SmartArt that the vendored fixtures lacked.

## Full Test Command

```bash
pnpm test
```

This is `vitest run` with no target list, so it runs **every** suite Vitest
discovers under `test/` — regression, read, schema, and tooling. There is no
maintained list of suite paths to keep in sync; adding a `test/**/*.test.js`
file is enough to get it run.

## Package Boundary Checks

```bash
pnpm run check:package   # package:lint + test:package + test:demos
```

`package:lint` runs package export/type validation. `test:package` creates a packed package with pnpm,
installs it with npm and pnpm, verifies that the ESM entries and declarations
are present, verifies that old generated artifacts are absent, runs an ESM
import smoke test, checks that the package has no CJS export condition, and
typechecks a minimal TypeScript consumer.

The TypeScript consumer fixture is generated inline by `scripts/package-smoke.mjs`
(`type-smoke.ts`). It is the only consumer of the public API that is not in
`test/`, so it does not move when the API does — a rename or a removed overload
in `src/` will not fail any unit test but *will* fail `test:package`. When you
change a public export, grep that fixture.

There is deliberately no separate `pnpm pack --dry-run` check. `package:lint`
already packs the tarball for real and then runs publint and
`@arethetypeswrong/cli` over it, so a dry-run that only asserted "pack exits 0"
added a build and a pack invocation for no signal. Do not re-add one.

Packing goes through `packPackage()` in `scripts/script-utils.mjs`, which passes
`--config.ignore-scripts=true` to skip the `prepack` rebuild — the callers have
already ensured `dist/` is current. Note the spelling: pnpm 11 rejects a plain
`--ignore-scripts` on `pack` and honours only the `--config.` form.

### Running scripts that spawn subprocesses

Every script subprocess goes through `run()` in `scripts/script-utils.mjs`. It
deliberately avoids a shell where it can, because Windows cannot exec the
`.cmd`/`.ps1` shims that package managers and `node_modules/.bin` entries ship
as: a bare name fails with `ENOENT`, and appending `.cmd` fails with `EINVAL`
(Node >=18.20/20.12 refuses to exec batch files without a shell, per the
CVE-2024-27980 hardening).

So `run()` resolves a bin one of three ways:

- **A local devDependency bin** — resolved to its JS entry via the declaring
  package's `bin` field and run on `process.execPath`. Bin name to package name
  is not derivable (`attw` lives in `@arethetypeswrong/cli`), so the mapping is
  the explicit `localBinPackages` table. **Add new local bins there**, otherwise
  they fall through to the shell path and emit Node 24's `DEP0190` warning.
- **An external package manager** (`pnpm`, `npm`) — `.cmd` plus `shell: true`,
  passed as a single pre-quoted command line, since `DEP0190` deprecates an args
  array alongside `shell: true`.
- **An absolute path** (e.g. `process.execPath`) — spawned directly.

### Which of these CI runs on Windows

`ci.yml`'s `package` job has an OS matrix and runs `check:package` on both
`ubuntu-latest` and `windows-latest`, so the `run()` behaviour above — the part
of the repo that is Windows-specific by design — is exercised on the platform it
exists for. The `static` and `test` jobs remain Linux-only: they are
platform-independent, and the validator installer is a bash script.

That narrows, but does not remove, the gap: a Windows-only break outside the
package and demo scripts is still invisible to CI.

## Demo Smoke Tests

Run both maintained demo smoke tests:

```bash
pnpm run test:demos
```

Run one target:

```bash
pnpm run test:demo:node
pnpm run test:demo:vite
```

The demo smoke command ensures `dist/` is current first, then runs the maintained
workspace demos with pnpm. The Node demo validates ESM package usage in a Node
application. The Vite demo validates a modern browser app path through React,
TypeScript, and Vite.

## Manual Visual Checks

Automated tests prove package shape and generated XML structure. Manual visual
checks are still useful for user-visible PowerPoint behavior:

1. Generate a small deck from the Node demo or Vite demo.
2. Open it in Microsoft PowerPoint when available.
3. Check import behavior in Keynote, LibreOffice Impress, or Google Slides when
   the change affects cross-app compatibility.
4. For browser download behavior, prefer `demos/vite-demo`.

Node demo decks are written to `demos/node/output/`, which is ignored by git.
Re-running a demo command replaces the previous deck with the same name.
