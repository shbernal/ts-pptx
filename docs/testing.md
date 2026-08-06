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
pnpm run verify       # per-change loop: typechecks, backlog validation, raw-XML ratchet, all suites
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

### Reading a red number on the write side

The section above is about the read model, where the dominant question is whether
an input is even possible. The emitters fail differently: their input is whatever
a caller passes to the public builder, and every `src/gen/**` file sits behind one
or more *doors* — `addTable`, `addChart`, the `ts-pptx/measure` subpath — that
normalize before the emitter runs. So the write-side question is not "can this
input exist?" but **"which door reaches this, and is that door in scope?"**

Ask these in order before writing a case for a red emitter branch:

- **Is the file low on _statements_, not just branches?** That is a different
  signal, and a stronger one. A branch gap means arms are unexercised; a statement
  gap means a whole input shape or a whole outcome has never run. Both of the big
  finds in this area were statement gaps: `gen/define/hyperlinks.ts` sat at 48%
  statements because nothing anywhere put a hyperlink on a *table cell*, and
  `gen/define/zoom.ts` and `comment.ts` were low because every existing fixture fed
  them *valid* input, so the entire refusal half of each definer — the guards that
  drop an unresolvable target with a warning — had never executed. Estimate the
  phase from the statement number; the branch percentage will understate it.
- **On a small pure emitter, low statements means a caller never arrives.** The
  first question there is "which caller is supposed to reach this and doesn't?",
  not "which test am I missing?". `gen/drawingml/line.ts` at 65% statements looked
  untested and was in fact *unreachable*: four `define/` rebuilds dropped `cap`,
  `pattern` and `image` off the caller's object, so a documented public option was
  ignored library-wide and `line: { type: 'pattern' }` threw. The tell was a
  literal zero execution count on an arm a fixture demonstrably exercised.
- **A branch dead from one door can be live from another.** `text-fit.ts`'s newline
  handling and its `lineSpacingPct` / `spaceBefore` / `spaceAfter` defaults are
  genuinely unreachable from a deck, because `buildFitParagraphs` pre-splits every
  `\n` and always fills those fields — and fully reachable from `ts-pptx/measure`,
  a published subpath that exists precisely so a consumer can hand-build
  `FitParagraph[]`. Covering it *there* took the file to 100%. Same shape in
  `normalizeRuns`, unreachable from `addTable` (which normalizes) and reachable
  from `tableLayout()` (which does not). Enumerate the doors before calling an arm
  dead.
- **…and then check whether the other door is in scope.** The mirror case:
  `gen/table/autopage.ts`'s scalar-`colW` arms, `?? 0` fallbacks and no-rows guards
  are reachable only from `tableToSlides`, the browser-only path excluded from the
  report, while `addTable` resolves and defaults all of it first. Those really are
  dead. Both questions, not just the first one.
- **Is it a dev-only debug dump?** `verbose` is a documented `TableProps` option
  whose trace does `.toFixed()` arithmetic on props that may legitimately be
  percentage strings, and no test set it — seven uncovered functions and ~50
  uncovered statements in one file. Worth a case, because a dump that throws on a
  documented option is a real bug, but the test pins *"the trace survives its
  inputs"* and must say so in its header. It is not a substitute for the engine
  branches underneath.
- **Is the red arm dead code you could delete instead?** Two functions in
  `gen/drawingml/color.ts` were callable only from the browser path and measured on
  the Node chunk anyway; moving them into their sole caller (`gen/table/html-dom.ts`,
  already coverage-excluded) removed the false signal with no fence and no test.
  Thirteen `x = x || !x ? x : <default>` ternaries in `gen/define/chart.ts` were
  identity assignments whose alternative no value of `x` can reach; deleting them
  removed thirteen permanently-red branches and changed no emitted byte. Deleting a
  false signal beats fencing it, and beats testing it. Gate any such `src` edit on
  `byte-identity:baseline` / `byte-identity:check` (see AGENTS.md).

Two traps specific to this side, both of which have cost real time:

**A branch counter records that an operand was _evaluated_, not that it was true.**
The second arm of `!border.width || isNaN(border.width)` *looks* unreachable — `NaN`
is falsy, so the first arm should always catch it — and it is not: any truthy width
reaches the second operand and marks it. Two arms were written up as "unreachable by
construction" on that reasoning before a probe showed both green. Applies to every
`!x || isNaN(x)` and `!x || x.length` shape in the repo; probe before classifying.

**The src-import trap above applies to emitters too, and is easy to misread as a
gap.** `gen/slide/comments.ts` sits near 64% branches and is not untested —
`comments-xml.test.mjs` exercises every one of its red arms with stub slides, but it
imports from `src/`, so it can never move a `dist` number. That is the third
instance in the repo, after `html-dom.ts`'s helpers and `zoom-links.test.js`. Check
the import path of the tests that already name a file before calling its number a
gap.

Finally, **re-measure; do not reason from whatever is in `coverage/`.** A run started
with `--coverage.reportsDirectory` both writes to that directory *and* cleans it, so a
narrow probe left behind (or a later full run wiping `coverage/probe`) can leave a
stale `coverage-final.json` that reports *fewer* covered branches than a three-file
probe does. A subset beating the full run is the tell. Recorded per-file numbers drift
the same way: two files in this package were logged at 73 and 47 missed branches and
were actually at 83 and 23 when the work started.

## Raw-XML Ratchet

```bash
pnpm run raw-xml:check   # part of verify and check:static
pnpm run raw-xml:list    # every occurrence, with line numbers
pnpm run raw-xml:freeze  # rewrite the budget from source
```

`src/gen/oxml/el.ts` exists to end hand-concatenated OOXML — its header names
escaping, attribute-order and child-sequence bugs as the motivation. Most of
`src/gen/` now builds through it; the chart emitters do not, so "no raw XML
anywhere" cannot be turned on. `scripts/raw-xml-budget.json` therefore freezes a
per-file count of XML tag delimiters (`<ns:name`, `</ns:name`) appearing in string
and template literals under `src/`. A file may go down or vanish; it may never go
up, and a file absent from the budget must be at zero.

The check also fails when a count goes **down** — re-freeze in the same commit, so
the budget never accumulates slack it could later hide a regression behind.

Two exemptions. `src/gen/oxml/` itself, because emitting those delimiters is its
job; and a literal handed straight to `warn`, `notes.note`, or `new *Error`,
because a diagnostic that names the element it is about is prose, and prose forced
to dodge a gate gets written worse. The scan walks the TypeScript AST rather than
the file text, so an `<a:bodyPr>` in a doc comment was never a finding to begin
with.

This is a ratchet, not a correctness check: it says nothing about whether the XML
is right, only that the amount built by hand is not growing. Correctness is
[schema validation](#ooxml-schema-validation)'s job, and byte-stability during a
migration is the byte-identity harness's.

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

### The conformance target is pinned

Everything validates at **`Microsoft365`**, pinned as `FILE_FORMAT` in
`test/validator.js` and passed explicitly on every CLI invocation.

This is worth stating because two upstream defaults disagree: the Open XML SDK's
`new OpenXmlValidator()` defaults to `Office2007`, while the OOXMLValidatorCLI
wrapper defaults to `Microsoft365`. Passing nothing left the project's conformance
bar owned by the wrapper, where a bump to
`tools/ooxml-validator/version.json` could have moved it without a line changing
here.

`Microsoft365` is also the strongest available setting, not merely the newest. The
per-version schemas differ in how much markup they **model**, not in what they
accept — an older version skips markup it has never heard of rather than rejecting
it. Two consequences:

- Error count is monotonically non-decreasing in version, so validating below
  `Microsoft365` can only lose coverage.
- **Version-clean does not mean version-compatible.** A chartEx (Office 2016) deck
  reports zero errors at `Office2007` because that schema set cannot see the
  markup. Do not read a clean low-version run as "opens in Office 2007" — the
  validator cannot answer that question, and `mc:Choice Requires=` decisions must
  still be made against `[MS-PPTX]` and real PowerPoint.

### The validator does not descend into `mc:Choice`

**Nothing inside an `<mc:Choice>` is validated at all — only the `<mc:Fallback>`
branch is.** This is not a version-coverage effect, and raising `FILE_FORMAT` does
not help. Measured at `Microsoft365` against a deck containing an `am3d:model3d`
graphic frame, one mutation at a time:

| Mutation | Errors |
| --- | --- |
| bogus attribute on a `<p:sp>` outside the `mc:AlternateContent` | 1 |
| bogus attribute on `<p:blipFill>` **inside `mc:Fallback`** | 1 |
| bogus attribute on `<p:xfrm>` inside `mc:Choice` (plain DrawingML!) | **0** |
| bogus attribute on `<am3d:model3d>` | **0** |
| unknown child element inside `<am3d:model3d>` | **0** |
| `<am3d:camera>` moved out of document order | **0** |
| non-numeric `<am3d:perspective fov="wide"/>` | **0** |
| `<am3d:model3d>` with its **required** `r:embed` deleted | **0** |

The SDK does model `am3d` (Office2019+); it simply never reaches it. Since the
`mc:Choice` branch is skipped wholesale, this applies equally to the **zoom**
(`p:graphicFrame` in the 2016 zoom namespaces) and **OLE** (`p:oleObj`) emitters —
their real payloads are unvalidated too, and only their fallback pictures are checked.

So for any construct that lives in an `mc:Choice`, a green `test:schema` run says
nothing about it. Cover it the way `model3d` is covered:

1. diff the emitted subtree byte-for-byte against a PowerPoint-authored fixture
   (`test/read/model3d-roundtrip.test.js`), and
2. run `pnpm run test:com`, which opens the deck in the real application.

And for anything that *renders*, the COM read-back is not sufficient either — see
the desktop-check section below.

### Version coverage probe

```bash
pnpm run schema:versions              # built-in fixtures
pnpm run schema:versions --file d.pptx  # any deck
```

Validates across all seven accepted versions and prints the coverage profile:

```
fixture                              O2007 O2010 O2013 O2016 O2019 O2021  M365
base (plain text slide)                  0     0     0     0     0     0     0
classic bar chart (2007 feature)         0     0     0     0     0     0     0
chartEx pareto (2016 feature)            0     0     0     4     4     4     4
core-construct corruption (control)      1     1     1     1     1     1     1
```

Two uses: re-verifying monotonicity after a validator bump (the script exits
non-zero if any row *decreases*, which would break the premise the pin rests on),
and dating a known divergence to the schema generation that introduced it — row 3
locates chartEx at Office2016, and those 4 errors are the tolerated
`cx:axisId` divergence documented in `test/schema-cases.js`.

The last row is a control: a deliberately corrupted `<p:sp>` attribute that every
schema generation catches. Without it an all-zero table is indistinguishable from a
validator that silently stopped running.

Not part of `verify` — seven validator spawns per fixture, and it asserts nothing
about emitted markup that `test:schema` does not already assert at `Microsoft365`.

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

Fixtures authored here with desktop PowerPoint COM keep their recipe in
`test/read/fixtures/authoring/` (see that directory's README). Land the recipe there
rather than leaving it in `.tmp/`, which is gitignored — otherwise the fixture becomes
unreproducible on the next clean checkout. The same directory holds the scripts that
derive the committed `*.oracle.json` / `*.cases.json` sidecars; each regenerates its
sidecar byte-for-byte after a Prettier pass, which is how you check that a recipe still
matches what it claims to produce.

## Converter And Read-Coverage Harnesses

Three runnable measurement tools back the `ts-pptx/script` subsystem. Their
suites are already inside `pnpm run verify`; run them directly to iterate or to
point them at your own decks.

```bash
pnpm run script:roundtrip                     # deck → script → run it → deck → diff the two IRs
pnpm run script:roundtrip -- --tier a         # the standalone printer instead of template-anchored
pnpm run read:census                          # QNames present in a deck that no read accessor names
pnpm run read:append-ceiling                  # what survives fromTemplate + appendSlides
```

All three take `--json`, so a test can assert on them rather than re-derive the
numbers. `script:roundtrip` and `read:census` also take `--fixture` and `--dir`,
the latter so `pptx-bank/` decks can be measured the moment they land;
`read:census` adds `--all` (include layouts, masters, theme and notes) and
`append-ceiling` takes `--template <path>` instead.

`script:roundtrip` gates on **undeclared** differences: the printer's fidelity
notes are the exclusion list, so a difference no note predicted is a defect.
Read a clean run precisely — both IRs come from the same reader, so it detects
*asymmetry* and not loss in general. `read:census` is what measures the other
half, and it under-reports by construction (it counts element names appearing in
comments and error strings as read), so a listed element is a real gap while an
absent one is not proof of coverage. Full contract in
[PPTX To Script](reference/pptx-to-script.md).

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
pnpm run check:package   # package:lint + test:package
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
exists for. The `static`, `test` and `browser` jobs remain Linux-only: they are
platform-independent, and the validator installer is a bash script.

That narrows, but does not remove, the gap: a Windows-only break outside the
package scripts is still invisible to CI.

Note that the Windows leg has never actually executed — `ci.yml` has not run at
all yet. Treat it as the highest-risk part of the first push; see "Common
Commands" in [the development guide](./development.md) for what to watch.

## Browser Lane

```bash
pnpm run test:browser   # ensure-dist, build demos/vite-demo, then Playwright
```

Everything above this section runs under Node. `dist/browser.js` and its runtime
adapter (`src/runtime/browser.ts`) cannot: they call `fetch`, `FileReader`,
`<canvas>`, `URL.createObjectURL` and click a synthetic `<a download>`. This lane
is the only thing that executes them.

It is **not** part of `verify` or `verify:full`, deliberately — it needs a
~120 MB Chromium download, and putting that in the per-change loop would tax
every iteration for a surface that changes rarely. Install the browser once:

```bash
pnpm exec playwright install chromium
```

Config is `playwright.config.ts` (root); specs are `test/browser/*.spec.mjs`.
Vitest excludes `test/browser/**` by directory, so the two harnesses never
collect each other's files. In CI it is the `browser` job in `ci.yml`.

The fixture is `demos/vite-demo`, driven through a `vite preview` server. Two
assertions run against one deck build:

| Spec | Claim |
|---|---|
| `deck-download.spec.mjs` | the object-URL download is a real OPC package — read back with **jszip**, an implementation independent of the `fflate` the library writes with |
| `cross-runtime-bytes.spec.mjs` | the browser-built deck is **byte-identical** to the Node-built one, all 113 parts |

The second is the one worth the lane. `demos/vite-demo` imports the same showcase
module `pnpm demos:build quarterly-review` runs, and `src/zip.ts` pins
`FIXED_MTIME`, so the two packages are directly comparable — one diff asserts that
every serializer, the zip writer, part ordering and relationship numbering are
runtime-invariant. A runtime-dependent code path anywhere in `src/gen/` surfaces
here as a named part.

That comparison is the byte-identity gate's, not a second one: both go through
`scripts/pptx-parts.mjs` (same explode, same normalizers, same diff). Keep it that
way. Two hand-rolled comparisons would drift, and they would drift silently — one
gate tolerating a difference the other still calls a regression, with nothing to
say which was right. The three normalized values are the same three as ever:
`core.xml` timestamps and the two `Math.random` GUIDs (`p14:section` ids,
`c16:uniqueId`).

What this lane does **not** cover, and must not be read as covering:

- **Live-DOM layout fidelity.** No assertion here depends on a rendered page —
  no `offsetWidth` after layout, no resolved cascade, no browser-chosen font. That
  remains out of active scope ([project target](project-target.md)), and *runtime
  support* and *layout fidelity* are separate claims that should stay separate.
- **`loadMedia`, `createSvgPngPreview`, `loadFontData`.** The demo builds
  `quarterly-review`, which draws every asset rather than loading one, so the deck
  never crosses those three adapter functions. That is also why the byte
  comparison converges — it never has to reconcile Node's raw base64 with the
  browser's `FileReader` data URI. They stay uncovered, and `vitest.config.ts`
  still excludes `dist/browser*.js` from coverage.
- **Engines other than Chromium.** A deliberate decision, recorded in
  `playwright.config.ts`: the APIs in play are uncontroversial across engines, and
  a matrix would cost CI time for a divergence nobody has observed. Add Firefox or
  WebKit when something concrete surfaces.

## Demos Are Not Tests

The showcase decks have no test role. There is no demo smoke command, no
verification aggregate builds one, and a broken showcase fails nothing.

The one exception is `demos/vite-demo`, which the browser lane above uses as its
fixture and CI therefore builds. It is a fixture, not a showcase-with-assertions:
nothing checks how the page *looks*, only that the deck it builds is the right
bytes. (The byte-identity harness likewise *builds* the showcase decks without
asserting anything about them — a showcase that throws simply takes the harness
down with it.)

The test role used to belong to `scripts/demo-smoke.mjs`, which generated one deck
from `demos/node` and ran `vite build`. Both signals it produced are now covered
directly, and more precisely: `test:package` imports all nine export subpaths out
of an installed tarball and forces the `browser` condition, `package:lint`
validates types resolution with attw, and the browser lane puts a real bundler
(Vite/Rolldown) in front of the package and then *runs what it emitted*. That last
one closes a gap this section previously recorded as accepted — nothing proved
Rollup/esbuild could resolve and tree-shake the runtime entry. Something does now,
though only for the `browser` condition; the `node` entry still has no bundler
check, and if that ever bites, the fix belongs in `scripts/package-smoke.mjs` as a
bundler step against the installed tarball, not in a resurrected demo smoke.

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

### The object model is not a render oracle

For a question of the form *"does PowerPoint honour this construct?"*, reading the
answer back out of the COM object model does not answer it. The object model reports
the **resolved model** — what the file says, as PowerPoint parsed it. The renderer is
free to disagree, and when a construct is unimplemented it does exactly that: the
property round-trips perfectly and nothing paints.

This is not hypothetical. Custom table styles were designed, implemented, shipped and
then removed, and a COM read-back agreed with them the whole way: `Table.Style.Name`,
`Table.Style.Id`, `Cell().Shape.Fill.ForeColor.RGB` and `Cell().Borders()` all reported
the custom style's own values on decks that render completely unstyled — a black
hairline grid on white. The read-back was accurate and useless.

Export the slide and read pixels instead:

```powershell
$slide.Export("$PWD\slide1.png", "PNG")   # then compare pixels, not properties
```

A rendered pair beats a rendered single: hold everything constant but the one variable
(the same deck under a built-in GUID vs. a custom one; a PowerPoint-authored fixture
with one identifier rewritten, bytes otherwise identical). That is what turns "it looks
wrong" into "the gallery is consulted and the package part is not."

Applies to `pnpm run test:com` too, which asserts on shape state read back over COM. It
is the right tool for *package* health — a deck PowerPoint reports as corrupt, an
`hlinkClick` that resolves to the wrong `PpActionType` — and the wrong tool for whether
a construct is painted. See [tables.md → Table styles](tables.md#table-styles) for the
worked case.

A second, sharper demonstration came from the 3D-model work, and it is why the
`model3d` leg of `test:com` exports a PNG rather than stopping at the read-back. Take a
working model deck and replace the `.glb` payload with garbage, changing nothing else.
PowerPoint opens it without complaint, enumerates the shape as `Shape.Type = 30`
(`msoShape3DModel`), and reports `Model3DFormat.CameraPositionZ = 2.2630334` — the exact
value the good deck reports. The exported slide contains **zero** drawn pixels. Every
property the object model exposes was correct about a slide that renders nothing.

That leg therefore asserts on pixels, with the preview picture deliberately set to solid
magenta so the two distinct failures separate cleanly: magenta in the frame means
PowerPoint fell back to the picture, and a blank frame means the payload never
rasterized. Both absent is the only passing state.
