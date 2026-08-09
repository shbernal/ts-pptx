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
pnpm run verify       # per-change loop: typechecks, raw-XML ratchet, all suites
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
pnpm exec vitest run test/regression/api/object-identity.test.js   # one file
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

They are grouped into one directory per subject — `chart/`, `table/`, `text/`, `image/`,
`shape/`, `master-layout/`, `color-fill/`, `media/`, `slide-content/`, `html/`, `package/`,
and `api/` for the cross-cutting rest. The grouping follows the filename prefix, so a new
file's home is normally obvious from its name; when a file could sit in two groups, put it
with the subsystem whose *emission* it asserts on. Nothing in the tooling keys on the
directory — Vitest globs the tree — so a file can be moved between groups freely.

Paths inside a suite are relative to its group directory: `../../helpers.js`,
`../../../dist/node.js`, `../../read/fixtures/…`.

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

### File extensions

Every Vitest file is `*.test.js`. The package is `"type": "module"`, so `.js` is already ESM
and needs no `.mjs` to say so. A subset used to carry `.mjs` to mark tests that drive built
entry points with Vitest's `describe`/`test` API rather than the shared
`defineRegressionSuite()` harness, but the suffix had no functional effect — Vitest resolved
and ran both identically — so the distinction cost a paragraph of explanation and bought
nothing a reader could rely on. Which harness a file uses is visible in the file.

The one exception is the browser lane: `test/browser/*.spec.mjs` are **Playwright** specs, not
Vitest ones, and are matched by name in `playwright.config.ts` (and excluded from Vitest's
`include`). There the different extension marks a genuinely different runner.

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

The four numbers in `vitest.config.ts` are the **Node suite's floor**. The gate
the repo is judged on is the merged one below.

### Merged coverage

```bash
pnpm run test:coverage    # the Node suite  -> coverage/coverage-final.json
pnpm run test:browser     # the browser lane -> .tmp/browser-coverage/
pnpm run coverage:gate    # merge both, then check scripts/coverage-gates.json
```

The Node suite cannot execute `src/runtime/browser.ts` — it needs `fetch`,
`FileReader` and a canvas — but that file is in its denominator, correctly, since
nothing of this repo's own is excluded from the report any more. A denominator
with no collector for part of it understates the truth: dropping the exclusion
took functions 98.33 → 97.35 while tested-ness went *up*. The fix is to give that
part a collector, not to hide it again and not to move the gate.

`scripts/coverage-merge.mjs` does that, on one rule:

> **The Node report defines the shape. The browser lane contributes hits.**

Both sides remap V8 coverage with the same `ast-v8-to-istanbul` version Vitest
itself uses (pinned as a devDependency for exactly that reason), and the browser
side's file coverage is then projected onto the Node report's own statement,
function and branch maps by source location. So the merged denominator is
*identical* to the Node report's and the merged percentage is directly comparable
to it; only the numerator can move, and only upward.

About 1% of browser locations do not line up, because `dist/node.js` and the
`dist/*.js` chunks the browser loads are different bundles of the same source and
a few mappings resolve a column differently. Those hits are dropped, not
relocated — the count prints on every run, and past 5% the run fails instead,
because at that point the two lanes are looking at different builds.

### The point of slack, as a gate

`scripts/coverage-gate.mjs` checks the merged report against
`scripts/coverage-gates.json` and fails two ways:

| Failure | Meaning | Fix |
|---|---|---|
| below the notch | coverage regressed past the gate | cover it, or explain what changed |
| inside the point of slack | still above the notch, but by less than 1.00 | coverage has to come back up |

The second one is the rule every threshold in this repo was already set by — a
notch always sits at least a full point below its measured number — and it used
to live only in prose. Prose does not fail a build, so when the exclusion drop
took `functions` to 0.35 of slack and `lines` to 0.67, an acceptance criterion of
"thresholds still pass" was satisfied by a state the doctrine forbids. One was
noticed; the other sat in a stale comment. Encoding it means that cannot recur.

The rule is held against the merged report and **not** against
`vitest.config.ts`'s numbers, deliberately: demanding a point of slack on a
report whose denominator includes code its collector cannot reach would leave
only two ways to comply, and both — lowering the notch, re-hiding the file — are
what the doctrine exists to prevent.

In CI this is the `coverage` job, which runs after `test` and `browser` and
consumes their artifacts (including the browser job's `dist/`, so the merge reads
the exact bundles the browser ran).

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
  **Cover this one.** Either promote a fixture from a real-world deck or synthesize the
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
`comments-xml.test.js` exercises every one of its red arms with stub slides, but it
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
living under `test/`; the runner `test/schema-validation.test.js` consumes it).

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

### Working against real-world decks

The committed fixtures are construct-targeted and minimal by design. To probe
OOXML structures or reproduce read/round-trip behaviour against decks that are
messier and larger, point the measurement harnesses at a directory of your own
with `--dir` (`script:roundtrip` and `read:census` both take it, absolute paths
included).

Keep that directory **outside the repo**. Real decks are routinely large,
copyrighted, or client-confidential, and the working tree is the one place they
should not sit: a gitignore rule is a single edit away from not protecting them.
If you do want one in-repo, `.tmp/` is already ignored — but it is output
scratch, so treat anything you leave there as disposable.

The automated suites must only point at `test/read/fixtures/`, since no other
checkout or CI run will have your decks. When one of them proves a good minimal,
license-clean regression case, **promote** it: copy it into
`test/read/fixtures/`, add it to that directory's provenance table + SHA-256
list and purpose notes, and wire it into the harness (`FIXTURES` in
`test/read/roundtrip.test.js`). The `mixed.pptx` fixture arrived this way, to
cover connectors, nested groups, charts, and SmartArt that the vendored
fixtures lacked.

Fixtures authored here with desktop PowerPoint COM keep their recipe in
`test/read/fixtures/authoring/` (see that directory's README). Land the recipe there
rather than leaving it in `.tmp/`, which is gitignored — otherwise the fixture becomes
unreproducible on the next clean checkout. The same directory holds the scripts that
derive the committed `*.oracle.json` / `*.cases.json` sidecars; each regenerates its
sidecar byte-for-byte after an oxfmt pass, which is how you check that a recipe still
matches what it claims to produce.

## Converter And Read-Coverage Harnesses

Four runnable measurement tools back the `ts-pptx/script` subsystem. Their
suites are already inside `pnpm run verify`; run them directly to iterate or to
point them at your own decks.

```bash
pnpm run script:roundtrip                     # deck → script → run it → deck → diff the two IRs
pnpm run script:roundtrip -- --tier a         # the standalone printer instead of template-anchored
pnpm run script:census                        # how many decks raise each fidelity note, per tier
pnpm run read:census                          # QNames present in a deck that no read accessor names
pnpm run read:append-ceiling                  # what survives fromTemplate + appendSlides
```

All four take `--json`, so a test can assert on them rather than re-derive the
numbers. All but `append-ceiling` take `--dir`, so a corpus of your own decks can
be measured in place; `script:roundtrip` and `read:census` add `--fixture`,
`read:census` adds `--all` (include layouts, masters, theme and notes),
`script:census` adds `--names <count>` (name the decks behind the long tail), and
`append-ceiling` takes `--template <path>` instead.

`script:roundtrip` gates on **undeclared** differences: the printer's fidelity
notes are the exclusion list, so a difference no note predicted is a defect.
Read a clean run precisely — both IRs come from the same reader, so it detects
*asymmetry* and not loss in general. `read:census` is what measures the other
half, and it under-reports by construction (it counts element names appearing in
comments and error strings as read), so a listed element is a real gap while an
absent one is not proof of coverage.

`script:census` gates on nothing — it counts. The round trip cannot tell a note
that excuses a difference from one that never fires, so the per-construct numbers
the reference publishes drift silently as reader gaps close and fixtures land;
this is what re-measures them. Full contract in
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

That leg does execute, and is green on every CI run to date — including the
`workflow_call` gate inside `publish.yml`, so the Windows path is exercised on
every release rather than merely configured.

## Browser Lane

```bash
pnpm run test:browser   # ensure-dist, build demos/vite-demo, then Playwright
```

Everything above this section runs under Node. `dist/browser.js` and its runtime
adapter (`src/runtime/browser.ts`) cannot: they call `fetch`, `FileReader`,
`<canvas>`, `URL.createObjectURL` and click a synthetic `<a download>`. This lane
is the only thing that executes them — all four adapter functions, not just the
download path.

It is **not** part of `verify` or `verify:full`, deliberately — it needs a
~120 MB Chromium download, and putting that in the per-change loop would tax
every iteration for a surface that changes rarely. Install the browser once:

```bash
pnpm exec playwright install chromium
```

Config is `playwright.config.ts` (root); specs are `test/browser/*.spec.mjs`.
Vitest excludes `test/browser/**` by directory, so the two harnesses never
collect each other's files. In CI it is the `browser` job in `ci.yml`.

> **Drive the lane through `pnpm run test:browser`, not `pnpm exec playwright
> test`.** Only the package script runs `scripts/ensure-dist.mjs`, so the bare
> invocation serves Chromium whatever `dist/` was last built. That is the same
> stale-`dist/` trap as the Vitest inner loop above, with a sharper edge: the
> bare form is what you reach for when **sensitivity-checking** a new assertion,
> and a sabotaged `src/` that never got bundled leaves every spec green. A
> sensitivity check that cannot fail is indistinguishable from one that passed,
> and this has cost real time twice. Rebuild first, or go through the script.

### Three fixtures, three Playwright projects

They answer different questions, and none can answer another's:

| Project | Fixture | What only it can prove |
|---|---|---|
| `demo` | `demos/vite-demo` behind `vite preview` | the **bundled** path a real consumer takes — Vite resolving the `browser` export condition, Rollup tree-shaking it |
| `runtime-adapter` | `test/browser/harness/index.html` behind `scripts/browser-harness-server.mjs` | the shipped `dist/browser.js` loading **unbundled**, and the adapter loaders the demo cannot reach |
| `html-table` | `test/browser/harness/table.html`, same server | `tableToSlides` reading a **non-zero `offsetWidth`** — the one width basis no Node DOM can produce — and the end-to-end conversions that basis feeds |

Each project matches its specs by filename prefix (`deck-*`/`cross-runtime-*`,
`adapter-*`, `table-*`), and none of them matches by exclusion. That is deliberate:
`demo` was once spelled as "everything except `adapter-*`", which silently meant
"everything not yet invented" — the next prefix added would have run a second time
against the demo's `baseURL` and failed for reasons unrelated to what it tests.

The demo deck (`quarterly-review`) draws every asset it shows, so it never asks
the runtime to load one — which is why it cannot cover three of the four adapter
functions, and why the harness exists. The harness serves the repo with its real
layout and loads `dist/browser.js` over a plain `<script type="module">`, so what
runs is the file that ships rather than a re-bundling of it. A `node:*` import
reaching the browser entry would fail the page outright.

That is not hypothetical: building the harness is what surfaced `opentype.js`
being a *dynamic* bare import inside the measure/fit chunk. Bundling had always
hidden it; an unbundled consumer needs it in an import map, and now
[the docs say so](runtime-and-package-support.md#using-the-browser-entry-without-a-bundler).

| Spec | Project | Claim |
|---|---|---|
| `deck-download.spec.mjs` | demo | the object-URL download is a real OPC package — read back with **jszip**, an implementation independent of the `fflate` the library writes with |
| `cross-runtime-bytes.spec.mjs` | demo | the browser-built deck is **byte-identical** to the Node-built one, all 113 parts |
| `adapter-media.spec.mjs` | runtime-adapter | `loadMedia` and `createSvgPngPreview`: a fetched raster image lands as the same bytes Node reads off disk *and* as the source file's; the `<canvas>` rasterizer emits a real PNG where Node stubs a placeholder; 404, undecodable-SVG and zero-dimension-SVG each fail with the right code |
| `adapter-fonts.spec.mjs` | runtime-adapter | `loadFontData`: a font fetched over HTTP bakes the same `fontScale` and embeds the same `/ppt/fonts/` bytes as one read off disk; a 404 rejects with `font/fetch-failed` |
| `adapter-coverage.spec.mjs` | runtime-adapter | all four adapter functions ran, and `dist/browser.js`'s executed share stayed above its floor |
| `table-widths.spec.mjs` | html-table | `tableToSlides` against a table a browser laid out: the **measured** arm of `pickColWidthBasis` drives the emitted grid, `data-pptx-width` still wins outright (including divided across a `colspan`), and Node falls back to the CSS basis on the same markup — a *different* proportion, not a coarser one, because the two bases measure different boxes |
| `table-autopage.spec.mjs` | html-table | a table too tall for one slide pages with **one row budget on every page**, carries every row across exactly once, and — the assertion no other lane can make — reaches the *same* pagination in Chromium as on a DOM that renders nothing |

`table-autopage.spec.mjs` is worth reading for what a browser lane is *for*. It was
written as the headless repro a report dismissed as out of scope had invited
(`gitbrent/PptxGenJS#1200`, `tableToSlides` auto-paging overflow), it reproduced, and the
bug it found was **DOM-free**: the pager dropped one row's cell margins at every page
break, so a continuation slide accepted a row it had no room for. The browser's
contribution was the cross-runtime assertion — proving the report was never about a
rendered page, and moving it out of the browser bucket rather than deeper into it. Fixed in `src/gen/table/autopage.ts`; the regression that guards it is DOM-free too
(`test/regression/table/table-autopage-continuation-budget.test.js`). The triage rule that came
out of it — ask what the browser actually supplies to a code path before accepting a
report as a layout report — is stated with the scope line in
[project target](project-target.md).

The deck definitions the adapter specs use live in `test/browser/harness/decks.mjs`
and are built **twice** — once in Chromium, once in Node — from that one
definition. Writing them out on both sides would mean a divergence in the fixture
reading as a divergence in the runtime.

`cross-runtime-bytes` is the one worth the lane. `demos/vite-demo` imports the same showcase
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

### Where the coverage number for this lane lives

Two places, answering two questions.

**Per function, in the lane itself.** `adapter-coverage.spec.mjs` asserts on
Chromium's own V8 block coverage of the `dist/browser.js` script, collected across
every harness scenario. Two assertions, red for different reasons — every adapter
function was entered, and the file's executed share stayed above a floor (measured
92.74%, floor 90). The first catches a function losing its only test; the second
catches a function that is still entered but whose interesting arms are not. A
merged percentage can express neither, which is why this stays.

What keeps it off 100 is named in that spec: the missing-2d-context arm and
`FileReader.onerror`. Both are unreachable in a working browser — getting to them
means stubbing a DOM constructor, which asserts about the stub. `tableToSlides`
used to be on that list as well; it is covered now, by the `html-table` project,
which does not and should not move this number — the measurement here is of the
*adapter* harness page, whose DOM has no table in it.

**As a percentage, in the merged report** — see [Merged coverage](#merged-coverage).
Every spec in the `runtime-adapter` and `html-table` projects contributes its raw
V8 coverage to `.tmp/browser-coverage/` (the fixture in `test/browser/fixtures.mjs`,
auto-use, so a new spec contributes by existing), and `scripts/coverage-merge.mjs`
folds it into the Node report.

What this lane does **not** cover, and must not be read as covering:

- **Live-DOM layout fidelity.** The `html-table` project does depend on a rendered
  page, and the distinction between what it proves and what it does not is the
  whole point. It asserts that a real `offsetWidth` is *taken and honoured* — that
  the measured arm of `pickColWidthBasis` runs and the emitted grid is proportional
  to it. It asserts nothing about whether Chromium's numbers are the right numbers,
  or whether another engine would produce them. The auto-paging spec draws the same
  line vertically: it asserts that pages of identical rows get identical row budgets
  — arithmetic the pager owes itself — never that an estimated row height is the
  height PowerPoint will draw. That second claim is layout
  fidelity, it has no oracle, and it remains out of active scope
  ([project target](project-target.md)). *Runtime support* and *layout fidelity*
  are separate claims and must stay separate: a layout difference between two
  browsers is not a defect in this package; a `.pptx` a browser builds differently
  from Node is.
- **Engines other than Chromium.** A deliberate decision, written down in
  [Runtime And Package Support](runtime-and-package-support.md#which-browsers-the-lane-runs)
  so it is not re-opened every time CI time is discussed: the APIs in play are
  uncontroversial across engines, and a matrix would cost CI time to re-answer a
  question nothing has asked. Add Firefox or WebKit when something concrete
  surfaces. (`adapter-coverage.spec.mjs` is Chromium-only by construction —
  `page.coverage` is a CDP feature — which is a consequence of that decision, not
  a reason for it.)

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
directly, and more precisely: `test:package` imports all ten export subpaths out
of an installed tarball and forces the `browser` condition, `package:lint`
validates types resolution with attw, and the browser lane puts a real bundler
(Vite/Rolldown) in front of the package and then *runs what it emitted*. That
closes a gap this section previously recorded as accepted — nothing proved
Rollup/esbuild could resolve and tree-shake the runtime entry.

### Bundling the package for Node

Both conditions are answered now. The `node` entry's half is `bundleForNode()` in
`scripts/package-smoke.mjs`, which esbuild-bundles the *installed tarball* with
`platform: 'node'` and then runs what it emitted. It is not redundant with the
export matrix next to it, because the two use different resolvers asking different
questions:

| | resolves | when |
|---|---|---|
| export matrix | `node`'s own resolver | at call time, off disk |
| bundler step | esbuild's, walking `exports` under `platform: 'node'` | at build time, statically |

A package can be perfectly importable and still be unbundlable — a dynamic bare
import is the canonical case, because Node just finds it and a bundler must
resolve it. That is not hypothetical: it is exactly what the browser harness hit
with `opentype.js`, and nothing had asked the same question of the `node` entry.

Three assertions, red for different reasons: it builds with **no warnings** (a
warning is a failure here — allow one by name if it ever must be, never mute the
channel); **nothing but a Node builtin stayed external**; and the emitted bundle
**runs and writes a real `.pptx`**. It runs against the npm *and* pnpm fixtures,
since pnpm's symlinked store is a different shape for a bundler to walk.

Two things worth knowing before editing it:

- **Builtins are tested with `isBuiltin`, not a `node:` prefix.** The prefix is a
  convention, not the rule — `fflate` imports `createRequire` from bare `module`,
  which a prefix test reads as an unresolvable specifier. That was this check's
  first finding, against itself.
- **The bundled subpath list is derived from `EXPORT_MATRIX`**, minus `/browser`
  (the browser lane owns that condition). Adding a subpath there gets it bundled;
  there is no second list to keep in sync.

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

#### "PowerPoint will never paint this" is a claim that needs render evidence

The claim *this construct is valid OOXML, we can emit it, and PowerPoint does not
implement it, so the markup can never affect the render* is worth making — it is what
stops the next person re-attempting the construct. But it may only be made on **render
evidence**. Schema reasoning does not establish it (the markup being correct is the
premise, not the finding), and a COM read-back does not either, for the reason this whole
section exists. Export a slide and compare pixels, ideally as a rendered pair, or the
claim is unproven.

Keep it apart from its two neighbours: the construct being *out of this project's scope*
is a decision about us and can be revisited when the target changes; the markup being
*invalid* is a defect we can fix. This one is a property of PowerPoint, and nothing we
do here moves it.

#### Where such a finding has to live

A finding of that shape has to end up in front of whoever is about to re-attempt the
construct, and nobody reads an issue tracker before writing an emitter. So it does not
go in a tracker: distil it into the doc that the feature's own workflow already sends
them to — [tables.md → Table styles](tables.md#table-styles) for the custom-table-style
case — and put any reusable *method* note (how the render evidence was obtained) here in
this section. An issue that merely records the negative result is filed and forgotten;
a paragraph in the feature's own doc is read at exactly the moment it matters.
