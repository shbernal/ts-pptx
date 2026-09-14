# AGENTS.md

## Repository expectations

- This repository builds ts-pptx, a JavaScript/TypeScript library that emits PowerPoint
  `.pptx` packages using OOXML.
- Use `pnpm` for repository scripts. The package declares Node `>=24`.
- Keep source changes focused in `src/` and tests in `test/`. Treat `dist/` as generated
  package artifacts unless the task explicitly asks for refreshed release outputs.
- Preserve unrelated dirty state. Do not revert user changes.
- Three trees are not the library and are easy to confuse. `docs/` is **content**: markdown
  under a frontmatter schema the docs kit validates. One subtree of it, `docs/contributing/`,
  is **repository-only**: its pages are checked like any other but are read on GitHub, never
  built into the site or listed in `llms.txt` (`repoOnly` in `docs/docs.json`). `www/` is the **site's application
  code**, the VitePress theme and the Vue components a page mounts, including the demos
  page that previews a deck via `pptx-html` (`www/README.md`). `demos/` is **clone-and-run
  only**: someone runs a script and gets a `.pptx`. Do not put an application in `docs/`,
  and do not re-grow a browser app under `demos/`. That one already existed once, and it
  is now a page of the site.

## Scope: Node-first (two out-of-active-scope domains)

- This project is **Node-first**: it runs and is tested without a browser or any
  office application. Two domains are out of *active* maintenance scope. Do not
  build features there, do not go hunting for fixes there, and never block other
  work on them. When a task lands in one, say so and treat it as out of scope,
  unless the user explicitly opts in. Neither is rejected on merit, and outside
  contributors are welcome to submit PRs. The maintainer is simply not driving them.
  See `docs/project-target.md` ("Out of active scope") for the full statement.
  - **Live-DOM / browser-layout features.** Anything whose answer comes from a
    *rendered* page: real `offsetWidth` after layout, the resolved cascade, fonts
    as the browser chose them. `tableToSlides()` is NOT in this category any more:
    it ships as a free function on `pptx-ts/html`, runs under Node with any DOM,
    and is covered against happy-dom (`test/regression/html/html-to-slides-node.test.js`).
    Only real *measurement* is out of scope there. Without a layout engine
    `offsetWidth` is `0`, so column widths fall back to computed CSS widths then to
    an equal split. Say "fall back", not "degrade": `offsetWidth` is the border box
    and computed `width` the content box, so the two bases can put a table's columns
    in different *proportions*, not merely at a coarser precision. Keep extracting
    the DOM-independent part of anything here into
    a pure helper and unit-testing it (pattern: `resolveHtmlColWidth`,
    `pickColWidthBasis`, `cssColorToHex`).
  - **Third-party office-suite interop quirks.** Breakage that only appears after
    a round-trip through another app (e.g. WPS copy/paste, then PowerPoint) when
    the generated package is itself valid OOXML. The supported bar is that output
    opens cleanly in Microsoft PowerPoint. Such an item only becomes in-scope with
    a repro pinning the defect to invalid OOXML the library itself emits.

## API evolution policy

- This project is maintained for our own use; there is no external backward-compat
  obligation. Prefer fixing root causes here rather than asking a downstream
  consumer to work around them. A fix in this public package helps every
  consumer.
- Breaking changes are acceptable and encouraged when they make the API clearer
  or safer. Do not block an improvement on reverse compatibility. When you make
  one, record it (with migration guidance and downstream impact) in `CHANGELOG.md`.
  When you only *propose* one, or want to track a not-yet-implemented candidate,
  open a GitHub issue (see Tracking Work below).
- Silent coercion of invalid input is a footgun, not a feature: prefer warning or
  failing on `NaN` / `undefined` / out-of-range values over emitting a degenerate
  result (e.g. a zero-size object).
- **When you give an option value a meaning it did not have, work out what its
  *absence* now means.** An emitter that defaults a missing option through a
  ternary has already assigned one of the states to silence, so correcting the
  explicit spelling moves the boundary without anyone editing the default. #9 gave
  `fill: { type: 'none' }` its `<a:noFill/>` back and, in the same stroke, left
  *inherit* with no spelling at all on that path, because omission already meant
  no-fill. So 3.1.0's own migration advice, "omit `fill` instead", pointed at the
  wrong state. #10 is that hole. A property with three states needs three
  spellings. "The caller said nothing" is not one of them: it is an *alias* for one
  of them, and which one is a decision to make deliberately and write down, not a
  fact to read off a ternary. Before calling such a fix complete, read the other
  arms of the branch you changed and ask what is now unreachable.
- Before adding, widening, or removing an escape hatch (raw XML, a passthrough
  string, direct DOM access), read the "Escape hatches" section of
  `docs/project-target.md`. It states the convenience-vs-guarantee rule and why
  the read path gets a deep raw hatch while the write path does not. That
  document is the whole reasoning. Do not re-derive it from anywhere else.

## OOXML and PowerPoint work

- Before changing emitted OOXML, read `docs/ooxml-agent-context.md`.
- Do not vendor full standards PDFs or large extracted specification text into this
  repository as agent context. Small, repo-specific notes with section references
  instead.
- Prefer executable evidence over prose alone. Inspect minimal PowerPoint-authored
  `.pptx` packages, compare package XML, add focused regression or schema fixtures.
- Some behaviour can only be judged against genuine PowerPoint output: a read accessor
  validated against real Office XML, or a write-side path whose target is "whatever
  PowerPoint authors". Where that fixture does not exist yet, do not implement against
  synthetic or round-tripped XML. Open a GitHub issue naming the construct the oracle
  has to contain, and stop until the fixture is authored. See
  `docs/evidence-and-fixtures.md`.

### MCP tool selection

Two MCP servers cover complementary parts of the OOXML/PowerPoint space. Work
through them in order before falling back to web search.

**Step 1: the `ooxml` MCP** (`mcp-server-ooxml`, source: the ECMA-376 / ISO 29500 XSDs)

Runs locally over stdio with the schema graph inside the package: no account, no
network, and deterministic. Use it for questions whose answer lives in the
*schema itself*:
- Element and complexType definitions (`ooxml_element`, `ooxml_type`)
- Legal child elements in schema order, with cardinality (`ooxml_children`)
- Attribute names, types, defaults, and required/optional (`ooxml_attributes`)
- The legal value space of a type, meaning facets, patterns and unions (`ooxml_values`),
  or just the enumeration (`ooxml_enum`)
- Namespace URIs and prefixes (`ooxml_namespace`)
- Whether a construct survives in Strict as well as Transitional
  (`ooxml_diff_profiles`)
- Turning an `ooxml-validate` diagnostic into "what would have been legal here"
  (`ooxml_explain`, passing the report's `id`, `description`, `xpath` and `partUri`)
- Finding a half-remembered name (`ooxml_search`, a **substring match on
  names**, not a semantic or full-text search)

It answers from the schema graph and nothing else. It has **no specification
prose, no PDFs, and no OPC part/content-type/relationship catalogue**, so
"which ECMA-376 section describes this rule" and "which content type belongs to
this `.pptx` part" are not questions it can take. For those, go to Step 2 and
then Step 3. It equally does not cover Microsoft-proprietary details such as
built-in style GUIDs, behavior differences between Office versions,
[MS-OE376] / [MS-PPTX] deviations, or Open XML SDK usage; those are Step 2.

**Step 2: the `microsoft_learn` MCP** (source: Microsoft Learn / Open Specifications)

If the `ooxml` MCP returns incomplete or no answer, always try this before web search.
Use it for:
- Microsoft Open Specifications ([MS-OE376], [MS-PPTX], [MS-OFFCRYPTO], …). These
  document how Office *implements* or *deviates from* ECMA-376, and they carry the
  Microsoft-proprietary enumerations: built-in table style GUIDs, preset shape
  adjustment ranges, behavior flags that are nowhere in the standard.
- PowerPoint-specific rendering behavior, repair heuristics, and version-gated features.
- Open XML SDK (`DocumentFormat.OpenXml`) API usage and samples.
- Azure / Microsoft 365 platform documentation.

Use `microsoft_docs_search` for a broad query first, then `microsoft_docs_fetch` on
a returned URL when you need the full page content.

**Step 3: web search (`WebSearch` / `WebFetch`)**

Only after both MCPs have been tried and the information is still missing or
ambiguous. Useful for community discoveries (e.g. undocumented GUIDs found by
reverse-engineering), third-party library behaviour, and content that postdates the
MCPs' corpora.

## Tracking work

- Not-yet-built work goes in a **GitHub issue**: a bug, a proposed API change, a
  missing PPTX behaviour, a fixture that has to be authored before a feature can
  be implemented. There is no local ledger. Do not add one, and do not re-add an
  upstream sync step (upstream tracking is retired).
- Three forms in `.github/ISSUE_TEMPLATE/` cover the usual shapes. **bug** takes
  wrong output, repair prompts, regressions and fidelity limits, and carries a
  severity plus a "possible fixes" slot. **api-gap** takes a missing accessor, or a
  property the write side authors that the read side cannot read back. A fixture
  that has to be authored first is a bug of severity *fidelity limit*. If neither
  fits, file a blank issue rather than bending one of them.

  The third form, **agent-report**, is not for work started here. It is where
  `InternalError`'s message sends an agent working in a *consumer* repo, and it is
  the only form that asks which class and code was thrown, and where an attached
  file came from. Do not route local work through it. It labels `agent-reported`, which
  records *where an issue came from* rather than what it is. It reaches feature requests
  and doc mismatches as well as defects, so triage still adds `bug` or `enhancement`
  after reading.
- The `skills/ts-pptx-upstream/` skill is the consumer-side half of that form. It
  ships in the tarball (`files` includes `skills`), so it is what
  `npx skills add ./node_modules/pptx-ts` installs. Keep it in step with
  the taxonomy and the forms when either changes. Everything under `.agents/skills/`
  is for working *on* this repo and is flagged `metadata.internal: true` so the same
  command does not offer it to a consumer.
- Work you implement on the spot needs no issue. Its record is the project's own
  commit history, its tests, and `CHANGELOG.md`.
- **Describe a downstream consumer's need ANONYMOUSLY.** Issues are public; the
  consumer is not. State the missing PPTX behaviour and how *any* consumer would
  reproduce it. Never the consumer's name, its file paths, deck or client names,
  or its content. See `docs/agent-development.md` for the checklist that moves
  such a need into the project.

## Verification

### The default loop

- **`pnpm run verify`** (~44s) is the per-iteration check: a `dist/` freshness guard →
  `check:core` (`typecheck` → `typecheck:scripts` → `typecheck:test` → `typecheck:site` →
  `raw-xml:check` → `path-refs:check` → `docs:check` → `comparison:check`) →
  the whole test suite (`vitest run`, which discovers every suite
  including schema). Run this instead of hand-composing four or five separate commands.
  Hand-composed sets come out slightly different every time and end up re-running the
  same suite twice.
- **`check:core` is the one list of cheap checks**, shared by `verify` and CI's
  `check:static` (which adds `lint`, `lint:chars` and `format:check`). The two used to
  copy it by hand, and `comparison:check` joined `verify` alone, so CI never ran it.
  `test/scripts/gate-parsers.test.js` fails if a check reaches `verify` without reaching
  `check:static`. Add a cheap check to `check:core`, not to either aggregate.
- **`pnpm run verify:full`** (~77s) before pushing or for a release/package-boundary
  change: everything in `verify`, plus `docs:build`, `script:roundtrip:all`,
  `package:lint`, `test:package`, `bundle-size:check` and `bundle-tier:check`. The split
  is only about cost. Those pack and install the tarball, build the production site, or
  run the whole read corpus twice; everything cheaper already lives in `verify`.
- **`docs:build` is in `verify:full`, not `verify`.** 19.7s of its 26.6s is
  `vitepress build`, a production static-site build that proves something about the
  *site*, not the library; `docs:check` keeps the docs themselves validated every
  iteration for ~3.5s. CI never depended on the loop tier for it. `docs.yml` builds the
  site on every pull request, so carrying it in `check:static` too just built it twice.
- Both aggregates are assembled by `scripts/run-steps.mjs`, which expands script names
  into their leaf commands and runs them in one process tree. `package.json` remains the
  single definition of every step. The runner only removes the 0.7 to 1.3s package-manager
  relaunch between them (`verify` was paying that 13 times). It prints a per-step
  breakdown when it passes, and `node scripts/run-steps.mjs --list verify` shows the
  expansion without running anything. Add a step by editing the name list in
  `package.json`, never by inlining its command. Inlining is how a second, drifting
  copy of a gate gets created.
- **The test suite sizes its own worker pool from free memory** (`vitest.config.ts`),
  not from the core count, because cores decide how fast it could run while memory
  decides whether the host survives it. On an idle machine this lands on the CPU bound
  and costs nothing; under memory pressure it scales down instead of driving the box
  into swap. `VITEST_MAX_WORKERS` pins it explicitly and is never clamped. Do not
  "fix" a slow run by raising `maxConcurrency`. Since validator batching landed, that
  knob no longer buys spawn parallelism. See docs/testing.md "Suite cost and the worker
  ceiling".
- **The suite runs `isolate: false`**, so one module registry is shared per worker rather
  than rebuilt per test file. `dist/` is over 1 MB of JS, and all 235 files were each
  re-evaluating it (3.5 to 4x on the `import` phase, 22 to 37% of wall clock). What that
  costs you: module-level state in a test helper is now shared with every other file in
  that worker. A cache wants that (`test/validator.js`'s batch queue and `corpus.js`'s
  `irFor` memo both got better for it); state carrying one test's *intent* does not.
  `test/setup-globals.js` resets `setDiagnosticHandler` after every test, and
  `sequence.shuffle.files` randomizes file order so an order dependence fails rather than
  hides. See docs/testing.md "One module registry per worker".
- **`pnpm run script:roundtrip:all`** (~25s, in `verify:full` and CI) is the script
  converter's gate: for every read fixture it prints a script, runs it, and diffs the
  result against the source with the printer's own fidelity notes as the exclusion list.
  A difference no note predicted is a defect. Both tiers run, tier B against the source
  deck as template and tier A with no template at all, because they gate different claims.
  It certifies "nothing the converter can see was lost", never "nothing was lost";
  `read:census` is what measures the second thing. It is also the **only** copy of that
  sweep: the two Vitest suites each carried a byte-equivalent one, so a `verify:full` ran
  the same 44 decks × 2 tiers twice. Those suites keep what the round trip rests on and
  cannot itself establish (the diff fails when perturbed, a note excuses only its own
  field, the canonicaliser is an equivalence). So `verify` alone does **not** cover the
  corpus round trip. Run `script:roundtrip:all` before pushing `src/script/` changes.
- **`pnpm run script:census`** (~2s, in no gate) counts how many fixtures raise each
  fidelity note, per tier. It gates nothing on purpose. It is the number
  `docs/reference/pptx-to-script.md` publishes, and the round trip cannot keep it true,
  because a note that excuses a difference and a note that never fires look identical to
  it. Run it after closing a reader gap, retiring a note, or landing a fixture, and
  refresh that page's tables in the same commit.
- Three of those are **ratchets**, and each fails on a change you did not intend as much
  as on one you did: `raw-xml:check` (in `verify`), `bundle-size:check` and
  `bundle-tier:check` (both in `verify:full`). A ratchet failure is not automatically a
  defect. It is a prompt to decide whether the number moved for a reason, then re-freeze
  in the same commit (`raw-xml:freeze`, `bundle-size:freeze`, `bundle-tier:freeze`) if it
  did. Do not re-freeze to make a gate quiet without knowing which change moved it.
- **The two size gates answer different questions and neither replaces the other.**
  `bundle-size:check` weighs the closure of each published entry without bundling, so its
  figure is an upper bound on what the package *ships* and does not move when code merely
  becomes unreachable. `bundle-tier:check` bundles three real consumer programs against
  `dist/` and lets a bundler shake them, so it is the only gate that can see reachability
  change. Expect the two numbers to disagree.
- Builds are never something you sequence by hand. Every gate starts with
  `scripts/ensure-dist.mjs`, which rebuilds only when `src/` or a build config is newer
  than `dist/` and is otherwise a ~0.1s no-op. Do not prefix anything with
  `pnpm run build &&`.
- `pnpm run build` passing is **not** evidence of type-correctness. tsdown's `.d.ts`
  pass does not typecheck, so a genuine type error still builds clean. `typecheck` is
  the only thing that catches it; never treat a green build as a substitute.

### Coverage: probe in the loop, gate before the commit

- **`pnpm run test:coverage` (~2min, observed 60 to 185s) is a commit gate, not a loop
  step.** It runs every test file, the schema suite included, which drives a .NET
  validator over every fixture. None of that answers the question you actually have,
  which is whether the test you just wrote reaches the line you wrote it for. Run it
  once, immediately before each commit, and never before an edit. A coverage session
  that runs the full gate after every edit spends more wall clock waiting on validators
  than on everything else combined.
- **`pnpm run coverage:probe <paths…>` is the loop step.** Same v8 instrumentation over
  the same `dist/` bundle, but scoped to the files you name and with the thresholds
  zeroed so a partial run cannot fail on a number it was never going to reach. One test
  file is ~15s; all of `test/read` is ~50s. It writes to `coverage/probe/` rather than
  `coverage/`, so a probe never overwrites the numbers the real gate produced.
- **Read the probe in one direction only.** A line the probe reports as *covered* is
  covered. That is conclusive, and it is the question the loop actually asks ("does my
  new case reach this branch?"). A line it reports as *uncovered* means only that the
  files you named do not reach it; another suite may. Never delete or rewrite a test on
  the strength of a probe miss. Confirm against the full gate first.
- Read specific lines out of `coverage/probe/coverage-final.json`; the per-file rollup
  is in `coverage-summary.json`. Both are the same shape the full gate writes, so a
  helper script works against either.
- **`test:coverage` is a floor, not the gate the repo is judged on.** Its denominator
  includes `src/runtime/browser.ts`, which no Node run can execute, so the number it
  prints understates the truth. The real gate is `pnpm run coverage:gate`, which merges
  the browser lane's V8 coverage into that report and additionally fails when a number
  clears its notch by less than a full point. It needs both lanes' output on disk, so
  running it locally means `test:coverage` **and** `test:browser` (~120 MB Chromium) have
  run first; otherwise leave it to CI's `coverage` job. See docs/testing.md
  "Merged coverage".

### Do not run these, the git hooks already own them

- **`format`, `format:check`, `lint`.** Pre-commit runs oxlint `--fix` then oxfmt
  `--write` over staged files and re-stages what they change (`stage_fixed: true`), and
  pre-push re-verifies the whole repo (`lefthook.yml`). An agent running `format:check`
  therefore cannot improve the committed result. It only burns a
  check→fix→re-check cycle on files that were going to be fixed anyway. This is not a
  lowered standard: the gate still runs, just not from your shell.
- **`lint:chars`.** Pre-commit scans the staged content of what you are committing and
  pre-push re-scans the repo (`lefthook.yml`), so running it by hand buys nothing. The
  one form worth invoking deliberately is `lint:chars:fix`, which is not in any hook on
  purpose: see the em-dash gate below.
- What the hooks do *not* cover, and `verify` therefore does: **tests** (no hook runs
  any), **`typecheck:test`**, and **`docs:check`** (pre-push checks `lint`,
  `lint:chars`, `format:check`, `typecheck`, `typecheck:scripts` and `typecheck:site`
  only).
  **`docs:build`** is covered by neither. It lives in `verify:full` and in CI's
  `docs.yml`, which runs on every pull request.

### Shell habit: do not pipe a verification command on its first run

Do not pipe a verification command through `| Select-Object -Last N` (or `tail`) the
first time you run it. Failures often explain themselves in a `beforeAll` or setup
message near the *top* of the output, and a tail filter scrolls exactly that off. The
concrete case: `test:schema` needs the OOXML oracle, and a machine that cannot obtain
one says so up front. Piping the first run turned one legible missing-validator error
into three blind 40s re-runs. Run bare first; filter only on a re-run, once you know
what you are looking for.

### Shell habit: reach a binary directly, keep `pnpm run` for scripts

`pnpm run X` and `pnpm exec X` add a flat ~0.7s to whatever they wrap. About 0.45s of
that is pnpm's own CLI startup. `pnpm --version` resolves nothing and still costs that
much. The rest is script lookup plus the extra `.CMD` shim and second `node` process
pnpm's shell line goes through. It is **not** dependency verification, so
it does not shrink with the size of the job: forcing `verify-deps-before-run` on or off
moves nothing, and on a one-file oxlint the wrapper costs more than the lint does.

So when you invoke one of the repo's binaries yourself, name it directly:

```
node node_modules/oxlint/bin/oxlint <paths>      # 1.15s repo-wide, against 1.86s for `pnpm run lint`
```

Output and exit status are byte-identical, since the `lint` script is exactly `oxlint .`.
That path is the `bin` entry oxlint's own `package.json` declares, the file the
`node_modules/.bin` shims are generated from, so it is the published contract rather than
a reach into the package. The pre-commit hook resolves its tools the same way, with the
measurements written down beside it in `lefthook.yml`.

This is a habit about *how* to call a tool, not permission to call more of them. The
hooks still own `lint` and `format:check`, per the section above. It applies to the
narrower cases that do warrant a direct call: iterating on `.oxlintrc.jsonc`, or checking
the single file you just rewrote.

It also stops at binaries. Keep `pnpm run` for `verify`, `check:static` and the other
composites. Those name **scripts**, not executables. They chain steps, and each opens
with the `ensure-dist` guard. Bypassing one means re-deriving its definition in your shell, which
is how a second, drifting copy of a gate gets created.

### Commit messages go through a file, never through a shell

Write the message to `.git/COMMIT_MSG_DRAFT` with your file-writing tool, then `git commit -F
.git/COMMIT_MSG_DRAFT`. Only a one-line `git commit -m "subject"` may be typed inline.

This is not style. On Windows, where this repo is primarily developed, an agent has both a
POSIX shell and PowerShell available, and the two disagree on here-doc syntax (`<<'EOF'` vs
`@'…'@`). Picking the wrong dialect for the tool being called does not error. The
delimiter passes through as text and lands in the subject line. Routing the message
through a file removes the shell from the path entirely, so there is no dialect left to
get wrong.

A `commit-msg` rule catches the damage when it happens. `no-shell-quoting-leak` comes
from [`shbernal/lefthook-rules`](https://github.com/shbernal/lefthook-rules) through the
`remotes:` block in `lefthook.yml`, and it rejects the shapes a leaked delimiter makes: a
line that is only `@`, `@'` or `@"`, an opener glued to the subject, a closer at the end
of a line. Treat it as a backstop, not as the place to learn the rule. By the time it
fires you have already burned the commit attempt. Five commits in this history got past
the shell before it existed, and the rule rejects all five and nothing else in 4361.

It is shared rather than local because ts-xlsx carries the same damage. It replaced a
local Node script that read the subject only, missed a closer on the last line, and would
have rejected three legitimate subjects that open with a code span.

### Targeted checks the above does not cover

- For OOXML serialization changes, add or update a fixture in `test/schema-cases.js` and
  run `pnpm run test:schema`. It needs the OOXML oracle, which `ooxml-validate` fetches
  and caches on first use, so there is nothing to install. `verify` already covers the
  schema suite. Run `test:schema` alone only to iterate on a fixture.
- For a *behavior-preserving* refactor of the `src/gen/` emitters, gate every step on the
  byte-identity harness. `pnpm run byte-identity:baseline` before the refactor, then
  `pnpm run byte-identity:check` after each step. It builds every deck registered in
  `demos/showcases/lib/showcases.mjs`, recurses into each embedded `.xlsx`, and diffs
  every part. Only three nondeterministic patterns are normalized: core.xml timestamps,
  `p14:section` ids, `c16:uniqueId`. Any other byte change is a real regression. Do not
  accept one as cleanup. The corpus is only what those decks emit, so before trusting a
  PASS, confirm the part you touched is in it (`.tmp/byte-identity/baseline/`). An emitter
  no showcase reaches is unproven, not proven unchanged.
- `pnpm run raw-xml:check` (part of `verify` and `check:static`) is a **ratchet, not a
  ban**. `scripts/raw-xml-budget.json` freezes the per-file count of hand-built XML tag
  delimiters in string literals, and the check fails if any file goes up, or if a file
  outside the budget has any at all. Most of `src/gen/` builds through `gen/oxml/el.ts`,
  but the chart emitters do not, so a flat prohibition is not yet possible. The budget is
  what stops the migration quietly un-doing itself. It also fails when a count goes
  **down**: re-freeze with `pnpm run raw-xml:freeze` in the same commit, so the budget
  never carries slack. `pnpm run raw-xml:list` prints every occurrence with a line
  number. The scan walks the TypeScript AST, so doc comments and messages handed to
  `warn` / `notes.note` / `new *Error` are not findings.
- `pnpm run path-refs:check` (part of `verify` and `check:static`) resolves every
  **backticked repo path** in `docs/`, `src/`, `test/`, `scripts/`, `tools/`, `demos/`
  and the root markdown. This repo cites files in backticks rather than as links, and
  those citations are usually the evidence for the claim beside them. `docs-check.mjs`
  only validates markdown *links*, so seven had rotted before the gate existed. A
  citation needs a `/` and a source extension. It resolves root-relative,
  file-relative, or as a path suffix (comments write `gen/oxml/el.ts` without the
  `src/` prefix), and `.js` falls back to `.ts`. Build output and `CHANGELOG.md` are
  skipped on purpose, because `RELEASING.md` lists `dist/pptxgen.*` as negative space
  and those must *never* resolve. Anything else deliberately dead goes in `ALLOWLIST` in
  `scripts/path-refs.mjs` with its reason, and an allowlist entry that stops firing
  fails the gate too. `pnpm run path-refs:list` prints every citation with its verdict.
- `pnpm run lint:chars` (part of `check:static`, and of both git hooks) keeps **em
  dashes** out of what a reader sees: the README, `www/`, and prose in `docs/`. Config
  is `charcheck.config.js`. The rules are scoped, so a dash in a code fence, a comment
  or a stylesheet is not a finding. Every surface errors, and `--max-warnings 0` sits on
  the script so a rule added at `warn` later fails rather than scrolling past.

  The gate arrived on a tree holding 684 of them in `docs/`, which warned under a frozen
  count while they were worked off. That backlog is empty, so the second severity and its
  `DOCS_CLEAN` allowlist are gone. Do not reintroduce them. A dash is a one-line rewrite
  at the moment it is written, which is the only time the author is still in the room.

  `pnpm run lint:chars:fix` rewrites repo-wide. A positional path is *intersected* with
  each rule's `include`, never substituted for it, so it cannot widen the scan.
  **Read that diff.** The fixer guesses at prose and says so on every run, and clearing
  the backlog is the evidence. Of 684 rewrites it broke a bracketed `[B — implemented]`
  label into `[B (implemented]`, nested parentheses two deep in four places, converted
  only one half of five dash *pairs*, and turned a table cell whose whole content was a
  dash into `|: |`. All caught by reading, none by a check. No hook passes `--fix`,
  deliberately: rewriting a sentence on its way into a commit is not a guess a hook
  should make unsupervised.

  If the gate ever seems *wrong* rather than inconvenient, load the `charcheck-upstream`
  skill in `.agents/skills/` and file it. Wrong looks like this: it passes a file you
  know holds a dash, it flags a region its scope says it does not read, or it tempts you
  to add an `exclude` to make a finding go away. That failure is silent by construction,
  and it has already happened here once. A trailing `\s*` in `DASH_PATTERN` matched the
  newline, ran the match past the end of a hard-wrapped line, and made charcheck drop the
  finding without a word when the next line opened with a code span
  (shbernal/charcheck#16). It hid 11 real dashes behind a clean run. Reporting it got it
  fixed in 0.2.3, and that is the point: the workaround (`[ \t]` in place of `\s`) lived
  in the config for one release and is gone.
  `node node_modules/charcheck/dist/cli.js --report-issue` is the tool for this. It
  prints each rule as it *resolved*, with the file count each one matched, and a count of
  zero is usually the whole bug. When a green run is load-bearing, prove it can still
  fail by putting one dash back and watching it report.
- **Whitespace-only byte diffs are a STOP, not a known divergence.** Inter-element
  whitespace is semantically inert, but whitespace next to character data is content, and
  the emitters keep those cases apart. Pretty-printing exists only in structural regions,
  while every text-bearing element (`<a:t>`, `<vt:lpstr>`, `<c:v>`, `<si><t>`) is
  emitted flat. Waving through "harmless" whitespace is the exact reasoning that would
  also wave through a real content change. If a part resists byte-identity, leave it
  un-migrated on template strings and list it as an exception. Do not migrate it with an
  accepted diff. A gate that admits exceptions stops being a gate and becomes a judgment
  call, precisely where fatigue is highest.
- **The one way past that STOP is a program, never a reading.**
  `node scripts/byte-identity.mjs prove-whitespace` (built on
  `scripts/xml-equivalence.mjs`) passes only when every difference is a whitespace-only
  text node in a position where whitespace provably cannot be content. It compares raw
  undecoded text, attribute order, quote characters, self-closing form and whitespace
  *inside* start tags. It freezes text in any element with no element children, any
  element carrying mixed content, and any element on its text-bearing list. This is not a
  looser `check`, and running it because `check` went red is the misuse it is shaped to
  resist. `check` remains the gate for every write-side refactor. It has been used once,
  for the `src/gen/chart/` flatten (`docs/chart-whitespace-flatten.md`), and it earned
  itself there: it caught the codemod silently taking the space out of
  `<c:layoutTarget val="inner" />`, which nobody reading a 57-part whitespace diff would
  have. A second use adds a section to that doc first, so the carve-out stays a written
  record rather than a precedent.
- For release and package boundary changes, run `pnpm run verify:full` (it bundles
  `package:lint` and `test:package`) and consult `docs/testing.md` for what each one
  covers. Nothing under `demos/` is a gate. The demos are showcases with no test role,
  so a demo that breaks fails no check and a demo that passes proves nothing about the
  published package. The byte-identity harness does *build* the showcase decks, but it
  asserts on the bytes they emit, never on the decks themselves. A showcase that throws
  simply takes that harness down with it.
- **The measurement oracles need fonts, and now say so.**
  `test/read/autofit-calibration-oracle.test.js` and
  `test/read/cjk-line-breaking-oracle.test.js` measure with the genuine faces PowerPoint
  used (Aptos, Aptos SemiBold, Arial, Calibri, Tahoma, Malgun Gothic) where the machine
  has them. Where it does not, they fall back to
  `test/read/fixtures/autofit-font-metrics.json`, a committed table of the advance widths
  of exactly the code points the cases use. Both CI lanes set `FONT_ORACLES=required`,
  which turns "no font and no sidecar entry" from a skip into a failure. Before that, both
  suites resolved nothing on CI and passed. **Add or edit an autofit or CJK case and the
  sidecar goes stale, so the suite fails naming the face and the character.** Regenerate
  with `pnpm run font-metrics:build` (needs all six faces installed, so a Windows box with
  Microsoft 365) and reformat with
  `pnpm exec oxfmt --write "test/read/fixtures/*.json"`. `pnpm run test:oracles` runs the
  probe plus the three files. `FONT_ORACLES_SIDECAR_ONLY=1` reproduces the Linux path on a
  machine that has the fonts.
- **Desktop check (Windows plus PowerPoint only, not in CI).** `pnpm run test:com` drives
  the real PowerPoint app over COM to catch what schema validation cannot: a package
  PowerPoint reports as corrupt (`0x80070570`), or an element that is schema-valid but
  semantically dead. It generates focused decks from `dist/`, opens them headless, and
  reads shape state *back out* to assert behavior, so each action-button `hlinkClick`
  has to resolve to the expected `PpActionType`. Point it at any deck with
  `--file <deck.pptx>` for just the corruption-open check. It SKIPs cleanly off Windows,
  or when PowerPoint is not COM-registered.

  Two of its checks read pixels rather than the object model, for the reason
  `docs/testing.md` gives. A 3D model that resolved is not a 3D model that drew, and an
  out-of-range `<a:gd>` adjustment guide reads back through `Shape.Adjustments` exactly
  as stored whether PowerPoint honours it or pins it. Both compare exported PNGs, and the
  preset-geometry one carries its own sensitivity pair, two *in*-range values that must
  paint differently, so a run that rendered nothing fails instead of satisfying every
  equality it makes.
- **Second render oracle (needs LibreOffice, runs in CI).** `pnpm run test:lo` renders
  decks through LibreOffice and reads the painted text back with `pdftotext`. It exists
  because PowerPoint cannot be an oracle for markup PowerPoint *recomputes*. SmartArt is
  the case that forced it. A deck stores every drawn string twice, in the
  `dgm:dataModel` PowerPoint reads and in the `dsp:drawing` cache every renderer without
  a layout engine paints, and PowerPoint regenerates the cache on open. So `test:com`
  looks identical whether the mirror ran or not, and the only evidence was the bytes of
  `ppt/diagrams/drawing1.xml`. LibreOffice has no SmartArt layout engine, so it paints
  the cache and nothing else.

  The `stale` case is the sensitivity check, and it is load-bearing: it edits the data
  model alone and asserts LibreOffice keeps painting the *old* string, so a run that
  stopped rendering or stopped extracting goes red instead of passing empty. Read text
  through PDF, never PNG. LibreOffice's PNG export writes the first slide only and
  ignores a `PageRange` filter option. It SKIPs cleanly when either tool is missing.
  `TSPPTX_SOFFICE` and `TSPPTX_PDFTOTEXT` point it at them, and a set-but-wrong path is
  an error rather than a silent fallback. LibreOffice is a no-admin install here
  (`msiexec /a` into `%LOCALAPPDATA%\Programs\LibreOffice`) and `pdftotext` ships with
  Git for Windows. The `powerpoint-fixture-authoring` skill has the provisioning detail.
  Unlike `test:com` it also runs in CI, on `ubuntu-latest`, where both tools are
  apt-installable (~41s uncached, then ~2s to run). The `render-oracle` job sets
  `TSPPTX_RENDER_ORACLE=required` so a runner that stops shipping a tool fails the leg
  instead of skipping it green.
- **The render oracle covers more than SmartArt.** `byte-identity` only covers what the
  showcase decks emit, so a construct the corpus never reaches has no evidence beyond
  "the right bytes are in the part", which says nothing about whether an independent
  implementation *acts* on them. All six known zero-baseline blind spots were probed
  against a control deck. `a:buBlip`, `a:prstTxWarp` and `numCol`/`spcCol` are
  observable through text extraction and are now cases. `rtl="1"` (identical page either
  way for Latin text) and `altLang` (no visual semantics) are not observable at all.
  `a:buClr` changes the raster but not the extraction, so it needs a rasteriser channel
  nobody has built. The `PAIRS` header in the script records this so the probe is not
  repeated. **New cases are differential**: the construct against its own control,
  asserting the two extractions differ while both still paint a canary. That is what
  keeps them portable across the xpdf and poppler builds of `pdftotext`, and what stops
  a blank render from passing. Make any new case fail on purpose before trusting it.
