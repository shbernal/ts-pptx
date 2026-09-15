# AGENTS.md

These are the rules for working in this repository. Where a rule needs more than a line of
explanation, a pointer names the page that has it.

## Repository expectations

- This repository builds ts-pptx, a JavaScript and TypeScript library that emits PowerPoint
  `.pptx` packages as OOXML.
- Use `pnpm` for repository scripts. The package declares Node `>=24`.
- Keep source in `src/` and tests in `test/`. Treat `dist/` as generated unless the task asks
  for refreshed release outputs. Preserve unrelated dirty state, and never revert user
  changes. `docs/contributing/development.md` has the layout.
- `docs/` is content, `www/` is the site's application code, and `demos/` holds
  clone-and-run scripts. Do not put an application in `docs/`, and do not grow a browser app
  under `demos/` again. `docs/contributing/` is checked like the rest of `docs/` but read on
  GitHub, never built into the site. `docs/contributing/development.md` and `www/README.md`
  have the detail.

## Out of active scope

The project is Node-first. It runs and is tested without a browser or any office
application. Two domains are outside active maintenance, and the same rules apply to both:

- Do not build features there, do not hunt for fixes there, and never block other work on
  them.
- When a task lands in one, say so and treat it as out of scope, unless the user explicitly
  opts in.
- Neither is rejected on merit, and outside contributors are welcome to submit pull
  requests.

The two domains:

- **Live-DOM and browser-layout features.** Anything whose answer comes from a rendered
  page, such as `offsetWidth` after layout, the resolved cascade, or the fonts a browser
  picked. `tableToSlides()` is in scope and runs under Node with any DOM. Only real
  measurement is out of scope, and `docs/html-tables.md` describes how column widths fall
  back without it. Say "fall back", not "degrade", because the fallback basis can change a
  table's proportions. Extract the DOM-independent part of anything here into a pure helper
  and unit-test it, as `resolveHtmlColWidth`, `pickColWidthBasis` and `cssColorToHex` do.
- **Third-party office-suite interop quirks.** Breakage that appears only after another
  application round-trips a valid package, such as a copy and paste in WPS Office followed
  by opening the file in PowerPoint. The supported bar is that output opens cleanly in
  Microsoft PowerPoint. A report enters scope only with a repro that pins the defect to
  invalid OOXML the library itself emits.

`docs/contributing/scope-and-policy.md` ("Out of active scope") has the full statement and
the triage rules.

## API evolution policy

- The project is maintained for our own use and has no external backward-compatibility
  obligation. Fix root causes here rather than asking a downstream consumer to work around
  them.
- Breaking changes are acceptable, and encouraged when they make the API clearer or safer.
  Do not block an improvement on compatibility. Record each one in `CHANGELOG.md` with
  migration guidance and downstream impact. Open a GitHub issue for a breaking change you
  only propose.
- Silent coercion of invalid input is a footgun. Warn or fail on `NaN`, `undefined` or an
  out-of-range value instead of emitting a degenerate result such as a zero-size object.
  `docs/contributing/development.md` has the warn-or-throw rule.
- When you give an option value a meaning it did not have, work out what its absence now
  means. A property with three states needs three spellings. Omission is an alias for one of
  them, and which one is a decision to write down, not a fact to read off a ternary. Before
  calling such a fix complete, read the other arms of the branch you changed and ask what is
  now unreachable. Issues #9 and #10 are the worked example.
- Before adding, widening or removing an escape hatch (raw XML, a passthrough string, direct
  DOM access), read "Escape hatches" in `docs/contributing/scope-and-policy.md`. Do not
  re-derive that reasoning from anywhere else.

## OOXML and PowerPoint work

- Before changing emitted OOXML, read `docs/contributing/ooxml.md`.
- For a serialization change, add or update a fixture in `test/schema-cases.js`. `verify`
  runs the schema suite, so run `pnpm run test:schema` alone only to iterate on a fixture.
- Do not vendor standards PDFs or large extracted specification text into the repository as
  agent context. Write small, repo-specific notes with section references instead.
- Prefer executable evidence over prose. Inspect minimal PowerPoint-authored packages,
  compare package XML, and add focused regression or schema fixtures.
- Some behavior can only be judged against genuine PowerPoint output, such as a read
  accessor or a write path whose target is whatever PowerPoint authors. Where that fixture
  does not exist yet, do not implement against synthetic or round-tripped XML. Open a GitHub
  issue naming the construct the fixture has to contain, and stop until it is authored.
  "Evidence and fixtures" in `docs/contributing/ooxml.md` has the procedure.

### MCP tool selection

Look OOXML questions up in this order. "Lookup order" in `docs/contributing/ooxml.md` says
what each source answers and what it lacks.

1. **The `ooxml` MCP**, for schema structure: elements, legal children in order, attributes,
   enums, namespaces, and whether a construct survives in Strict. Hand an `ooxml-validate`
   diagnostic to `ooxml_explain`.
2. **The `microsoft_learn` MCP**, for the Microsoft Open Specifications, proprietary GUIDs and
   enumerations, PowerPoint behavior and the Open XML SDK. Always try it before web search.
3. **Web search**, only after both.

## Tracking work

- Not-yet-built work goes in a GitHub issue. That covers a bug, a proposed API change, a
  missing PPTX behavior, and a fixture that has to be authored before a feature. There is no
  local ledger. Do not add one, and do not re-add an upstream sync step.
- File through the forms in `.github/ISSUE_TEMPLATE/`. **bug** takes wrong output, repair
  prompts, regressions and fidelity limits. A fixture that has to be authored first is a bug
  of severity *fidelity limit*. **api-gap** takes a missing accessor, or a property the write
  side authors that the read side cannot read back. If neither fits, file a blank issue
  rather than bending one.
- Do not route local work through **agent-report**. That form is where `InternalError`'s
  message sends an agent working in a consumer repo. Its `agent-reported` label records where
  an issue came from, so triage still adds `bug` or `enhancement` after reading it.
- `skills/ts-pptx-upstream/` ships in the tarball as the consumer-side half of that form.
  Keep it in step with the forms and their taxonomy when either changes. Skills under
  `.agents/skills/` are for working on this repo, and each carries `metadata.internal: true`
  so a consumer's install does not offer it.
- Work you implement on the spot needs no issue. Its record is the commit history, its tests
  and `CHANGELOG.md`.
- Describe a downstream consumer's need anonymously. Issues are public and the consumer is
  not, so never name the consumer, its file paths, its decks, its clients or its content.
  "Promoting a downstream need" in `CONTRIBUTING.md` has the checklist.

## Verification

### The default loop

- Run `pnpm run verify` on every iteration instead of composing separate commands. It runs
  `ensure-dist`, `check:core` and the whole test suite.
- Run `pnpm run verify:full` before pushing, and for a release or package-boundary change. It
  adds the site build, the corpus script round trip, the package suites and both size gates.
- `check:core` is the one list of cheap checks, shared by `verify` and CI's `check:static`.
  Add a cheap check there, not to either aggregate.
- Add a check to an aggregate by naming its script in `package.json`. Never inline the
  script's command, because an inlined copy drifts from the original.
  `node scripts/run-steps.mjs --list verify` prints an expansion without running it.
- Do not prefix a command with `pnpm run build &&`. Every gate rebuilds a stale `dist/`
  itself.
- A green `pnpm run build` is not evidence of type-correctness. Only `typecheck` catches a
  type error.

The [gate matrix](docs/contributing/testing.md#gate-matrix) shows which aggregate, hook and
CI job runs each gate.

### Test suite habits

- The suite shares one module registry per worker (`isolate: false`). Module-level state in
  a test helper is shared with every other file in that worker. A cache belongs there, and
  state that carries one test's intent does not.
- Do not fix a slow run by raising `maxConcurrency`. The worker pool sizes itself from free
  memory, and `VITEST_MAX_WORKERS` pins it.
- Do not pipe a verification command through `tail` or `Select-Object -Last` on its first
  run. A setup failure explains itself near the top of the output. Filter only a re-run.
- `verify` does not run the corpus script round trip. Run `pnpm run script:roundtrip:all`
  before pushing a change to `src/script/`.
- `pnpm run script:census` is in no gate. Run it after closing a reader gap, retiring a
  fidelity note or landing a fixture, and refresh the tables in
  `docs/reference/pptx-to-script.md` in the same commit.

`docs/contributing/testing.md` explains the worker pool, the shared registry and both
converter harnesses.

### Ratchets

- `raw-xml:check`, `bundle-size:check` and `bundle-tier:check` fail on a change you did not
  intend as readily as on one you did. Each also fails when its number drops.
  `raw-xml:check` fails on any drop, and the size gates fail once a figure is far enough
  under its budget.
- Treat a ratchet failure as a prompt to find the change that moved the number. If it moved
  for a reason, re-freeze in the same commit with `raw-xml:freeze`, `bundle-size:freeze` or
  `bundle-tier:freeze`. Never re-freeze to quiet a gate without knowing which change moved
  it.
- The two size gates answer different questions, and neither replaces the other.
  `docs/contributing/testing.md` describes both.

### Coverage

- `pnpm run test:coverage` is a commit gate, not a loop check. Run it once, immediately
  before each commit.
- In the loop, run `pnpm run coverage:probe <paths>`. It runs only the test files you name,
  with the same instrumentation and the thresholds zeroed. It writes to `coverage/probe/`, so
  it never overwrites the full gate's report.
- Read a probe in one direction only. A line it reports as covered is covered. A line it
  reports as uncovered may still be reached by another suite, so never delete or rewrite a
  test on the strength of a probe miss. Confirm against the full gate first.
- The gate the repo is judged on is `pnpm run coverage:gate`, which merges the browser lane
  into the Node report. It needs `test:coverage` and `test:browser` output on disk, so leave
  it to CI's `coverage` job unless you have run both.

### Checks the hooks own

- Do not run `format`, `format:check`, `lint` or `lint:chars` yourself. The git hooks in
  `lefthook.yml` run them on every commit and push, so a manual run cannot improve the
  committed result.
- No hook runs the tests, `typecheck:test`, `docs:check` or `docs:build`. `verify` covers
  the first three. `verify:full` and CI's `docs.yml` cover `docs:build`.
- When a narrow case warrants calling a repo binary yourself, such as iterating on
  `.oxlintrc.jsonc` or checking the one file you rewrote, name its `bin` entry directly, as in
  `node node_modules/oxlint/bin/oxlint <paths>`. That is not permission to run what the hooks
  own. `lefthook.yml` records what the `pnpm` wrapper costs.
- Keep `pnpm run` for `verify`, `check:static` and the other composites. They name scripts
  that chain other scripts and open with the `ensure-dist` guard, and bypassing one creates a
  second copy of the gate.

### Byte identity

- Gate a behavior-preserving refactor of the `src/gen/` emitters on the byte-identity
  harness. Run `pnpm run byte-identity:baseline` before the refactor and
  `pnpm run byte-identity:check` after each change.
- The harness normalizes three nondeterministic values: `core.xml` timestamps, `p14:section`
  ids and `c16:uniqueId`. Any other byte change is a real regression. Do not accept one as
  cleanup.
- The corpus is only what the harness's decks emit, so before trusting a PASS, confirm the
  part you touched is in `.tmp/byte-identity/baseline/`. An emitter no deck reaches is
  unproven, not proven unchanged.
- Whitespace-only byte diffs are a STOP, not a known divergence. If a part resists byte
  identity, leave it on template strings and list it as an exception. Do not migrate it with
  an accepted diff.
- The one way past that STOP is a program, never a reading. `node scripts/byte-identity.mjs
  prove-whitespace` is for a change planned as whitespace-only before it was made. It is not
  a looser `check`, and running it because `check` went red is a misuse. "Proving a change is
  whitespace-only" in `docs/contributing/testing.md` has the procedure.
- Do not count a demo build as verification. Nothing under `demos/` is a gate.

### Other targeted checks

- `pnpm run path-refs:check` resolves every backticked repo path. Cite repo files in
  backticks. Put a path that must stay unresolved in `ALLOWLIST` in `scripts/path-refs.mjs`,
  with its reason.
- `pnpm run lint:chars` keeps em dashes out of reader-facing prose, as configured in
  `charcheck.config.js`. Write a sentence without one when you write it. Do not reintroduce a
  warning severity or an allowlist for existing dashes.
- `pnpm run lint:chars:fix` guesses at prose, so read its whole diff before committing. No
  hook passes `--fix`, on purpose.
- If charcheck seems wrong rather than inconvenient, load the `charcheck-upstream` skill in
  `.agents/skills/` and file the problem instead of adding an `exclude`.
  `node node_modules/charcheck/dist/cli.js --report-issue` prints each rule as it resolved,
  with its file count. When a green run is load-bearing, put one dash back and watch it
  report.
- Adding or editing an autofit or CJK case makes the font-metrics sidecar stale. Regenerate
  it with `pnpm run font-metrics:build` on a machine with all six faces installed, then run
  `pnpm exec oxfmt --write "test/read/fixtures/*.json"`.
- `pnpm run test:com` runs only on Windows with PowerPoint, and CI does not run it. It opens
  decks over COM to catch a package PowerPoint reports as corrupt. To learn whether
  PowerPoint paints a construct, compare exported PNGs, never properties read back over COM.
- `pnpm run test:lo` renders decks through LibreOffice and runs in CI. Read its text through
  PDF, never PNG. Write a new case as a pair that asserts the construct and its control
  extract differently while both paint a canary. Make any new case fail on purpose before
  trusting it. The `PAIRS` header in `scripts/libreoffice-render-smoke.mjs` records which
  constructs were already probed.

`docs/contributing/testing.md` covers the font oracles, the COM check, PNG evidence and the
LibreOffice oracle.

## Commit messages and release notes go through a file

- Write a commit message to `.git/COMMIT_MSG_DRAFT` with your file-writing tool, then run
  `git commit -F .git/COMMIT_MSG_DRAFT`. Only a one-line `git commit -m "subject"` may be
  typed inline.
- Pass release notes as a file too, with `gh release create --notes-file`, as the
  `release-publish` skill does.
- The POSIX shell and PowerShell disagree on here-doc syntax. A delimiter in the wrong dialect
  lands in the message without an error, and a file takes the shell out of the path.
- The `no-shell-quoting-leak` commit-msg rule, pulled from `shbernal/lefthook-rules` through
  `remotes:` in `lefthook.yml`, is a backstop. Do not rely on it to catch a leak.
