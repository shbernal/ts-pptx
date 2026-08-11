# AGENTS.md

## Repository Expectations

- This repository builds ts-pptx, a JavaScript/TypeScript library that emits PowerPoint `.pptx` packages using OOXML.
- Use `pnpm` for repository scripts. The package declares Node `>=24`.
- Keep source changes focused in `src/` and tests in `test/`. Treat `dist/` as generated package artifacts unless the task explicitly requires refreshing release outputs.
- Preserve unrelated dirty state. Do not revert user changes.
- Three trees are not the library and are easy to confuse. `docs/` is **content**: markdown
  under a frontmatter schema the docs kit validates. `www/` is the **site's application
  code** — the VitePress theme and the Vue components a page mounts, including the demos
  page that previews a deck via `pptx-html` (`www/README.md`). `demos/` is **clone-and-run
  only**: someone runs a script and gets a `.pptx`. Do not put an application in `docs/`,
  and do not re-grow a browser app under `demos/` — that one already existed once and is
  now a page of the site.

## Scope: Node-First (Two Out-Of-Active-Scope Domains)

- This project is **Node-first**: it runs and is tested without a browser or any
  office application. Two domains are out of *active* maintenance scope. Do not
  proactively build features or hunt for fixes in them, and do not block other
  work on them; when a task lands in one, say so and treat it as out of scope
  unless the user explicitly opts in. They are not rejected on merit — outside
  contributors are welcome to submit PRs — but the maintainer is not driving them.
  See `docs/project-target.md` ("Out Of Active Scope") for the full statement.
  - **Live-DOM / browser-layout features** — anything whose answer comes from a
    *rendered* page: real `offsetWidth` after layout, the resolved cascade, fonts
    as the browser chose them. `tableToSlides()` is NOT in this category any more:
    it ships as a free function on `ts-pptx/html`, runs under Node with any DOM,
    and is covered against happy-dom (`test/regression/html/html-to-slides-node.test.js`).
    Only real *measurement* is out of scope there — without a layout engine
    `offsetWidth` is `0`, so column widths fall back to computed CSS widths then to
    an equal split. Say "fall back", not "degrade": `offsetWidth` is the border box
    and computed `width` the content box, so the two bases can put a table's columns
    in different *proportions*, not merely at a coarser precision. Keep extracting
    the DOM-independent part of anything here into
    a pure helper and unit-testing it (pattern: `resolveHtmlColWidth`,
    `pickColWidthBasis`, `cssColorToHex`).
  - **Third-party office-suite interop quirks** — breakage that only appears after
    a round-trip through another app (e.g. WPS copy/paste, then PowerPoint) when
    the generated package is itself valid OOXML. The supported bar is that output
    opens cleanly in Microsoft PowerPoint. Such an item only becomes in-scope with
    a repro pinning the defect to invalid OOXML the library itself emits.

## API Evolution Policy

- This project is maintained for our own use; there is no external backward-compat
  obligation. Prefer fixing root causes here rather than asking a downstream
  consumer to work around them — a fix in this public package helps every
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
  *inherit* with no spelling at all on that path — omission already meant no-fill —
  and 3.1.0's own migration advice, "omit `fill` instead", therefore pointed at the
  wrong state; #10 is that hole. A property with three states needs three
  spellings. "The caller said nothing" is not one of them: it is an *alias* for one
  of them, and which one is a decision to make deliberately and write down, not a
  fact to read off a ternary. Before calling such a fix complete, read the other
  arms of the branch you changed and ask what is now unreachable.
- Before adding, widening, or removing an escape hatch (raw XML, a passthrough
  string, direct DOM access), read the "Escape Hatches" section of
  `docs/project-target.md`. It states the convenience-vs-guarantee rule and why
  the read path gets a deep raw hatch while the write path does not — that
  document is the whole reasoning; do not re-derive it from anywhere else.

## OOXML And PowerPoint Work

- Before changing emitted OOXML, read `docs/ooxml-agent-context.md`.
- Do not vendor full standards PDFs or large extracted specification text into this repository as agent context. Store small, repo-specific notes with section references instead.
- Prefer executable evidence over prose alone: inspect minimal PowerPoint-authored `.pptx` packages when needed, compare package XML, and add focused regression or schema fixtures.
- If a feature can only be tested against genuine PowerPoint output (a read accessor validated against real Office XML, or a write-side behaviour whose target XML is "what PowerPoint authors") and that fixture/oracle does not exist yet, do not implement against synthetic or round-tripped XML. Open a GitHub issue naming the construct the oracle must contain, and stop until the fixture is authored — see `docs/evidence-and-fixtures.md`.

### MCP Tool Selection

Two MCP servers cover complementary parts of the OOXML/PowerPoint space. Work
through them in order before falling back to web search.

**Step 1 — `ooxml` MCP** (source: ECMA-376 / ISO 29500 parsed XSDs and spec PDFs)

Use this for questions whose answer lives in the *standard itself*:
- Element and complexType definitions (`ooxml_element`, `ooxml_type`)
- Legal child elements in document order (`ooxml_children`)
- Attribute names, types, defaults, and required/optional (`ooxml_attributes`)
- Enum values for a type (`ooxml_enum`)
- Namespace URIs (`ooxml_namespace`)
- OPC package parts, content types, and relationships (`ooxml_package_part`, `ooxml_parts`)
- Free-text search across the spec PDFs (`ooxml_search`, `ooxml_section`)

The `ooxml` MCP does **not** cover Microsoft-proprietary details such as built-in
style GUIDs, behavior differences between Office versions, [MS-OE376] / [MS-PPTX]
deviations from the standard, or Open XML SDK usage. If the answer requires any of
those, move to Step 2 rather than falling back to web search.

**Step 2 — `microsoft_learn` MCP** (source: Microsoft Learn / Open Specifications)

If the `ooxml` MCP returns incomplete or no answer, always try this before web search.
Use it for:
- Microsoft Open Specifications ([MS-OE376], [MS-PPTX], [MS-OFFCRYPTO], …) —
  these document how Office *implements* or *deviates from* ECMA-376 and contain
  Microsoft-proprietary enumerations such as built-in table style GUIDs, preset
  shape adjustment ranges, and behavior flags not in the standard.
- PowerPoint-specific rendering behavior, repair heuristics, and version-gated features.
- Open XML SDK (`DocumentFormat.OpenXml`) API usage and samples.
- Azure / Microsoft 365 platform documentation.

Use `microsoft_docs_search` for a broad query first, then `microsoft_docs_fetch` on
a returned URL when you need the full page content.

**Step 3 — web search (`WebSearch` / `WebFetch`)**

Only after both MCPs have been tried and the information is still missing or
ambiguous. Useful for community discoveries (e.g. undocumented GUIDs found by
reverse-engineering), third-party library behaviour, and content that postdates the
MCPs' corpora.

## Tracking Work

- Not-yet-built work goes in a **GitHub issue**: a bug, a proposed API change, a
  missing PPTX behaviour, a fixture that has to be authored before a feature can
  be implemented. There is no local ledger — do not add one, and do not re-add an
  upstream sync step (upstream tracking is retired).
- Three forms in `.github/ISSUE_TEMPLATE/` cover the usual shapes: **bug** (wrong
  output, repair prompts, regressions, and fidelity limits — carries a severity
  and a "possible fixes" slot) and **api-gap** (a missing accessor, or a property
  the write side authors that the read side cannot read back). A fixture that has
  to be authored first is a bug of severity *fidelity limit*. If neither fits,
  file a blank issue rather than bending one of them. The third, **agent-report**,
  is not for work started here: it is where `InternalError`'s message sends an
  agent working in a *consumer* repo, and it is the only form that asks which class
  and code was thrown and where an attached file came from. Do not route local work
  through it. It labels `agent-reported`, which records *where an issue came from*
  rather than what it is — it reaches feature requests and doc mismatches as well as
  defects, so triage still adds `bug` or `enhancement` after reading.
- The `skills/ts-pptx-upstream/` skill is the consumer-side half of that form. It
  ships in the tarball (`files` includes `skills`), so it is what
  `npx skills add ./node_modules/@shbernal/ts-pptx` installs — keep it in step with
  the taxonomy and the forms when either changes. Everything under `.agents/skills/`
  is for working *on* this repo and is flagged `metadata.internal: true` so the same
  command does not offer it to a consumer.
- Work you implement on the spot needs no issue. Its record is the project's own
  commit history, its tests, and `CHANGELOG.md`.
- **Describe a downstream consumer's need ANONYMOUSLY.** Issues are public; the
  consumer is not. State the missing PPTX behaviour and how *any* consumer would
  reproduce it — never the consumer's name, its file paths, deck or client names,
  or its content. See `docs/agent-development.md` for the checklist that moves
  such a need into the project.

## Verification

### The default loop

- **`pnpm run verify`** (~44s) is the per-iteration check: a `dist/` freshness guard →
  `typecheck` → `typecheck:scripts` → `typecheck:test` → `typecheck:site` →
  `raw-xml:check` → `path-refs:check` → `docs:check` →
  the whole test suite (`vitest run`, which discovers every suite
  including schema). Run this instead of hand-composing four or five separate commands —
  hand-composed sets come out slightly different every time and end up re-running the
  same suite twice.
- **`pnpm run verify:full`** (~77s) before pushing or for a release/package-boundary
  change: everything in `verify`, plus `docs:build`, `script:roundtrip:all`,
  `package:lint`, `test:package` and `bundle-size:check`. The split is only about cost —
  those pack and install the tarball, build the production site, or run the whole read
  corpus twice; everything cheaper already lives in `verify`.
- **`docs:build` is in `verify:full`, not `verify`.** 19.7s of its 26.6s is
  `vitepress build`, a production static-site build that proves something about the
  *site*, not the library; `docs:check` keeps the docs themselves validated every
  iteration for ~3.5s. CI never depended on the loop tier for it — `docs.yml` builds the
  site on every pull request, so carrying it in `check:static` too just built it twice.
- Both aggregates are assembled by `scripts/run-steps.mjs`, which expands script names
  into their leaf commands and runs them in one process tree. `package.json` remains the
  single definition of every step; the runner only removes the ~0.7–1.3s package-manager
  relaunch between them (`verify` was paying that 13 times). It prints a per-step
  breakdown when it passes, and `node scripts/run-steps.mjs --list verify` shows the
  expansion without running anything. Add a step by editing the name list in
  `package.json` — never by inlining its command, which is how a second, drifting copy
  of a gate gets created.
- **The test suite sizes its own worker pool from free memory** (`vitest.config.ts`),
  not from the core count, because cores decide how fast it could run while memory
  decides whether the host survives it. On an idle machine this lands on the CPU bound
  and costs nothing; under memory pressure it scales down instead of driving the box
  into swap. `VITEST_MAX_WORKERS` pins it explicitly and is never clamped. Do not
  "fix" a slow run by raising `maxConcurrency` — since validator batching landed that
  knob no longer buys spawn parallelism. See docs/testing.md "Suite cost and the worker
  ceiling".
- **The suite runs `isolate: false`**, so one module registry is shared per worker rather
  than rebuilt per test file — `dist/` is >1 MB of JS and all 235 files were each
  re-evaluating it (3.5–4x on the `import` phase, 22–37% of wall clock). The consequence
  for you: module-level state in a test helper is now shared with every other file in
  that worker. A cache wants that (`test/validator.js`'s batch queue and `corpus.js`'s
  `irFor` memo both got better for it); state carrying one test's *intent* does not.
  `test/setup-globals.js` resets `setDiagnosticHandler` after every test, and
  `sequence.shuffle.files` randomizes file order so an order dependence fails rather than
  hides. See docs/testing.md "One module registry per worker".
- **`pnpm run script:roundtrip:all`** (~25s, in `verify:full` and CI) is the script
  converter's gate: for every read fixture it prints a script, runs it, and diffs the
  result against the source with the printer's own fidelity notes as the exclusion list.
  A difference no note predicted is a defect. Both tiers run — tier B against the source
  deck as template, tier A with no template at all — because they gate different claims.
  It certifies "nothing the converter can see was lost", never "nothing was lost";
  `read:census` is what measures the second thing. It is also the **only** copy of that
  sweep: the two Vitest suites each carried a byte-equivalent one, so a `verify:full` ran
  the same 44 decks × 2 tiers twice. Those suites keep what the round trip rests on and
  cannot itself establish (the diff fails when perturbed, a note excuses only its own
  field, the canonicaliser is an equivalence). So `verify` alone does **not** cover the
  corpus round trip — run `script:roundtrip:all` before pushing `src/script/` changes.
- **`pnpm run script:census`** (~2s, in no gate) counts how many fixtures raise each
  fidelity note, per tier. It gates nothing on purpose — it is the number
  `docs/reference/pptx-to-script.md` publishes, and the round trip cannot keep it true,
  because a note that excuses a difference and a note that never fires look identical to
  it. Run it after closing a reader gap, retiring a note, or landing a fixture, and
  refresh that page's tables in the same commit.
- Two of those are **ratchets**, and both fail on a change you did not intend as much as
  on one you did: `raw-xml:check` (in `verify`) and `bundle-size:check` (in
  `verify:full`). A ratchet failure is not automatically a defect — it is a prompt to
  decide whether the number moved for a reason, then re-freeze in the same commit
  (`raw-xml:freeze`, `bundle-size:freeze`) if it did. Do not re-freeze to make a gate
  quiet without knowing which change moved it.
- Builds are never something you sequence by hand. Every gate starts with
  `scripts/ensure-dist.mjs`, which rebuilds only when `src/` or a build config is newer
  than `dist/` and is otherwise a ~0.1s no-op. Do not prefix anything with
  `pnpm run build &&`.
- `pnpm run build` passing is **not** evidence of type-correctness — tsdown's `.d.ts`
  pass does not typecheck, so a genuine type error still builds clean. `typecheck` is
  the only thing that catches it; never treat a green build as a substitute.

### Coverage: probe in the loop, gate before the commit

- **`pnpm run test:coverage` (~2min, observed 60–185s) is a commit gate, not a loop
  step.** It runs every test file, including the schema suite that spawns a .NET
  OOXMLValidatorCLI process per concurrent case — none of which you need in order to
  learn whether the test you just wrote reaches the line you wrote it for. Run it
  **once, immediately before each commit**, and never before an edit. A coverage
  session that runs the full gate after every edit spends more wall-clock waiting on
  validators than on everything else combined.
- **`pnpm run coverage:probe <paths…>` is the loop step.** Same v8 instrumentation over
  the same `dist/` bundle, but scoped to the files you name and with the thresholds
  zeroed so a partial run cannot fail on a number it was never going to reach. One test
  file is ~15s; all of `test/read` is ~50s. It writes to `coverage/probe/` rather than
  `coverage/`, so a probe never overwrites the numbers the real gate produced.
- **Read the probe in one direction only.** A line the probe reports as *covered* is
  covered — that is conclusive, and it is the question the loop actually asks ("does my
  new case reach this branch?"). A line it reports as *uncovered* means only that the
  files you named do not reach it; another suite may. Never delete or rewrite a test on
  the strength of a probe miss — confirm against the full gate first.
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

### Do not run these — the git hooks already own them

- **`format`, `format:check`, `lint`.** Pre-commit runs oxlint `--fix` then oxfmt
  `--write` over staged files and re-stages what they change (`stage_fixed: true`), and
  pre-push re-verifies the whole repo (`lefthook.yml`). An agent running `format:check`
  therefore cannot improve the committed result — it can only burn a
  check→fix→re-check cycle on files that were going to be fixed anyway. This is not a
  lowered standard: the gate still runs, just not from your shell.
- What the hooks do *not* cover, and `verify` therefore does: **tests** (no hook runs
  any), **`typecheck:test`**, and **`docs:check`** (pre-push checks `lint`,
  `format:check`, `typecheck`, `typecheck:scripts` and `typecheck:site` only).
  **`docs:build`** is covered by neither — it is in `verify:full` and in CI's
  `docs.yml`, which runs on every pull request.

### Shell habit: do not pipe a verification command on its first run

Do not pipe a verification command through `| Select-Object -Last N` (or `tail`) the
first time you run it. Failures often explain themselves in a `beforeAll` or setup
message near the *top* of the output, and a tail filter scrolls exactly that off. The
concrete case: `test:schema` requires the validator installed with
`./tools/ooxml-validator/install.sh`, and its absence is reported up front — piping the
first run turned one legible missing-validator error into three blind 40s re-runs. Run
bare first; filter only on a re-run, once you know what you are looking for.

### Shell habit: reach a binary directly, keep `pnpm run` for scripts

`pnpm run X` and `pnpm exec X` add a flat ~0.7s to whatever they wrap. About 0.45s of
that is pnpm's own CLI startup — `pnpm --version`, which resolves nothing, costs that
much by itself — and the rest is script lookup plus the extra `.CMD` shim and second
`node` process pnpm's shell line goes through. It is **not** dependency verification, so
it does not shrink with the size of the job: forcing `verify-deps-before-run` on or off
moves nothing, and on a one-file oxlint the wrapper costs more than the lint does.

So when you invoke one of the repo's binaries yourself, name it directly:

```
node node_modules/oxlint/bin/oxlint <paths>      # 1.15s repo-wide, against 1.86s for `pnpm run lint`
```

Output and exit status are byte-identical — the `lint` script is exactly `oxlint .` — and
that path is the `bin` entry oxlint's own `package.json` declares, the file the
`node_modules/.bin` shims are generated from, so it is the published contract rather than
a reach into the package. The pre-commit hook resolves its tools the same way, with the
measurements written down beside it in `lefthook.yml`.

This is a habit about *how* to call a tool, not permission to call more of them — the
hooks still own `lint` and `format:check` per the section above. It applies to the
narrower cases that do warrant a direct call: iterating on `.oxlintrc.jsonc`, or checking
the single file you just rewrote.

It also stops at binaries. Keep `pnpm run` for `verify`, `check:static` and the other
composites — those name **scripts**, not executables: they chain steps and open with the
`ensure-dist` guard. Bypassing one means re-deriving its definition in your shell, which
is how a second, drifting copy of a gate gets created.

### Commit messages go through a file, never through a shell

Write the message to `.git/COMMIT_MSG_DRAFT` with your file-writing tool, then `git commit -F
.git/COMMIT_MSG_DRAFT`. Only a one-line `git commit -m "subject"` may be typed inline.

This is not style. On Windows, where this repo is primarily developed, an agent has both a
POSIX shell and PowerShell available, and the two disagree on here-doc syntax (`<<'EOF'` vs
`@'…'@`). Picking the wrong dialect for the tool being called does not error — the delimiter
is passed through as text and lands in the subject line. Routing the message through a file removes the shell from the path entirely,
so there is no dialect left to get wrong.

`scripts/check-commit-msg.mjs`, a lefthook `commit-msg` hook, rejects a subject that shows
the damage. It is a backstop, not the place to learn the rule: by the time it fires you have
already burned the commit attempt.

### Targeted checks the above does not cover

- For OOXML serialization changes, add or update a fixture in `test/schema-cases.js` and run `pnpm run test:schema` (which requires the validator installed with `./tools/ooxml-validator/install.sh`). `verify` already covers the schema suite; run `test:schema` alone only to iterate on a fixture.
- For a *behavior-preserving* refactor of the `src/gen/` emitters, gate every step on the byte-identity harness: `pnpm run byte-identity:baseline` before the refactor, then `pnpm run byte-identity:check` after each step. It builds every deck registered in `demos/showcases/lib/showcases.mjs`, recurses into each embedded `.xlsx`, and diffs every part; only three nondeterministic patterns (core.xml timestamps, `p14:section` ids, `c16:uniqueId`) are normalized. Any other byte change is a real regression — do not accept one as cleanup. The corpus is only what those decks emit, so before trusting a PASS, confirm the part you touched is in it (`.tmp/byte-identity/baseline/`) — an emitter no showcase reaches is unproven, not proven unchanged.
- `pnpm run raw-xml:check` (part of `verify` and `check:static`) is a **ratchet, not a ban**: `scripts/raw-xml-budget.json` freezes the per-file count of hand-built XML tag delimiters in string literals, and the check fails if any file goes up or a file not in the budget has any. Most of `src/gen/` builds through `gen/oxml/el.ts`, but the chart emitters do not, so a flat prohibition is not yet possible — the budget is what stops the migration silently un-doing itself. It fails when a count goes **down** as well: re-freeze with `pnpm run raw-xml:freeze` in the same commit, so the budget never carries slack. `pnpm run raw-xml:list` prints every occurrence with a line number. The scan is over the TypeScript AST, so doc comments and messages handed to `warn`/`notes.note`/`new *Error` are not findings.
- `pnpm run path-refs:check` (part of `verify` and `check:static`) resolves every **backticked repo path** in `docs/`, `src/`, `test/`, `scripts/`, `tools/`, `demos/` and the root markdown. This repo cites files in backticks rather than as links, and those citations are usually the evidence for the claim beside them — `docs-check.mjs` only validates markdown *links*, so seven had rotted before the gate existed. A citation needs a `/` and a source extension; it resolves root-relative, file-relative, or as a path suffix (comments write `gen/oxml/el.ts` without the `src/` prefix), and `.js` falls back to `.ts`. Build output and `CHANGELOG.md` are skipped on purpose — `RELEASING.md` lists `dist/pptxgen.*` as negative space and those must *never* resolve. Anything else deliberately dead goes in `ALLOWLIST` in `scripts/path-refs.mjs` with its reason, and an allowlist entry that stops firing fails the gate too. `pnpm run path-refs:list` prints every citation with its verdict.
- **Whitespace-only byte diffs are a STOP, not a known-divergence.** Inter-element whitespace is semantically inert, but whitespace adjacent to character data is content, and the emitters keep those cases separate: pretty-printing exists only in structural regions, while every text-bearing element (`<a:t>`, `<vt:lpstr>`, `<c:v>`, `<si><t>`) is emitted flat. Waving through "harmless" whitespace is therefore the exact reasoning that would also wave through a real content change. If a part resists byte-identity, leave it un-migrated on template strings and list it as an exception — do not migrate it with an accepted diff. A gate that admits exceptions stops being a gate and becomes a judgment call, precisely where fatigue is highest.
- For release/package boundary changes, run `pnpm run verify:full` (it bundles `package:lint` and `test:package`) and consult `docs/testing.md` for what each one covers. Nothing under `demos/` is a gate — the demos are showcases with no test role, so a demo that breaks fails no check and a demo that passes proves nothing about the published package. (The byte-identity harness does *build* the showcase decks, but it asserts on the bytes they emit, never on the decks themselves; a showcase that throws simply takes that harness down with it.)
- **Desktop check (Windows + PowerPoint only, not in CI):** `pnpm run test:com` drives the real PowerPoint app over COM to catch what schema validation can't — a package PowerPoint reports as corrupt (`0x80070570`), or an element that is schema-valid but semantically dead. It generates a focused deck from `dist/`, opens it headless, and reads shape state *back out* to assert behavior (e.g. each action-button `hlinkClick` resolves to the expected `PpActionType`). Point it at any deck with `--file <deck.pptx>` for just the corruption-open check. SKIPs cleanly off-Windows or when PowerPoint isn't COM-registered.
