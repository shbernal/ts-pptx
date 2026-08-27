---
doc-schema-version: 1
title: "Development Guide"
summary: "Setup, source layout, generated outputs, and contribution rules."
read_when:
  - Setting up the repository
  - Changing source layout or generated output policy
  - Updating development commands
doc_type: "guide"
---

# Development Guide

## Prerequisites

- Node.js `>=24`.
- Corepack-enabled `pnpm`.
- A local checkout of this repository.

Install dependencies:

```bash
pnpm install
```

## Repository Layout

- `src/`: TypeScript source.
- `test/`: regression tests, schema fixtures, and validator helpers.
- `docs/`: maintained project documentation.
- `demos/node`: Node.js ESM demo.
- `www/`: the site's theme and its Vue components, including the demos page.
- `scripts/`: build, package, demo, and smoke-test automation.
- `dist/`: generated package runtime and declaration artifacts.

Do not hand-edit generated `dist/` outputs unless the task explicitly asks to
refresh release artifacts.

This repository is developed on case-insensitive filesystems (Windows, default
macOS). Never introduce a file whose name differs from an existing one only by
case: for example a generated `INDEX.md` collides with the VitePress home page
`index.md` and silently overwrites it. The generated `read_when` discovery index
is named `doc-index.md` for exactly this reason.

## Source Conventions

The OOXML generators live under `src/gen/` as a layered tree (mirroring `src/read/`):
`gen/define/*` normalizes user options onto the slide model, and
`gen/{drawingml,slide,pres,opc,chart,table,anim}/*` serialize it to OOXML. (The old
`src/gen-{xml,objects,charts,tables}.ts` files are now re-export barrels.) These
modules use two comment conventions so a reader can navigate without scrolling:

- **File module map.** Each module opens with a TSDoc block that states its job and,
  for larger files, lists its regions. Read it first to orient.
- **Region banners.** File-level sections are marked with a single-line banner:

  ```ts
  // ===== Region Name =====
  ```

  The region names match the entries in that file's module-map header, so you can jump
  to a region by grepping `===== <name> =====`. Use these for coarse file structure
  (roughly one per group of related functions), not for every function.
- **Intra-function steps.** Sequential steps *inside* a function stay as `// STEP N:`
  (e.g. `// STEP 1: …`). Do not use `===== … =====` banners inside a function body.

When adding a new top-level region to one of these files, add both the banner and a
matching line in the module-map header. Prefer ASCII `=====` banners over box-drawing
characters. (A few enums, such as `TableStyle` in `enums.ts`, group their members
with `// ── Name ──` sub-headers; that is an intra-construct grouping, not a file region.)

### Trailing `_` marks an escape hatch

A public member whose name ends in `_` (today that is `element_` across the read
model) is a **deliberate escape hatch onto the internal representation**, not a
naming accident. The underscore is there to be slightly ugly: it makes hatch usage
greppable and makes it stand out in review, both in this repo and in consumer code.

Do not "tidy" such a member to the bare name; doing so silently converts a flagged
hatch into ordinary-looking API. Every `element_` is paired with a public
`markDirty()` on the same object, because handing out a live DOM node without the
obligation that comes with it is how an edit vanishes on save. See the
"Escape Hatches" section of [project target](project-target.md) for when a new
hatch is acceptable at all.

## Common Commands

Two aggregate commands cover almost every iteration; reach for the individual
scripts below only when you want one specific gate:

```bash
pnpm run verify       # ~45s — typechecks, the raw-XML ratchet, and the whole test suite
pnpm run verify:full  # ~65s — the above plus the package boundary suites
```

`verify` is the per-change loop; `verify:full` is what to run before pushing or
when touching the release/package boundary. Both deliberately omit `lint` and
`format:check`, which the git hooks already own (see [Static Checks](#static-checks)).

Two more aggregates exist for CI, and are occasionally useful locally:

```bash
pnpm run check:static   # lint, lint:chars, format:check, all four typechecks, the two ratchets, and the docs build
pnpm run check:package  # package:lint, test:package, bundle-size:check
```

`bundle-size:check` freezes what each published entry point and its chunks weigh,
gzipped, against `scripts/bundle-size-budget.json`. It fails only when an entry grows past
its budget, and asks for a re-freeze only when it comes in far enough under to be
worth banking: bytes move on every commit, and a gate that failed on every
commit would get switched off. `pnpm run bundle-size:list` shows the per-chunk
breakdown; `pnpm run bundle-size:freeze` re-baselines deliberately. The number is
a growth detector, not a download size: see
[Bundle Size](runtime-and-package-support.md#bundle-size).

Pass flags to a script as `pnpm run lint --fix`, never `pnpm run lint -- --fix`.
pnpm forwards the `--` **literally** to the underlying binary: `pnpm run lint -- --fix`
runs `oxlint . "--" "--fix"`.

With oxlint and oxfmt this fails **silently**, which is worse than how it used to
fail. Both tools accept the stray `--` and drop it along with everything after it:
that command lints the whole tree, exits 0, and applies no fixes, and
`pnpm run format:check -- --write` reports the tree clean while never writing
anything. The previous toolchain at least errored out loudly
(`No files matching the pattern "--fix" were found`). So the command looks like it
worked and did something other than what you asked: check the echoed command line
pnpm prints if a flag seems to have had no effect.

`ci.yml` runs on every push to `master` and on every pull request, and
`publish.yml` pulls in that same workflow through `workflow_call` instead of
keeping its own transcript of it. A release therefore passes the identical eight
jobs, which appear in the publish run prefixed `CI gate /`: that path is
exercised rather than assumed, and it is how v3.0.0 shipped.

The `windows-latest` leg of the `package` job is the one worth understanding. It
exists to cover the Windows-only branches of `run()` in
`scripts/script-utils.mjs`, which is exactly where this repo's one live
cross-platform bug lived (`run('node', …)` resolving to `node.cmd`). It has been
green on every run so far and has shown no flakiness, but it is also the only
job that can catch that class of bug, so if it ever does turn intermittent, mark
it `continue-on-error: true` rather than dropping the leg; a noisy signal beats
none.

Relatedly, `.tmp/*.tsbuildinfo` is deliberately **not** cached in CI, which leaves
`incremental: true` inert there. Whether an `actions/cache` step would earn its
keep depends on the size of the cold/warm gap, and under TypeScript 7 that gap is
small. Measured locally: cold (buildinfo deleted first) 1.9 / 1.5 / 2.8s for
`typecheck` / `typecheck:scripts` / `typecheck:test`, against warm runs of
1.2 / 1.4 / 1.6s, about 6.2s versus 4.2s for all three. Two seconds is well below
what restoring and saving a cache costs, so the default is to leave it alone.

Note those numbers are the *native Go* compiler's. Under TypeScript 6 the same
three gates took roughly 27.7s cold and 9.1s warm, where a cache was at least
arguable; the upgrade removed most of the reason to want one.

No script needs to be prefixed with a build. Every gate begins with
`scripts/ensure-dist.mjs`, which compares source and config mtimes against `dist/`
and rebuilds only when it is actually stale: a ~0.1s no-op otherwise. Run
`pnpm run build` directly only when you want the bundle for its own sake.

A green `build` is **not** evidence of type-correctness. tsdown's `.d.ts` pass does
not typecheck, so a real type error (`const x: number = 'a-string'`) still builds
successfully and is caught only by `typecheck`. Never substitute one for the other.

The individual gates (`build`, `typecheck`, `typecheck:scripts`, `typecheck:test`,
`test`, `test:unit`, `test:read`, `test:schema`, `test:coverage`, `package:lint`,
`test:package`, `raw-xml:check`, `bundle-size:check`) all still exist and are
worth running alone when iterating on one specific thing. `pnpm run` lists them.

One gate is in neither aggregate: `pnpm run test:browser`, the Playwright lane
that runs the package in a real Chromium (CI job `browser`). It needs a ~120 MB
browser download (`pnpm exec playwright install chromium`, once) and putting
that in the per-change loop would tax every iteration for a surface that changes
rarely. Run it when you touch `src/runtime/browser.ts`, `src/browser.ts`, the zip
writer, or anything that could plausibly emit different bytes on a different
runtime. See [Browser Lane](testing.md#browser-lane).

It starts two servers of its own (a `vite preview` for the demo and
`scripts/browser-harness-server.mjs` for the adapter harness), both on fixed
ports bound to `127.0.0.1`. Playwright manages their lifetime, so there is
nothing to start by hand, but a stale process holding 4173 or 4174 will fail the
run with `--strictPort`, which is the intended behaviour rather than silently
testing the wrong thing.

## Static Checks

Three gates keep the source statically sound. All are green and expected to stay
that way:

```bash
pnpm run typecheck     # tsc -p tsconfig.json --noEmit
pnpm run lint          # oxlint .
pnpm run lint:chars    # charcheck (em dashes in the README, www/ and docs/ prose)
pnpm run format:check  # oxfmt --check (includes src/**/*.ts)
```

`lint` is **type-aware**: `.oxlintrc.jsonc` sets `options.typeAware: true`, so the
rules that need type information (`no-floating-promises`, the `no-unsafe-*` family,
`no-misused-promises`) really do run. They are executed by `oxlint-tsgolint`, a Go
typechecker oxlint shells out to: which is why that devDependency exists and why
its version tracks TypeScript 7 rather than oxlint. The option lives in the config
rather than behind a `--type-aware` flag so an editor's oxlint integration reaches
the same verdict as `pnpm run lint`.

### Who runs which gate

`lint` and `format:check` do not normally need to be run by hand. Pre-commit runs
oxlint `--fix` and oxfmt `--write` over staged files and re-stages the result
(`stage_fixed: true`), and pre-push re-verifies the whole repo, so running
`format:check` yourself can only cost you a check→fix→re-check cycle on files that
were going to be fixed on commit anyway. What no hook covers is **tests** (none run
any), **`typecheck:test`** and **`docs:build`** (pre-push runs `lint`, `lint:chars`,
`format:check`, `typecheck`, `typecheck:scripts` and `typecheck:site` only); those are
`verify`'s job.

`lint:chars` is owned by the hooks the same way, but it is the one gate that runs at two
different scopes. Pre-commit scans the *staged content* of the files in the commit, which
is what makes it accurate about what you are actually shipping. Pre-push scans the whole
repo, which catches prose that reached the branch without passing that hook: a commit made
with `--no-verify`, one merged in, or a rebase that resurrected a line. Every rule errors,
and both runs carry `--max-warnings 0`. See [Static Checks](#static-checks) above and
`charcheck.config.js`, which carries the reasoning.

When the gate itself looks wrong, the thing to reach for is
`node node_modules/charcheck/dist/cli.js --report-issue`. charcheck's characteristic
failure is silence rather than an exception: a rule whose globs reach no file reports a
clean run and exits 0, which is indistinguishable from a scan that passed. That flag
prints every rule *as it resolved*, including how many files each one matched, so a rule
matching zero is visible instead of invisible. It reads no file's content and exits 0
whatever the tree holds. The `charcheck-upstream` skill in `.agents/skills/` covers the
triage and files the report; it ships inside the package, so refresh the copy
(`npx skills update charcheck-upstream`) in the same commit that bumps the pin, and
re-apply the local `metadata.internal: true` flag afterwards.

That silence is not hypothetical here. `DASH_PATTERN` in `charcheck.config.js` ends in
`\s*`, and `\s` matches a newline: until charcheck 0.2.3 that carried the match off the end
of a hard-wrapped line, and when the next line opened with an inline code span the finding
was dropped without a word (shbernal/charcheck#16). Eleven genuine dashes sat behind a clean
run that way, and the workaround was to match `[ \t]` instead. 0.2.3 fixed it and the
pattern is back to `\s*`, which is why the pin is exact. So when a green scan is the
evidence for a claim, prove the gate can still fail: put one dash back, watch it report,
take it out again.

Note that `format`/`format:check` carry an explicit file list while pre-commit's
oxfmt job uses an extension glob. Every extension in the former is covered by
the latter today, but the two are maintained separately: if they drift, so does
the advice above.

There is a third list, and it is a safety net rather than a definition:
`.oxfmtrc.jsonc`'s `ignorePatterns`. `format:run` hands oxfmt an explicit set of
globs instead of the bare `oxfmt` that would otherwise suffice, because bare oxfmt
considers a wider set than the explicit list covers: it reaches markdown,
CSS, HTML and `.vue` single-file components, none of which this repo
formats. A silently *wider* set is how a tool-written file gets clobbered, so both
defences are kept and they overlap on purpose.

Three of the four `tsc` projects are `incremental`, with their build state under the
gitignored `.tmp/` (one `tsBuildInfoFile` each: a shared one would thrash). `typecheck:site`
is not: it is small, and it is the one project whose inputs are mostly `node_modules`. A warm
`typecheck` runs in roughly a third of the cold time; a cold incremental run is not
slower than a non-incremental one, so CI loses nothing.

### Line endings (LF)

All text files are checked in and checked out as **LF**, enforced by
`.gitattributes` (`* text=auto eol=lf`, with binary asset types marked `binary`).
oxfmt writes LF and relies on this. Do not depend on your local `core.autocrlf`
setting: the repo config is self-contained.

On Windows, a working tree that predates the `.gitattributes` (or a fresh clone
with `core.autocrlf=true` and no attributes applied) can materialize files as
CRLF, which makes `pnpm run format:check` report every text file as mis-formatted
and makes `pnpm run format --write` rewrite all of them. If that happens, do **not**
run `format --write`; instead re-normalize the working tree to LF (the blobs are
already LF, so this changes only line endings, not content):

```bash
git rm -r --cached -q .
git reset --hard        # re-checks-out every tracked file as LF per .gitattributes
```

### Two TypeScript versions are installed, on purpose

`tsc` is **TypeScript 7**, the native Go compiler. Two things still hold a pinned
copy of **TypeScript 6**, and neither is an oversight:

- `tools/api-docs`: a private workspace package holding TypeDoc, its markdown
  plugin, and `typescript@6.0.3` for TypeDoc to use.
- `typescript-6` in the root devDependencies: an alias of `typescript@6` that
  `scripts/raw-xml-ratchet.mjs` imports.

The cause is the same for both. TypeScript 7's npm package ships `bin/tsc` plus
twenty platform binaries and **no JavaScript compiler API**: no `ts.SyntaxKind`,
no `ts.createProgram`. Anything that walks a syntax tree therefore cannot run on
it: TypeDoc dies at import time, and the raw-XML ratchet walks a tree by design.
Both are dev-only, so leaving them on 6.x costs consumers nothing, and the
published `.d.ts` output is unaffected.

They are separate pins that unwind separately (`tools/api-docs` frees up when
TypeDoc supports TypeScript 7, the alias when the ratchet does), so Renovate
holds each below 7 with its own rule and its own `description`. Note that
`pnpm.overrides` cannot express this: an override does not bind a peer
dependency, which is why the TypeDoc copy needs a whole workspace package rather
than one line of config. `tools/api-docs/README.md` has the full reasoning,
including what to delete when the day comes.

This is also why `.vscode/settings.json` points `typescript.tsdk` at the
`tools/api-docs` copy: TypeScript 7 ships no `tsserver.js`, so the editor cannot
use the root one.

### TypeScript strictness

Strictness is configured once in `tsconfig.base.json` and applies to all of
`src/`. Beyond `strict: true`, the codebase enables `strictNullChecks`,
`noUncheckedIndexedAccess`, `noPropertyAccessFromIndexSignature`,
`verbatimModuleSyntax`, and the zero-cost path/usage knobs
(`noImplicitReturns`, `noFallthroughCasesInSwitch`, `noImplicitOverride`,
`noUnusedLocals`, `noUnusedParameters`). Fix new errors with real narrowing or
guards: not `!` assertions or `as` casts (both are lint errors; see below).

`exactOptionalPropertyTypes` is deliberately **left off**. The interfaces it
flags (`IChartOptsLib`, `ObjectOptions`, `BorderProps`, …) are internal
*normalized* option state, and normalization is built around "`undefined` means
use-the-default / omit": the exact present-but-`undefined` pattern the flag
forbids. Enabling it either fights that design or risks output changes (e.g.
rewriting the latent `x || !x ? x : false` no-ops to `x ?? false` flips
`undefined → false`). Revisit only if the chart/shape option code is ever split
into distinct "raw input" and "resolved options" types, at which point the flag
becomes cheap on the input type.

### Lint policy

`src/**/*.ts` runs the type-aware set; `test/` and `scripts/` run a syntax-only
set. Two guardrail rules are pinned as **errors** to close the compile-time
escape hatches from the null-safety work:

- `typescript/no-non-null-assertion`: bans a bare `!`.
- `typescript/no-unnecessary-type-assertion`: bans a provably-redundant `as` (an
  intentional branding/`unknown as T` cast is not redundant and stays).

They are pinned together, and deliberately so: the point is the `!`/`as`
symmetry, and pinning both by name means it survives any upstream change to the
preset either one lives in.

A handful of type-aware rules are intentionally relaxed to `off`
(`require-await`, `no-base-to-string`, `no-redundant-type-constituents`), each
with an inline rationale in `.oxlintrc.jsonc`. There is no formatting-rule
conflict to manage: **oxlint ships no formatting rules at all**, so nothing needs
disabling and no compatibility package belongs in this repo. oxfmt is the sole
formatter of record.

### What the baseline is, and what it is not

oxlint's categories are not a drop-in for the old preset pair, and the gap was
measured rather than guessed. `correctness` alone reports 17 findings on this
tree; adding `suspicious` reports 753; adding `pedantic`, 1509. Those extra 1492
are not latent bugs that had been going unnoticed: they are a different
linter's house style, and several contradict decisions this repo has already
recorded (`suspicious` re-enables `require-await`; `pedantic` wants `eqeqeq`
across `read/api`). Adopting them would be a style migration wearing a toolchain
swap's clothes.

So the baseline is `correctness`, and parity is reached by **naming rules
explicitly**. Of the 90 rules previously enabled on `src/**/*.ts`, 57 are already
in `correctness` and 32 more are turned on by name. Exactly one is lost
(`no-octal`, which oxlint does not implement), and it is lost harmlessly: legacy
octal literals are a syntax error in strict mode, every file here is an ES
module, and modules are always strict, so the parser already forbids what the
rule forbade.

The swap also **tightens** the gate in one place. `scripts/**` and `test/**` stay
syntax-only, but `no-floating-promises` and `no-misused-promises` now stay on
there. Those two were always the pair worth having; under the previous linter
they could not be enabled without dragging the whole `no-unsafe-*` family along,
and oxlint lets them stand alone.

One structural difference is worth knowing before editing `.oxlintrc.jsonc`. The
old flat config scoped every block by `files`, so a file matching no block was
linted with **zero** rules: which is how `tools/**` and `docs/**` came to be
unlinted without anyone deciding it. oxlint inverts that: top-level `rules` apply
to every file that is not ignored. Both trees are therefore named in
`ignorePatterns`, to hold them at the enforcement level they have always had.
Starting to lint them is a decision worth taking on its own merits, not a side
effect of changing linters.

## OOXML Changes

Before changing emitted OOXML, read
[OOXML agent context](ooxml-agent-context.md).

For serialization changes:

1. Search the local source and tests first.
2. Use the configured OOXML MCP server for schema structure, children,
   attributes, enums, namespaces, and OPC package metadata.
3. Use the configured Microsoft Learn MCP server for PowerPoint and Open XML
   SDK behavior.
4. Add or update a focused fixture in `test/schema-cases.js`.
5. Run schema validation:

```bash
pnpm run test:schema
```

## Package Boundary Changes

The package is ESM-only. Changes to package exports, generated filenames, or
package contents should preserve the support contract documented in
[runtime and package support](runtime-and-package-support.md).

Package-boundary verification:

```bash
pnpm run check:package
```

## Demo Changes

The demos are showcases, not tests. Nothing under `demos/` gates a commit, and no
verification aggregate runs them: the published-package contract is covered by
`check:package` alone (see [Package Boundary Changes](#package-boundary-changes)).

Build the two showcase decks:

```bash
pnpm demos:build                    # both
pnpm demos:build quarterly-review   # one, by slug
```

The streaming demo runs from its own directory, and the browser story is a page of the
site rather than a workspace of its own:

```bash
pnpm --dir demos/node run demo-stream   # streams a deck over HTTP
pnpm run docs:dev                       # the site, including /demos — the same deck in a browser
```

See [demos/README.md](https://github.com/shbernal/ts-pptx/blob/master/demos/README.md)
for what each one is for.

## Site Changes

The project site is one VitePress build covering everything at
`https://shbernal.github.io/ts-pptx/`: the front page, the docs, and the demos page. It
is split across two trees on purpose:

- **`docs/`** is content. Markdown under the frontmatter schema, navigated from
  `docs.json`, validated by `docs:check`.
- **`www/`** is the code that renders it: the VitePress theme, its stylesheet, and the Vue
  components a page mounts. See
  [www/README.md](https://github.com/shbernal/ts-pptx/blob/master/www/README.md).

VitePress only looks for a theme at `<root>/.vitepress/theme`, so
`docs/.vitepress/theme/index.ts` is a one-line re-export of `www/theme`. That shim is the
whole cost of keeping an application out of the docs tree.

```bash
pnpm run docs:dev       # hot-reloaded, at http://localhost:5173/ts-pptx/
pnpm run docs:build     # what CI publishes; runs docs:check on both sides of the build
pnpm run docs:preview   # serve the built output, which is what the browser lane drives
pnpm run typecheck:site # tsc over www/**/*.ts and docs/.vitepress/**
```

Two things to know before changing the demos page:

- **It is a test fixture.** The Playwright `demo` project drives `/demos`: it is the only
  place `src/runtime/browser.ts`'s `writeFile` executes, and where the browser-built deck
  is compared byte for byte against the Node-built one. The specs bind by ARIA role
  (`getByRole('group', { name: 'Download' })`, then a button matching `/^Build /`, then
  `role="status"` / `role="alert"` inside that group). Rearrange the markup freely; keep
  those roles.
- **`.vue` files are typechecked by nothing.** `tsc` does not read single-file components
  and this repo carries no `vue-tsc`. So the page's logic lives in
  `www/demos/deck-preview.ts` (which `typecheck:site` reads and
  `test/regression/www/deck-preview.test.js` covers) and the component is markup around
  it. Adding logic to the SFC quietly moves it outside both.
