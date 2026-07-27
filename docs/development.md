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
- Python `>=3.9`, for the `docs:*` scaffolding scripts only. `scripts/run-python.mjs`
  finds an interpreter for you (`py -3`, `python`, or `python3`, in that order on
  Windows so the non-functional Microsoft Store `python3` alias is skipped). Set
  `TSPPTX_PYTHON` to an absolute interpreter path to override the search.

Install dependencies:

```bash
pnpm install
```

## Repository Layout

- `src/`: TypeScript source.
- `test/`: regression tests, schema fixtures, and validator helpers.
- `docs/`: maintained project documentation.
- `demos/node`: Node.js ESM demo.
- `demos/vite-demo`: React, TypeScript, and Vite demo.
- `scripts/`: build, package, demo, and smoke-test automation.
- `tools/ooxml-validator`: OOXML validator installer and wrapper.
- `dist/`: generated package runtime and declaration artifacts.

Do not hand-edit generated `dist/` outputs unless the task explicitly asks to
refresh release artifacts.

This repository is developed on case-insensitive filesystems (Windows, default
macOS). Never introduce a file whose name differs from an existing one only by
case — for example a generated `INDEX.md` collides with the VitePress home page
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
characters. (A few enums, such as `TableStyle` in `core-enums.ts`, group their members
with `// ── Name ──` sub-headers; that is an intra-construct grouping, not a file region.)

### Trailing `_` marks an escape hatch

A public member whose name ends in `_` — today that is `element_` across the read
model — is a **deliberate escape hatch onto the internal representation**, not a
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
pnpm run verify       # ~45s — typechecks, backlog validation, and the whole test suite
pnpm run verify:full  # ~65s — the above plus the package boundary suites
```

`verify` is the per-change loop; `verify:full` is what to run before pushing or
when touching the release/package boundary. Both deliberately omit `lint` and
`format:check`, which the git hooks already own (see [Static Checks](#static-checks)).

Two more aggregates exist for CI, and are occasionally useful locally:

```bash
pnpm run check:static   # lint, format:check, all three typechecks, backlog:validate
pnpm run check:package  # package:lint, test:package
```

Pass flags to a script as `pnpm run lint --fix`, never `pnpm run lint -- --fix`.
pnpm forwards the `--` **literally** to the underlying binary, where it turns the
following flag into a positional argument (`No files matching the pattern
"--fix" were found`). This bites eslint and prettier identically.

`ci.yml` has **never actually run.** The repo has no remote yet, so the workflow
is verified only by parsing plus every command it invokes being green locally.
On the first push, watch the three things with no local equivalent: the
`workflow_call` from `publish.yml`, the job split (that `check:static` and
`check:package` are each self-sufficient on a fresh runner), and above all the
`windows-latest` package leg. That leg exists to cover the Windows-only branches
of `run()` in `scripts/script-utils.mjs`, which is exactly where this repo's one
live cross-platform bug lived (`run('node', …)` resolving to `node.cmd`) — expect
it to find more. If it proves flaky, mark it `continue-on-error: true` rather than
dropping it; a noisy signal beats none.

Relatedly, `.tmp/*.tsbuildinfo` is deliberately **not** cached in CI, which leaves
`incremental: true` inert there. Whether an `actions/cache` step earns its keep
depends on a *cold* CI typecheck nobody has measured yet; warm local runs
(2.1 / 1.7 / 2.3s) put it near the not-worth-it line, so the default is to leave
it alone until that first push produces a number.

No script needs to be prefixed with a build. Every gate begins with
`scripts/ensure-dist.mjs`, which compares source and config mtimes against `dist/`
and rebuilds only when it is actually stale — a ~0.1s no-op otherwise. Run
`pnpm run build` directly only when you want the bundle for its own sake.

A green `build` is **not** evidence of type-correctness. tsdown's `.d.ts` pass does
not typecheck, so a real type error (`const x: number = 'a-string'`) still builds
successfully and is caught only by `typecheck`. Never substitute one for the other.

The individual gates — `build`, `typecheck`, `typecheck:scripts`, `typecheck:test`,
`test`, `test:unit`, `test:read`, `test:schema`, `test:coverage`, `package:lint`,
`test:package`, `backlog:validate` — all still exist and are worth
running alone when iterating on one specific thing. `pnpm run` lists them.

## Static Checks

Three gates keep the source statically sound. All are green and expected to stay
that way:

```bash
pnpm run typecheck     # tsc -p tsconfig.json --noEmit
pnpm run lint          # eslint . --no-warn-ignored
pnpm run format:check  # prettier --check (includes src/**/*.ts)
```

### Who runs which gate

`lint` and `format:check` do not normally need to be run by hand. Pre-commit runs
eslint `--fix` and prettier `--write` over staged files and re-stages the result
(`stage_fixed: true`), and pre-push re-verifies the whole repo — so running
`format:check` yourself can only cost you a check→fix→re-check cycle on files that
were going to be fixed on commit anyway. What no hook covers is **tests** (none run
any), **`typecheck:test`** and **`backlog:validate`** (pre-push runs `lint`,
`format:check`, `typecheck` and `typecheck:scripts` only); those are `verify`'s job.

Note that `format`/`format:check` carry an explicit file list while pre-commit's
prettier job uses an extension glob. Every extension in the former is covered by
the latter today, but the two are maintained separately — if they drift, so does
the advice above.

The three `tsc` projects are `incremental`, with their build state under the
gitignored `.tmp/` (one `tsBuildInfoFile` each — a shared one would thrash). A warm
`typecheck` runs in roughly a third of the cold time; a cold incremental run is not
slower than a non-incremental one, so CI loses nothing.

### Line endings (LF)

All text files are checked in and checked out as **LF**, enforced by
`.gitattributes` (`* text=auto eol=lf`, with binary asset types marked `binary`).
Prettier's default `endOfLine: "lf"` relies on this. Do not depend on your local
`core.autocrlf` setting — the repo config is self-contained.

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

### TypeScript strictness

Strictness is configured once in `tsconfig.base.json` and applies to all of
`src/`. Beyond `strict: true`, the codebase enables `strictNullChecks`,
`noUncheckedIndexedAccess`, `noPropertyAccessFromIndexSignature`,
`verbatimModuleSyntax`, and the zero-cost path/usage knobs
(`noImplicitReturns`, `noFallthroughCasesInSwitch`, `noImplicitOverride`,
`noUnusedLocals`, `noUnusedParameters`). Fix new errors with real narrowing or
guards — not `!` assertions or `as` casts (both are lint errors; see below).

`exactOptionalPropertyTypes` is deliberately **left off**. The interfaces it
flags (`IChartOptsLib`, `ObjectOptions`, `BorderProps`, …) are internal
*normalized* option state, and normalization is built around "`undefined` means
use-the-default / omit" — the exact present-but-`undefined` pattern the flag
forbids. Enabling it either fights that design or risks output changes (e.g.
rewriting the latent `x || !x ? x : false` no-ops to `x ?? false` flips
`undefined → false`). Revisit only if the chart/shape option code is ever split
into distinct "raw input" and "resolved options" types, at which point the flag
becomes cheap on the input type.

### Lint policy

`src/**/*.ts` runs the type-aware set (`recommendedTypeChecked`), wired to type
info via `parserOptions.projectService`. `test/` and `scripts/` run the plain
recommended set. Two guardrail rules are pinned as **errors** to close the
compile-time escape hatches from the null-safety work:

- `@typescript-eslint/no-non-null-assertion` — bans a bare `!`.
- `@typescript-eslint/no-unnecessary-type-assertion` — bans a provably-redundant
  `as` (an intentional branding/`unknown as T` cast is not redundant and stays).

A handful of type-aware rules are intentionally relaxed to `off`
(`require-await`, `no-base-to-string`, `no-redundant-type-constituents`), each
with an inline rationale in `eslint.config.mjs`. Prettier is the sole formatter
of record; `eslint-config-prettier` disables any formatting rules that would
conflict.

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
./tools/ooxml-validator/install.sh
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
verification aggregate runs them — the published-package contract is covered by
`check:package` alone (see [Package Boundary Changes](#package-boundary-changes)).

Build the two showcase decks:

```bash
pnpm demos:build                    # both
pnpm demos:build quarterly-review   # one, by slug
```

The other two workspaces run from their own directory:

```bash
pnpm --dir demos/node run demo-stream   # streams a deck over HTTP
pnpm --dir demos/vite-demo run dev      # the same review deck, built in a browser
```

See [demos/README.md](../demos/README.md) for what each one is for.
