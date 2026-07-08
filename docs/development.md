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
- `demos/vite-demo`: React, TypeScript, and Vite demo.
- `scripts/`: build, package, demo, and smoke-test automation.
- `tools/ooxml-validator`: OOXML validator installer and wrapper.
- `dist/`: generated package runtime and declaration artifacts.

Do not hand-edit generated `dist/` outputs unless the task explicitly asks to
refresh release artifacts.

## Common Commands

Build the source bundle used by tests:

```bash
pnpm run build
```

Typecheck source:

```bash
pnpm run typecheck
```

Run regression tests:

```bash
pnpm run test:unit
```

Check package contents:

```bash
pnpm run package:lint
pnpm run pack:check
pnpm run test:package
```

Smoke-test the maintained demos against the built workspace package:

```bash
pnpm run test:demos
```

## Static Checks

Three gates keep the source statically sound. All are green and expected to stay
that way:

```bash
pnpm run typecheck     # tsc -p tsconfig.json --noEmit
pnpm run lint          # eslint . --no-warn-ignored
pnpm run format:check  # prettier --check (includes src/**/*.ts)
```

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
4. Add or update a focused fixture in `test/schema.test.js`.
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
pnpm run build
pnpm run package:lint
pnpm run pack:check
pnpm run test:package
```

## Demo Changes

For Node demo changes:

```bash
pnpm run test:demo:node
```

For Vite demo changes:

```bash
pnpm run test:demo:vite
```

For both:

```bash
pnpm run test:demos
```
