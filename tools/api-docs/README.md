# `tools/api-docs`

This package exists for one reason: **TypeDoc cannot run on TypeScript 7, and this repo
compiles with TypeScript 7.** It holds TypeDoc, its markdown plugin, and a pinned
TypeScript 6 for TypeDoc to use. Nothing imports it, nothing builds it, and nothing it
produces reaches the published package.

## Why the split exists

TypeScript 7 is the native Go compiler. Its npm package ships `bin/tsc` plus twenty
platform binary packages and **no JavaScript compiler API** — no `ts.SyntaxKind`, no
`ts.createProgram`. TypeDoc is built on that API, so under TypeScript 7 it fails at import
time with `TypeError: Cannot read properties of undefined (reading 'PropertyDeclaration')`.
TypeDoc 0.28's peer range stops at `6.0.x` and is honest about it; there is no released
TypeDoc that supports TypeScript 7.

So the two need different `typescript` resolutions. TypeDoc is a devDependency that reads
source and writes markdown, so leaving *documentation generation* on TypeScript 6 while the
*compiler* moves to 7 costs consumers nothing.

## Why a workspace package and not something smaller

`pnpm.overrides` was tried and **does not work**. With `"overrides": { "typedoc>typescript":
"6.0.3" }` and root TypeScript 7, pnpm still resolves TypeDoc's peer to the root 7.0.2,
still warns, and never downloads TypeScript 6 at all — an override does not bind a peer
dependency. Aliasing (`"typescript-6": "npm:typescript@^6"`) does not help either, because
TypeDoc resolves the bare specifier `typescript`.

A workspace member does work: pnpm installs `typedoc@0.28.20_typescript@6.0.3` here and
`typescript@7` at the root, with **no peer warning** — pnpm considers the graph correct,
which the override approach never achieved.

## How it is wired

`scripts/docs-api.mjs` spawns `tools/api-docs/node_modules/.bin/typedoc` with `cwd` set to
the repo root. That single path is the whole integration:

- TypeDoc resolves its plugin relative to **its own install**, not to `cwd`, so
  `typedoc-plugin-markdown` loads from here.
- Because `cwd` stays the repo root, every root-relative path in `typedoc.docs.json`
  (`entryPoints`, `out`, `tsconfig`) keeps working unchanged. That file stays at the root.

The `typescript` version here is pinned exactly (`6.0.3`, not `^6.0.3`). The point of this
package is that this copy does not move.

## The editor depends on this too

`.vscode/settings.json` points `typescript.tsdk` at
`tools/api-docs/node_modules/typescript/lib`. TypeScript 7's package ships no `tsserver.js`
and no language server, so the editor cannot use the root copy; this is the nearest pinned
one. The trade-off is written down in that file.

## When to delete it

The day TypeDoc supports TypeScript 7. Move `typedoc` and `typedoc-plugin-markdown` back to
the root devDependencies, point `typedocBin` in `scripts/docs-api.mjs` at the root
`node_modules/.bin`, drop this member from `pnpm-workspace.yaml`, resolve
`.vscode/settings.json` (see above), and delete this directory.

Note that this package is **not** the only thing pinned to the old compiler:
`scripts/raw-xml-ratchet.mjs` walks a syntax tree and imports the root's aliased
`typescript-6` devDependency for the same underlying reason. It is independent of this
package and has to be dealt with separately.
