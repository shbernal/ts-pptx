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
- `dist/`: generated package runtime artifacts.
- `types/`: generated package declaration artifacts.

Do not hand-edit generated `dist/` or `types/` outputs unless the task
explicitly asks to refresh release artifacts.

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

Build package artifacts:

```bash
pnpm run build:dist
pnpm run types:build
```

Check package contents:

```bash
pnpm run pack:check
pnpm run test:package
```

Smoke-test the maintained demos against the built workspace package:

```bash
pnpm run test:demos
```

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
pnpm run build:dist
pnpm run types:build
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
