---
doc-schema-version: 1
title: "Agent development guide"
summary: "Expectations for agent-assisted ts-pptx changes and verification."
read_when:
  - Starting agent work in this repository
  - Updating agent workflow or evidence requirements
  - Reviewing documentation or implementation handoff rules
doc_type: "guide"
---

# Agent development guide

Humans and coding agents both maintain this repository. The bar is the same for
either: ground the change in evidence from this checkout, and prove it with a
check someone else can run.

## Default workflow

1. Inspect the current checkout before answering or editing.
2. Preserve unrelated dirty state.
3. Keep source changes in `src/` and tests in `test/`.
4. Treat `dist/` as generated artifacts unless the task explicitly asks to
   refresh package outputs.
5. Prefer existing repository scripts and local patterns over new tooling.
6. Verify the narrowest relevant behavior before finishing.

## OOXML workflow

Before changing emitted OOXML, read
[OOXML agent context](ooxml.md).

Use:

- the OOXML MCP server for ECMA-376 structure, schema children, attributes,
  enums, namespaces, and OPC metadata;
- the Microsoft Learn MCP server for Microsoft Open Specifications,
  PowerPoint-specific behavior, and Open XML SDK behavior;
- minimal generated fixtures in `test/schema-cases.js` for regression evidence.

Avoid copying large standards text into the repository. Keep notes small and
reference section numbers or source pages when they matter.

## Package support guardrails

The package ships one ESM build, and every consumer reaches it through that:
`require()` on Node 24+ goes through Node's own ESM interop, and a browser loads it
from a bundler or an ESM CDN. Keep it that way unless the project direction is
explicitly changed. A second CommonJS artifact, an IIFE bundle assigning a `window`
global, and a classic CDN script tag are all upstream shapes this package replaced
rather than kept.

The package smoke test should continue to prove that those old artifacts
are absent and that `pptx-ts`, `pptx-ts/inspect`, `pptx-ts/node`, and
`pptx-ts/browser` resolve. The `./measure`, `./read`, `./math`, and `./zip`
subpaths are in package exports but `pnpm run test:package` does not cover them
yet. See [Where it runs](../getting-started/runtime.md).

## Promoting a downstream need

The checklist and the anonymity rule are in [CONTRIBUTING.md](https://github.com/shbernal/ts-pptx/blob/master/CONTRIBUTING.md#promoting-a-downstream-need).

## Communication expectations

When proposing or making a change, distinguish:

- current supported behavior;
- legacy behavior still present in demos or old docs;
- desired future behavior;
- verification that was actually run.
