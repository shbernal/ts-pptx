# Agent Development Guide

This repository is intended to be maintainable by both humans and coding
agents. Agent-driven changes should be grounded in local evidence and verified
with executable checks.

## Default Workflow

1. Inspect the current checkout before answering or editing.
2. Preserve unrelated dirty state.
3. Keep source changes in `src/` and tests in `test/`.
4. Treat `dist/` as generated artifacts unless the task explicitly asks to
   refresh package outputs.
5. Prefer existing repository scripts and local patterns over new tooling.
6. Verify the narrowest relevant behavior before finishing.

## OOXML Workflow

Before changing emitted OOXML, read
[OOXML agent context](ooxml-agent-context.md).

Use:

- the OOXML MCP server for ECMA-376 structure, schema children, attributes,
  enums, namespaces, and OPC metadata;
- the Microsoft Learn MCP server for Microsoft Open Specifications,
  PowerPoint-specific behavior, and Open XML SDK behavior;
- minimal generated fixtures in `test/schema.test.js` for regression evidence.

Avoid copying large standards text into the repository. Keep notes small and
reference section numbers or source pages when they matter.

## Package Support Guardrails

Do not reintroduce the dropped upstream targets unless the project direction is
explicitly changed:

- CommonJS support;
- IIFE/global browser bundle support;
- direct CDN script-tag support as a maintained package workflow.

The package smoke test should continue to prove that old artifacts are absent
and that `@shbernal/pptxgenjs`, `@shbernal/pptxgenjs/core`,
`@shbernal/pptxgenjs/node`, `@shbernal/pptxgenjs/browser`, and
`@shbernal/pptxgenjs/standalone` resolve.

## Upstream Signals

When reviewing upstream issues or pull requests for potential local work, use
[upstream signal workflow](upstream-signal-workflow.md) and record decisions in
`docs/upstream-signals.yml`.

Treat upstream as a source of evidence for PPTX generation bugs and missing
features, not as a source of package-target decisions. Signals about PowerPoint
repair prompts, invalid OOXML, content types, relationships, chart/table/media
serialization, and current TypeScript or ESM behavior can be target candidates.
Signals that depend on CommonJS, IIFE/global bundles, direct CDN script tags, or
legacy generated artifact names are non-target unless the documented project
target changes first.

## Communication Expectations

When proposing or making a change, distinguish:

- current supported behavior;
- legacy behavior still present in demos or old docs;
- desired future behavior;
- verification that was actually run.
