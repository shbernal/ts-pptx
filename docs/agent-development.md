---
doc-schema-version: 1
title: "Agent Development Guide"
summary: "Expectations for agent-assisted ts-pptx changes and verification."
read_when:
  - Starting agent work in this repository
  - Updating agent workflow or evidence requirements
  - Reviewing documentation or implementation handoff rules
doc_type: "guide"
---

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
- minimal generated fixtures in `test/schema-cases.js` for regression evidence.

Avoid copying large standards text into the repository. Keep notes small and
reference section numbers or source pages when they matter.

## Package Support Guardrails

Do not reintroduce the dropped upstream targets unless the project direction is
explicitly changed:

- CommonJS support;
- IIFE/global browser bundle support;
- direct CDN script-tag support as a maintained package workflow.

The package smoke test should continue to prove that old artifacts are absent
and that `pptx-ts`, `pptx-ts/inspect`,
`pptx-ts/node`, and `pptx-ts/browser` resolve.
The `./measure`, `./read`, `./math`, and `./zip` subpaths exist in package
exports but are not yet covered by `pnpm run test:package`: see
[Runtime And Package Support](runtime-and-package-support.md).

## Promoting A Downstream Need

A downstream consumer hitting a generic PPTX gap is the main source of new work
here: an OOXML serialization fix, an API/typing gap, a repeated layout primitive,
media/SVG handling, post-processing that patches generated XML. Before moving one
into this project:

1. Prove the need with a minimal, consumer-agnostic reproduction.
2. Reduce the behavior to a minimal ts-pptx fixture.
3. Add a ts-pptx regression or schema test.
4. Pack or link the project into the downstream consumer to verify.
5. Run the consumer's build/render/lint/eval path against the linked project.
6. Keep only generic code in ts-pptx; keep project policy downstream.

Report such a gap as a GitHub issue and describe it **anonymously**: the missing
PPTX behavior and how *any* consumer would reproduce it, never the consumer's
name, file paths, deck or client names, or content.

Treat a report as evidence about PPTX generation bugs and missing features, not
as a source of package-target decisions. PowerPoint repair prompts, invalid
OOXML, content types, relationships, chart/table/media serialization, and current
TypeScript or ESM behavior are candidates. Anything that depends on CommonJS,
IIFE/global bundles, direct CDN script tags, or legacy generated artifact names is
not, unless the documented project target changes first, and see
[Project target → What Stays In The Consumer](project-target.md#what-stays-in-the-consumer)
for what this package refuses to absorb regardless of merit.

## Communication Expectations

When proposing or making a change, distinguish:

- current supported behavior;
- legacy behavior still present in demos or old docs;
- desired future behavior;
- verification that was actually run.
