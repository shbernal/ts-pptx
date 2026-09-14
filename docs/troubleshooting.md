---
doc-schema-version: 1
title: "Troubleshooting"
summary: "Observable symptoms, checks, causes, and fixes for ts-pptx."
read_when:
  - Debugging a user-visible failure
  - Adding recovery guidance
  - Explaining known failure signatures
doc_type: "troubleshooting"
---

# Troubleshooting

Start from the observable failure, then verify the relevant package boundary,
runtime, or OOXML layer before changing code.

## Import or runtime failures

Checks:

- Review [Runtime and package support](runtime-and-package-support.md).
- Run `pnpm run test:package`.
- Inspect `package.json` exports and generated declarations.

Likely causes:

- A consumer is loading a legacy upstream artifact by path (`dist/pptxgen.bundle.js`,
  `dist/pptxgen.cjs.js`) or expecting a `window.TsPptx` global, instead of importing
  the package through its exports.
- A consumer is deep-importing an internal source file.
- Generated declarations or package exports are stale.

## OOXML or PowerPoint failures

Checks:

- Review [OOXML Agent Context](https://github.com/shbernal/ts-pptx/blob/master/docs/contributing/ooxml.md).
- Add or update a focused fixture.
- Run the relevant schema or regression test.

Likely causes:

- Emitted XML is structurally invalid.
- PowerPoint accepts a structure differently from the schema.
- A fix changed package parts or relationships without matching fixtures.

## Docs or API drift

Checks:

- Run `pnpm run docs:api`.
- Run `pnpm run docs:check`.
- Compare generated API docs against `package.json` exports.

Likely causes:

- Public exports changed without regenerating docs.
- A hand-written page describes a planned API as current behavior.
- A generated reference page was edited instead of fixing the source types.
