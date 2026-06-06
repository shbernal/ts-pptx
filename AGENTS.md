# AGENTS.md

## Repository Expectations

- This repository builds PptxGenJS, a JavaScript/TypeScript library that emits PowerPoint `.pptx` packages using OOXML.
- Use `pnpm` for repository scripts. The package declares Node `>=24`.
- Keep source changes focused in `src/` and tests in `test/`. Treat `dist/` and `types/` as generated package artifacts unless the task explicitly requires refreshing release outputs.
- Preserve unrelated dirty state. Do not revert user changes.

## OOXML And PowerPoint Work

- Before changing emitted OOXML, read `docs/ooxml-agent-context.md`.
- Use the configured `ooxml` MCP server for ECMA-376 / ISO 29500 structure, schema, legal children, attributes, enum values, namespaces, and OPC package metadata.
- Use the configured `microsoft_learn` MCP server for Microsoft Open Specifications, PowerPoint implementation behavior, Open XML SDK behavior, and compatibility notes.
- Do not vendor full standards PDFs or large extracted specification text into this repository as agent context. Store small, repo-specific notes with section references instead.
- Prefer executable evidence over prose alone: inspect minimal PowerPoint-authored `.pptx` packages when needed, compare package XML, and add focused regression or schema fixtures.

## Verification

- For source changes, run `pnpm run build` and `pnpm run typecheck` when practical.
- For behavior changes, run `pnpm run test:unit`.
- For OOXML serialization changes, add or update a fixture in `test/schema.test.js` and run `pnpm run test:schema`.
- `pnpm run test:schema` requires the validator installed with `./tools/ooxml-validator/install.sh`.
- For release/package boundary changes, consult `TESTING.md` and run the relevant package or demo smoke commands.
