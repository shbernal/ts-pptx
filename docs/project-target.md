# Project Target

PptxGenJS generates PowerPoint `.pptx` packages from TypeScript and modern
JavaScript. The project target is a maintained, ESM-first library for
applications that need to create presentations programmatically.

## Goals

- Generate `.pptx` packages without requiring PowerPoint at runtime.
- Keep the public package boundary explicit and easy to verify.
- Provide TypeScript declarations that work in modern app code.
- Support Node.js `>=24` and modern bundler-driven front-end applications.
- Preserve broad OOXML feature coverage: slides, text, tables, charts, images,
  SVGs, media, masters, and HTML table conversion.
- Make OOXML changes testable through regression tests, schema fixtures, and
  package-level smoke tests.
- Support agent-driven maintenance by documenting local evidence, validation
  commands, and OOXML research paths.

## Non-Goals

- Shipping a CommonJS build.
- Shipping a standalone IIFE/global browser build.
- Supporting direct CDN script tags as the primary browser story.
- Rebuilding the upstream release matrix around every historical artifact name.
- Treating generated `dist/` outputs as hand-edited source.

## Maintenance Posture

The repository should be understandable to a maintainer or an agent starting
from a clean checkout:

- package support is documented in `docs/runtime-and-package-support.md`;
- development commands are documented in `docs/development.md`;
- verification commands are documented in `docs/testing.md`;
- OOXML source-of-truth lookup is documented in `docs/ooxml-agent-context.md`;
- legacy autoloop files are retained but not promoted as the default workflow.
