# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.0.0] - 2026-07-24

Initial public release of ts-pptx — an ESM-first, TypeScript-first library for
generating PowerPoint `.pptx` files from Node.js and modern JavaScript
toolchains.

ts-pptx descends from [gitbrent/PptxGenJS](https://github.com/gitbrent/PptxGenJS)
(MIT) and has been developed independently since; see the README for lineage and
the [project target](docs/project-target.md) for scope. It ships its own API and
makes no backwards-compatibility guarantee with the original project.

### Added

- Slide authoring: slides, layouts, masters, sections, speaker notes, and
  presentation metadata.
- Content: text, tables, shapes, connectors, groups, images, SVGs, charts,
  media, and OLE objects.
- Outputs: file, stream, buffer, Blob, base64, and browser download, depending
  on the runtime.
- A `.pptx` read model for opening, inspecting, and round-tripping existing
  decks (`@shbernal/ts-pptx/read`).
- Standalone text measurement and table-fit helpers
  (`@shbernal/ts-pptx/measure`).
- Native equation authoring from LaTeX or MathML (`@shbernal/ts-pptx/math`).
- Explicit ESM package boundary with typed subpath exports for `core`,
  `inspect`, `measure`, `read`, `math`, `node`, `browser`, and `standalone`.

[Unreleased]: https://github.com/shbernal/ts-pptx/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/shbernal/ts-pptx/releases/tag/v1.0.0
