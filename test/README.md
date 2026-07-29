# Test suite map

Tests are organized **by behavior**, not by source module — file names describe
the contract under test (`object-identity.test.js`, `content-type-defaults.test.js`),
not the `src/*.ts` they exercise. This index answers the two questions that
naming-by-behavior makes hard: *"does feature X already have a test?"* and
*"where do I add one?"*

For the authoritative "which lines are still uncovered" view, generate the
machine-readable coverage report and read `coverage/coverage-summary.json`
(per-file rollup) — see [Finding coverage gaps](#finding-coverage-gaps) below.

## Layout

| Path | Harness | What it covers |
|---|---|---|
| `test/regression/*.test.js` \| `.mjs` | `defineRegressionSuite()` (`helpers.js`) — see [docs/testing.md](../docs/testing.md) | write side: public API → emitted OOXML/package parts |
| `test/read/*.test.js` | Vitest `describe`/`test` | `src/read/**` lossless read + edit round-trip |
| `test/schema-cases.js` (+ `schema-validation.test.mjs`) | fixture data module | OOXML schema validation of emitted parts |
| `test/backlog-ledger.test.mjs` | Vitest | `scripts/backlog-ledger.mjs` tooling |

## `src/*.ts` → representative regression tests

Most regression tests import the whole library from `dist/node.js` and assert on
generated XML, so a single test often touches several modules. The table lists
the **primary** module each group exercises.

| Source module | Representative tests |
|---|---|
| `presentation.ts` (top-level API) | `object-identity`, `repeated-writes`, `presentation-child-order`, `presentation-layouts`, `entry-export-surface` |
| `slide.ts` | `slide-backgrounds`, `slide-hyperlinks`, `slide-title-placeholder`, `add-section-duplicate`, `object-locks` |
| `gen/define/*.ts` (add*Definition) | `addchart-signature`, `shape-presets`, `shape-text-body`, `text-formatting`, `image-shape`, `connector-shape`, `group-shapes` |
| `gen/slide/*.ts`, `gen/opc/*.ts` (spTree / part emission) | `master-*`, `notes-master-placeholders`, `notes-hyperlinks`, `slide-master-*`, `placeholder-type-attr`, `content-type-defaults`, `border-shadow-ppt-props` |
| `gen/chart/*.ts` | `chart-*`, `combo-charts`, `radar-style` |
| `gen/table/autopage.ts` (auto-paging) | `table-autopage-*`, `table-fit-columns`, `table-colwidth-distribution`, `table-header-row`, `table-merged-cell-borders`, `table-span-border-structure` |
| `gen/media.ts`, `media/*.ts` | `image-svg-source`, `image-data-dedup`, `media-load-error`, `media-loop`, `chart-embedding-parts` |
| `gen/drawingml/{color,fill,line,effect}.ts` | `hash-colors`, `alpha-colors`, `gradient-fills`, `shadow-scheme-color` |
| `measure/*.ts` (fit pass, simulator, metrics) | `text-fit*`, `measure-text-api`, `measured-fit-*`, `font-heuristic`, `font-metrics-registry`, `table-cell-fit` |
| `embedded-fonts.ts` | `embed-font`, `text-fontface-ea-cs` |
| `math.ts` | `math-latex-omml` |
| `enums.ts` / `types/index.ts` | `shape-presets`, `bullet-glyphs`, `bullet-options`, `entry-export-surface` (types) |
| `units.ts` / `units-internal.ts` | `coordinate-units`, `table-cell-margin-inches`, `table-margins` |
| `zip.ts` | `zip-compression`, `zip-output-types` |
| `index.ts` / `node.ts` / `browser.ts` (entries) | `neutral-entry`, `entry-export-surface` |
| `node.ts` / `runtime/*` | `node-runtime`, `node-runtime-fetch` |
| `read/group-transform.ts` | `group-shapes` |
| `inspect.ts` | `pptx-inspection` |
| `slide.ts` theme wiring | `theme-color-scheme`, `theme-relationships`, `theme-ea-cs-fonts` |

## `src/read/**` → read tests

| Source area | Representative tests |
|---|---|
| `read/opc/**` (package, parts, content-types, relationships) | `roundtrip`, `zip-input`, `zip-no-fs`, `model` |
| `read/oxml/**` (dom, theme, fill, color-transform) | `color-transform`, `shape-fill-edit`, `custgeom`, `template-masters` |
| `read/api/slide` + `presentation` | `edit`, `edit-existing`, `clone-slide`, `remove-slide`, `import-slide*`, `append-onto-existing` |
| `read/api/shapes` / `table` / `text` / `chart` | `shapes-edit`, `import-shape`, `table`, `chart`, `picture-edit`, `style-accessors` |
| `read/api/animation` / `transition` | `animations-transitions` |
| `read/api/theme-context` + `placeholder` inheritance | `placeholder-inherit`, `import-slide-masters`, `import-slide-restyle` |
| embedded fonts (read) | `embedded-fonts`, `append-embedded-fonts` |

Fixture provenance (real PowerPoint-authored decks) is in
[`test/read/fixtures/README.md`](read/fixtures/README.md).

## Finding coverage gaps

`vitest.config.ts` emits `json-summary` and `json` coverage reporters:

```bash
pnpm run test:coverage
```

Then read `coverage/coverage-summary.json` for a per-file
statements/branches/functions/lines rollup (no HTML scraping needed). Note the
dist-vs-src trap documented in [docs/testing.md](../docs/testing.md#coverage-gate):
a line shown red in the dist report may already be covered by a `src/`-importing
test.

## Adding a test

1. Check the table above (and `coverage-summary.json`) for an existing test of
   the behavior.
2. Name the new file after the **contract**, not a bug number
   (`slide-master-placeholders.test.js`, not `bug-123.test.js`). Record legacy
   provenance in the second arg to `defineRegressionSuite()` if relevant.
3. Prefer public-API deck generation + focused package/XML assertions. See
   [docs/testing.md § Regression Suite Layout](../docs/testing.md#regression-suite-layout).
