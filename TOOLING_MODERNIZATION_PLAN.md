# PptxGenJS Tooling Modernization Plan

This plan tracks the move toward a TypeScript-first, modern, maintainable development setup. The goal is not cosmetic churn; every step should either reduce release risk, improve source correctness, make contribution easier, or make npm consumers safer.

## Target State

- Source of truth is TypeScript source, not generated artifacts or hand-maintained type mirrors.
- Runtime target is active Node.js LTS and modern browsers, not ES2016-era compatibility by default.
- Build output is generated on demand and verified before packaging.
- npm package contents are tested as the product boundary.
- Tests, linting, type checks, builds, and packaging checks are scriptable and CI-ready.
- Tooling is boring, current, documented, and easy for maintainers to update.

## Runtime Baseline

Use the active Node.js LTS line as the minimum runtime target.

As of 2026-06-06, the official Node.js release schedule lists Node 24.x as Active LTS and Node 26.x as Current. Start by targeting Node 24 LTS. Revisit the minimum when Node 26 enters LTS.

Reference: https://github.com/nodejs/Release

## Non-Negotiables

- Keep commits small and focused.
- Validate after each step.
- Do not reintroduce Gulp, Bower, checked-in `dist/`, or custom release artifact mutation.
- Do not manually edit generated package artifacts as source.
- Do not hide package correctness behind demo-specific filesystem patches.
- Preserve dual package output unless a deliberate compatibility decision says otherwise:
  - ESM import entry
  - CJS require entry
  - browser IIFE bundle
  - TypeScript declarations

## Current Baseline

Already completed:

- Runtime dependency cleanup:
  - removed npm `https`
  - moved `@types/node` out of runtime dependencies
  - removed unused `image-size`
  - dropped `bower.json`
- Build cleanup:
  - replaced `rollup-plugin-typescript2` with `@rollup/plugin-typescript`
  - removed Gulp and Gulp plugins
  - replaced Gulp release tasks with `scripts/build.mjs`
  - ignored and untracked `dist/`
  - removed the brittle `release-test` harness
  - added `pack:check` using `npm pack --dry-run --json`

Current minimum validation gate:

```bash
npm test
npm run build:dist
npm run pack:check
```

## Phase 1: Lock Publish And Pack Safety

Objective: make npm publishing safe even though `dist/` is ignored by git.

Changes:

- Add `engines.node` for the active LTS floor.
- Add `packageManager` once the package manager choice is explicit.
- Add `prepack` or `prepublishOnly` so npm packaging always runs `npm run build:dist`.
- Keep `pack:check` as the package manifest audit.
- Add a package smoke script that builds, packs, installs into a temporary fixture, and imports both ESM and CJS entries.

Suggested commits:

1. `chore: declare Node LTS engine`
2. `build: build package artifacts before packing`
3. `test: add packed package smoke test`

Validation:

```bash
npm test
npm run build:dist
npm run pack:check
npm pack --pack-destination test-tmp
```

## Phase 2: Make Type Declarations Generated

Objective: stop maintaining `types/index.d.ts` as a manual parallel API surface.

Changes:

- Split TypeScript configs:
  - `tsconfig.base.json`
  - `tsconfig.build.json`
  - `tsconfig.types.json`
  - `tsconfig.test.json` if tests move to TypeScript
- Generate declarations from `src/**/*.ts`.
- Emit declarations into the package-owned type path.
- Add a `types:build` script.
- Add a `types:check` script if generated declarations are committed, or make `prepack` generate them if they stay ignored.
- Update `exports.types` and `types` to match the generated output.

Preferred direction:

- Long term, generate declarations into `dist/` or a generated `types/` directory.
- Avoid hand edits to public declarations.
- If bundled declarations become necessary, evaluate API Extractor or Rollup declaration tooling deliberately.

Suggested commits:

1. `build: split TypeScript configs`
2. `build: generate public declarations`
3. `chore: remove manual type drift from release docs`

Validation:

```bash
npm run typecheck
npm run types:build
npm test
npm run pack:check
```

## Phase 3: Raise TypeScript And JavaScript Targets

Objective: stop compiling as if the project targets ES2016-era runtimes.

Changes:

- Set `target` and `lib` based on active Node LTS and modern browsers.
- Use `moduleResolution: "bundler"` or a modern Node-compatible setting after checking Rollup behavior.
- Add `typecheck` with `tsc --noEmit`.
- Begin reducing TypeScript looseness:
  - keep `strict: true`
  - audit `strictNullChecks: false`
  - audit `noImplicitAny: false`
- Do not flip strictness flags globally until failures are understood and staged.

Suggested commits:

1. `build: target active Node LTS`
2. `build: add TypeScript typecheck gate`
3. `refactor: prepare strict null checks incrementally`

Validation:

```bash
npm run typecheck
npm test
npm run build:dist
npm run pack:check
```

## Phase 4: Lint And Formatting Discipline

Objective: make style and correctness checks explicit and CI-ready.

Changes:

- Add scripts:
  - `lint`
  - `lint:fix`
  - `format`
  - `format:check`
- Clean `eslint.config.mjs`:
  - remove old commented legacy configs
  - decide whether style belongs in ESLint, Prettier, or both
  - add ignores for generated outputs and demos where appropriate
- Consider Prettier for formatting and ESLint for correctness.
- Add typed ESLint rules only after `typecheck` is stable.

Suggested commits:

1. `lint: add explicit lint scripts`
2. `lint: remove legacy ESLint config comments`
3. `format: add Prettier checks`

Validation:

```bash
npm run lint
npm run format:check
npm test
```

## Phase 5: Replace The Custom Test Runner

Objective: use a standard test runner with better watch mode, reporting, filtering, and TS support.

Preferred tool:

- Vitest

Migration path:

- Port existing `test/bug-*.test.js` cases without changing assertions.
- Keep OOXML schema validation as a separate suite at first.
- Add scripts:
  - `test:unit`
  - `test:schema`
  - `test:watch`
  - `test:coverage`
- Only after parity, consider moving tests to TypeScript.

Suggested commits:

1. `test: introduce Vitest runner`
2. `test: port regression fixtures to Vitest`
3. `test: keep schema validation as a dedicated suite`

Validation:

```bash
npm test
npm run test:unit
npm run test:schema
```

## Phase 6: Test The Real NPM Package

Objective: replace demo patching with packed-package fixtures.

Changes:

- Create temporary test fixtures from `npm pack` output.
- Install the produced `.tgz` into clean Node and Vite fixture directories.
- Smoke-test:
  - ESM import
  - CJS require
  - browser bundle load
  - type declaration resolution
- Keep demos as examples, not as the package correctness harness.

Suggested commits:

1. `test: add packed package import smoke tests`
2. `test: add packed package Vite fixture`
3. `test: add browser bundle smoke fixture`

Validation:

```bash
npm run build:dist
npm pack --pack-destination .tmp-pack
npm run test:package
```

## Phase 7: Build System Final Shape

Objective: keep build tooling modern while preserving the artifact contract.

Current state:

- Rollup remains a good fit because the project needs controlled ESM, CJS, and IIFE outputs.
- `scripts/build.mjs` is acceptable short term, but it should not grow into an untyped build framework.

Evaluation criteria for future bundler changes:

- Can produce ESM, CJS, minified IIFE, and browser bundle outputs.
- Can externalize `jszip` for library builds and globalize it as `JSZip` for IIFE builds.
- Can generate or integrate declaration output cleanly.
- Keeps source maps correct.
- Does not hide package exports behind framework conventions.

Candidates to evaluate later:

- Keep Rollup with cleaner config sharing.
- Move build orchestration to TypeScript using `tsx`.
- Evaluate `tsup`, `tsdown`, or `unbuild` only if they match the artifact contract without hacks.

Suggested commits:

1. `build: share Rollup config between CLI and script`
2. `build: type build orchestration`
3. `build: evaluate bundled declaration tooling`

Validation:

```bash
npm run build
npm run build:dist
npm run pack:check
```

## Phase 8: CI And Maintenance Automation

Objective: make every PR run the same reliable gates.

Changes:

- Add GitHub Actions workflow for:
  - install
  - lint
  - typecheck
  - test
  - build:dist
  - pack:check
- Cache npm dependencies.
- Run against the active LTS baseline.
- Optionally run an additional latest-current Node job as allowed-failure or advisory.
- Add dependency update strategy after the baseline is stable.

Suggested commits:

1. `ci: add Node LTS validation workflow`
2. `ci: add package smoke gate`
3. `chore: document dependency update cadence`

Validation:

```bash
npm ci
npm run lint
npm run typecheck
npm test
npm run build:dist
npm run pack:check
```

## Phase 9: Source Modernization After Tooling

Objective: use the modern tooling to make safer library changes.

Work areas:

- Remove broad `any` usage where it hides real contracts.
- Move public API types closer to implementation.
- Add focused tests around OOXML regressions before changing emitter internals.
- Isolate environment-specific code:
  - browser media loading
  - Node filesystem/HTTPS loading
  - SVG conversion behavior
- Convert fragile string/XML assembly areas only with tests in place.

This phase can change runtime behavior. Treat it as feature/refactor work, not build tooling.

## Definition Of Done

The modernization is complete when:

- Fresh clone plus one documented install can run all checks.
- `dist/` is generated, not tracked.
- `npm pack` contains exactly the intended runtime files.
- Public declarations are generated from source.
- CI enforces lint, typecheck, tests, build, and package checks.
- Tests use standard tooling.
- Node target is active LTS, documented, and enforced.
- There are no legacy tooling artifacts left for normal development.
