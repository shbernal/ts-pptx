# Static-Check Hardening Plan

Plan for closing the remaining gaps in TypeScript strictness and ESLint coverage.
Both current gates pass clean; this document tracks *optional* hardening that
sharpens the tooling, not fixes for active breakage.

## Current State (baseline)

All static-check gates are green as of this writing:

| Gate                  | Command                                | Result       |
| --------------------- | -------------------------------------- | ------------ |
| Typecheck             | `tsc -p tsconfig.json --noEmit`        | clean        |
| Lint                  | `eslint . --no-warn-ignored`           | clean (0 warnings) |
| Format                | `prettier --check`                     | clean        |

> **Scope note:** `format:check` runs Prettier over configs, `scripts/`, and
> `test/` — **not** `src/**/*.ts`. Source formatting is enforced by the
> `@stylistic` ESLint rules in the lint gate, so the two gates partition the
> tree rather than overlap.

> **Update:** Gap 1 is fully closed — `strictNullChecks` is now enabled globally
> in `tsconfig.base.json`, so the standard `typecheck` gate above enforces
> null-safety across all of `src/`. The separate `strict-null:check` gate and
> `tsconfig.strict-null.json` ratchet have been retired.
>
> **Update:** `noUncheckedIndexedAccess` is now also enabled globally in
> `tsconfig.base.json` (see the former "out of scope" note below). All 276
> resulting errors were fixed with real narrowing/guards — no `!` assertions or
> `as` casts. Notable root-cause fixes: `OOXML_NS` was retyped so named prefixes
> (`OOXML_NS.p`) resolve to `string` instead of `string | undefined`; the
> chart-series accessors (`dataLabels`/`dataValues`/`dataSizes`, plus a new
> `firstLabelGroup`) now tolerate an absent first series; and the image-header
> byte parser routes its bounds-guarded reads through a `?? 0` reader helper.

`tsconfig.base.json` already enables `noImplicitAny`, `strictNullChecks`, and
`noUncheckedIndexedAccess` globally, so the codebase is stricter than average.
The core null-safety knobs are all on, and type-aware ESLint (Gap 2) is now
enabled. The remaining opportunity below is a set of not-yet-evaluated strictness
options (Gap 4).

---

## Gap 1 — `strictNullChecks` is globally disabled

### The gap

`tsconfig.base.json` sets:

```jsonc
"strict": true,           // but...
"strictNullChecks": false // ...explicitly turned back off
```

`strict: true` normally *implies* `strictNullChecks`; it has been deliberately
disabled. With null-checking off, the compiler is blind to the entire class of
`undefined` / `null` bugs — exactly the "silent coercion / degenerate result"
footguns that `AGENTS.md` says to fail on.

A migration scaffold already exists: `tsconfig.strict-null.json` opts in **3 of 24**
source files (`gen-utils.ts`, `gen-media.ts`, `slide.ts`). The ratchet has stalled.

### Size of the remaining job

Flipping `strictNullChecks` on across all of `src/` produces **403 errors**,
heavily concentrated:

| File                        | Errors |
| --------------------------- | -----: |
| `gen-xml.ts`                |    197 |
| `gen-charts.ts`             |    174 |
| `pptxgen.ts`                |     22 |
| `read/api/shapes.ts`        |      4 |
| `read/api/theme-context.ts` |      3 |
| `read/api/chart.ts`         |      2 |
| `read/api/presentation.ts`  |      1 |

~92% of the pain is two files (`gen-xml`, `gen-charts`). The newer `read/*`
subsystem is nearly null-clean already.

### Approach — incremental ratchet, not big-bang

Extend the `files` list in `tsconfig.strict-null.json` one batch at a time. Each
file added becomes a permanent ratchet that CI prevents from regressing. Avoid a
single global flip: it invites `!` non-null assertions that trade a compiler check
for a runtime footgun.

Suggested order (cheapest / highest-confidence first):

- [x] **Batch A — `read/*` (10 errors).** Added the four `read/api/*` files
      (`shapes.ts`, `theme-context.ts`, `chart.ts`, `presentation.ts`) to
      `tsconfig.strict-null.json`; fixed all 10 errors with narrowing/`?? ''`
      guards (no `!` assertions). The read subsystem is the most valuable to keep
      null-safe since it parses untrusted OOXML.
- [x] **Batch B — `pptxgen.ts` (22 errors).** Public entry surface; worth the
      null-safety.
- [x] **Batch C — `gen-charts.ts` (174 errors).** Long tail; do when touching
      chart code anyway.
- [x] **Batch D — `gen-xml.ts` (197 errors).** Longest tail; last.
- [x] **Final flip.** All of `src/**/*.ts` now typechecks clean with
      `strictNullChecks` on, so the flag moved into `tsconfig.base.json`
      (`strictNullChecks: true`), `tsconfig.strict-null.json` was deleted, and the
      `strict:null:check` script was removed. Null-checking is now enforced by the
      standard `typecheck` gate for every source file, with no ratchet to maintain.

### Guardrails

- Fix errors with real narrowing / guards, **not** `!` assertions or `as`.
- Keep `strict-null:check` non-blocking in CI until the file list is complete, so
  contributors are not tempted to suppress.

---

## Gap 2 — ESLint runs only the non-type-aware TypeScript rules

### The gap (closed)

`eslint.config.mjs` extended `tseslint.configs.recommended` — the purely syntactic
rule set. rule set. The type-aware set (`recommendedTypeChecked`) was off, so rules that need
type information did not run: `no-floating-promises`, `no-unsafe-*` (untyped `any`
flowing through), `await-thenable`, etc. These are genuinely valuable for OOXML
string-assembly code and pair naturally with the `strictNullChecks` work (both
track null/undefined flow).

### Resolution — ✅ **done**

- [x] Switched the `src/**/*.ts` block to `tseslint.configs.recommendedTypeChecked`
      and wired type info via `parserOptions.projectService: true` +
      `tsconfigRootDir`. The `test/` + `scripts/` block stays on the plain
      (non-type-aware) recommended set — extending coverage there is out of scope.
- [x] Triaged the **135** initial findings. `eslint --fix` cleared 23
      `no-unnecessary-type-assertion` (redundant `as` casts — the Gap 4 goal), then
      the remaining 112 were resolved by the root fixes and rule relaxations below.

**Root fixes (real type improvements):**

- `pptxgen.ts` `defineSlideMaster` cloned props with `JSON.parse(JSON.stringify())`
  (widens to `any`, 20 unsafe errors) → `structuredClone` preserves the
  `SlideMasterProps` type. This surfaced two masked latent bugs, both fixed:
  `newLayout.background`/`bkgd` were assigned `|| null` into optional-only fields,
  and `addBackgroundDefinition`'s signature claimed a required `props` while the
  body already treated it optional (widened to `BackgroundProps | undefined`).
- Three `new Array(n).fill(x)` sites (`any[]`) → `new Array<T>(n).fill(x)`
  (`gen-utils.ts`, `gen-tables.ts`, `read/api/chart.ts`).
- `gen-xml.ts` read `cell._optImp` (typed `any`) for fill resolution. Retyped
  `_optImp: any → unknown` and added a typed `importedFillColor(unknown)` helper
  that narrows structurally (no casts), deduping the two identical blocks. **This
  exposed a real bug:** the old `x || opts.fill ? opts.fill : ''` precedence was
  provably equivalent to `opts.fill ?? ''`, i.e. the imported (`_optImp`) fill was
  silently discarded for merged/normal table cells. Corrected to
  `importedFillColor(...) || opts.fill || ''`. Full unit + schema + read suites
  pass, but this changes emitted XML for cells carrying an `_optImp` fill — worth a
  targeted regression fixture.

**Guardrail note — type-aware `--fix` is not safe to trust blindly:** it removed
two *load-bearing* assertions (a `IChartOptsLib` subtype cast in `slide.ts` that
two mutually-assignable interfaces made look redundant, and an
`Element`-narrowing cast in `read/api/animation.ts`). Both were restored as a type
annotation and an explicit null-guard `throw`, respectively.

**Rules relaxed to `off` (systemically noisy / intentional patterns), each with an
inline rationale in `eslint.config.mjs`:**

| Rule | Count | Why off |
| ---- | ----: | ------- |
| `no-unsafe-enum-comparison`    | 60 | `CHART_NAME`(union)/`CHART_TYPE`(enum) + string/`SCHEME_COLORS` comparisons are same-valued-string safe (see Gap 4) |
| `require-await`                |  6 | async methods conform to a uniform Promise-returning contract (runtime adapters, zip/opc save) with no `await` |
| `no-base-to-string`           |  6 | deliberate `String()`/`.toString()` coercion of `unknown`/union values in OOXML string-assembly paths |
| `no-redundant-type-constituents` | 2 | public color types are `literal-union \| string` on purpose (autocomplete + escape hatch) |

The high-value type-flow rules (`no-unsafe-*`, `no-floating-promises`,
`no-misused-promises`, `await-thenable`, `no-unnecessary-type-assertion`) remain
**on and enforced**.

---

## Gap 3 — non-null assertions (`!`) were not forbidden

### The gap (closed)

A bare postfix `!` is a compile-time opt-out of exactly the null-checking that
Gap 1 turned on: it strips `null | undefined` from a type with no runtime check,
so it trades a compiler guarantee for a latent NPE. With `strictNullChecks` and
`noUncheckedIndexedAccess` now global, unguarded `!` is the remaining way to
silently defeat them.

### Resolution — ✅ **done**

- The read subsystem carried **16** non-null assertions (AST-verified via the
  rule itself, not regex; the write-side `gen-*.ts` had none). All were removed
  with real checks, **no** `as` casts:
  - Added `ownerDocumentOf()` and `replaceInParent()` to `read/oxml/dom.ts` —
    these encode the DOM's parsed-node invariants (`ownerDocument` / `parentNode`
    are non-null for tree-reachable nodes) with a loud throw instead of a silent
    `!`. This covered 13 sites across `theme.ts`, `fill.ts`, `presentation.ts`.
  - The other 3 were fixed structurally: hoisting a repeated `this.#chart()` call
    to a `const` (`read/api/chart.ts`), guarding `sourceRels.get(id)` with an
    explicit error (`read/api/presentation.ts`), and narrowing `obj.options` once
    in the measure-fit loop (`measure-fit.ts`).
- `@typescript-eslint/no-non-null-assertion` is now enabled as an **error** in
  `eslint.config.mjs`, so a bare `!` fails the standard `lint` gate. Any
  genuinely unavoidable future case must use an `eslint-disable-next-line` with a
  justification comment, which surfaces it in review rather than hiding it.

---

## Gap 4 — other strictness knobs are not yet evaluated

### The gap

Gaps 1–3 closed the null-safety knobs, but `tsconfig.base.json` leaves several
other strictness options off. None have been assessed for cost/value, so their
status is simply unknown — this gap tracks *deciding* on each, not a commitment
to enable them.

| Option                              | Guards against                                   |
| ----------------------------------- | ------------------------------------------------ |
| `exactOptionalPropertyTypes`        | `{x?: T}` silently accepting an explicit `undefined` |
| `noImplicitReturns`                 | a code path that falls off the end without a value |
| `noFallthroughCasesInSwitch`        | an unintended `case` fallthrough                 |
| `noImplicitOverride`                | a method that shadows a base method by accident  |
| `noUnusedLocals` / `noUnusedParameters` | dead bindings (partly covered by ESLint today) |
| `noPropertyAccessFromIndexSignature`| natural companion to `noUncheckedIndexedAccess`  |
| `verbatimModuleSyntax`              | type-only imports leaking into the JS output     |

### Approach

- [ ] Flip each on in isolation, count the errors, and record cost per knob.
- [ ] Enable the cheap/high-value ones (likely `noImplicitReturns`,
      `noFallthroughCasesInSwitch`, `noImplicitOverride`) directly.
- [ ] Treat `exactOptionalPropertyTypes` and `verbatimModuleSyntax` as their own
      ratchets if the error count is large.

### The `as`-cast enforcement asymmetry

Gap 1/3 guardrails repeatedly say "fix with real narrowing, **not `!` or `as`**."
The `!` half is now enforced (`no-non-null-assertion`), but `as` casts have no
corresponding lint gate. Consider `@typescript-eslint/no-unnecessary-type-assertion`
(type-aware, arrives free with Gap 2's `recommendedTypeChecked`) to catch
provably-redundant casts, and evaluate `consistent-type-assertions` for the rest.
This is folded into Gap 2 triage since the rule ships with the type-checked set.

---

## Explicitly out of scope (for now)

- **`noUncheckedIndexedAccess`** — ✅ **done.** Enabled globally in
  `tsconfig.base.json` after Gap 1's final flip; see the baseline update above.
- **Blocking `strict-null:check` in CI** before the file list is complete — moot;
  the ratchet was retired when `strictNullChecks` went global.
- **Stylistic / formatting changes** — Prettier + `@stylistic` already pass clean;
  no changes warranted.

---

## Suggested sequencing

Gaps 1, 2, and 3 are done. Remaining:

1. **Gap 4 (other strictness knobs)** — survey error counts per knob, then enable
   the cheap ones and ratchet the expensive ones. The `CHART_NAME`/`CHART_TYPE`
   unification noted here is also what would let `no-unsafe-enum-comparison` be
   turned back on.
2. **Regression fixture** for the table-cell `_optImp` fill fix landed under Gap 2
   (behavior change to emitted XML; currently covered only indirectly).
