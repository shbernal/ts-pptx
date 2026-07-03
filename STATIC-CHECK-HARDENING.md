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
rule set. The type-aware set (`recommendedTypeChecked`) was off, so rules that need
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
- `gen-xml.ts` read `cell._optImp` (typed `any`) for table-cell fill resolution.
  Retyping `any → unknown` forced the read into a narrowing helper, which in turn
  exposed that `TableCell._optImp` is **dead**: it is *read* at two fill sites but
  never *written* anywhere in the fork (`git log -S"_optImp ="` is empty), and
  `addTableDefinition` rebuilds every cell as a fresh object that drops it, so it is
  always `undefined` at emission. It is a vestige of the browser HTML-import path
  (`tableToSlides`, out of active scope), which no longer populates it. Rather than
  keep an always-`undefined` field and a helper guarding an unreachable branch, the
  field, the helper, and both reads were removed; the fill lines are now plain
  `cellOpts.fill || ''` / `originOpts.fill || ''`. Full unit suite (406) unchanged,
  confirming zero observable-output impact.

**Guardrail note — type-aware `--fix` is not safe to trust blindly:** it removed
two *load-bearing* assertions (a `IChartOptsLib` subtype cast in `slide.ts` that
two mutually-assignable interfaces made look redundant, and an
`Element`-narrowing cast in `read/api/animation.ts`). Both were restored as a type
annotation and an explicit null-guard `throw`, respectively.

**Rules relaxed to `off` (systemically noisy / intentional patterns), each with an
inline rationale in `eslint.config.mjs`:**

| Rule | Count | Why off |
| ---- | ----: | ------- |
| `require-await`                |  6 | async methods conform to a uniform Promise-returning contract (runtime adapters, zip/opc save) with no `await` |
| `no-base-to-string`           |  6 | deliberate `String()`/`.toString()` coercion of `unknown`/union values in OOXML string-assembly paths |
| `no-redundant-type-constituents` | 2 | public color types are `literal-union \| string` on purpose (autocomplete + escape hatch) |

> `no-unsafe-enum-comparison` was on this list (15 findings) but is now **on and
> enforced** — see the "`no-unsafe-enum-comparison` — ✅ resolved" section below.

The high-value type-flow rules (`no-unsafe-*`, `no-floating-promises`,
`no-misused-promises`, `await-thenable`, `no-unnecessary-type-assertion`) remain
**on and enforced**. `no-unnecessary-type-assertion` is additionally pinned
explicitly in the rules block (alongside `no-non-null-assertion`) so the `!`/`as`
symmetry is legible and survives any upstream preset change.

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

## Gap 4 — other strictness knobs (surveyed; all resolved)

### The gap

Gaps 1–3 closed the null-safety knobs, but `tsconfig.base.json` left several
other strictness options off. Each has now been surveyed in isolation
(`tsc -p tsconfig.json --<flag>`, tsc 6.0.3), the cost recorded below, and a
decision reached for every one. Seven of the eight are enabled; the eighth
(`exactOptionalPropertyTypes`) is deliberately deferred with the rationale below.

| Option                              | Errors | Guards against                                   | Status |
| ----------------------------------- | -----: | ------------------------------------------------ | ------ |
| `noImplicitReturns`                 |      0 | a code path that falls off the end without a value | ✅ enabled |
| `noFallthroughCasesInSwitch`        |      0 | an unintended `case` fallthrough                 | ✅ enabled |
| `noImplicitOverride`                |      0 | a method that shadows a base method by accident  | ✅ enabled |
| `noUnusedLocals`                    |      0 | dead local bindings                              | ✅ enabled |
| `noUnusedParameters`                |      0 | dead parameters                                  | ✅ enabled |
| `noPropertyAccessFromIndexSignature`|     25 | natural companion to `noUncheckedIndexedAccess`  | ✅ enabled |
| `verbatimModuleSyntax`              |     63 | type-only imports leaking into the JS output     | ✅ enabled |
| `exactOptionalPropertyTypes`        |    112 | `{x?: T}` silently accepting an explicit `undefined` | ⛔ deferred (deliberate) |

### Approach

- [x] Flip each on in isolation, count the errors, and record cost per knob
      (table above).
- [x] Enable the five zero-cost knobs directly in `tsconfig.base.json`
      (`noImplicitReturns`, `noFallthroughCasesInSwitch`, `noImplicitOverride`,
      `noUnusedLocals`, `noUnusedParameters`). Build, typecheck, and lint all stay
      green — no source changes were required.
- [x] `noPropertyAccessFromIndexSignature` (25) — converted the 25 dot-accesses of
      index-signature properties to bracket notation (24 dynamic `fast-xml-parser`
      `XmlNode` reads in `inspect.ts`, one DOM `DOMStringMap` `dataset` access in
      `runtime/browser.ts`). These are genuinely dynamic maps, so bracket access is
      correct, not a workaround. Enabled in `tsconfig.base.json`.
- [x] `verbatimModuleSyntax` (63) — marked the 63 type-only imports (all TS1484):
      converted the two all-type import blocks (`core-interfaces` in `pptxgen.ts`
      and `slide.ts`) to `import type`, and added inline `type` modifiers to the
      type names mixed into value-import blocks. Enabled in `tsconfig.base.json`.

### `exactOptionalPropertyTypes` — why it is deferred (deliberate, not un-surveyed)

The survey produced 112 errors, but — unlike the other seven knobs — enabling it
is **low-value and high-risk on this codebase**, so it is intentionally left off:

- **The targets are internal *normalized* types, not the public input surface.**
  The flagged interfaces (`IChartOptsLib`, `ObjectOptions`, `IChartPropsTitle`,
  `BorderProps`, `ShapeLineProps`, …) are the mutable working state built up during
  option normalization, and they legitimately hold `T | undefined` mid-flight.
- **The codebase idiom is the opposite of what the flag enforces.** Normalization
  is written around "`undefined` means use-the-default / omit" (e.g. the eight
  explicit `… : undefined` assignments and the many `x || !x ? x : default` lines
  in `gen-objects.ts`). `exactOptionalPropertyTypes` exists to forbid a
  present-but-`undefined` property, which is exactly this pattern.
- **The "clean" fixes carry real behavior risk.** The `x || !x ? x : false` lines
  are latent no-ops: `x || !x` is always truthy, so the `: false` default is dead
  and `undefined` is preserved. Rewriting them to `x ?? false` (the tidy way to
  satisfy the flag) would change output (`undefined → false`). The only zero-risk
  alternative — widening the internal interfaces back to `| undefined` — just
  re-admits the present-but-`undefined` state the flag is meant to catch, so it
  enforces the constraint only at the public boundary and is largely cosmetic.

Net: the flag fights the library's own normalization design, and closing 112 sites
buys little while risking output regressions. Revisit only if the chart/shape
option-normalization code is ever refactored to separate "raw input props" from a
"resolved options" type — at which point `exactOptionalPropertyTypes` on the input
type becomes cheap and genuinely useful.

### The `as`-cast enforcement asymmetry — ✅ resolved

Gap 1/3 guardrails repeatedly say "fix with real narrowing, **not `!` or `as`**."
The `!` half is enforced by `no-non-null-assertion`; the `as` half is now covered
by `@typescript-eslint/no-unnecessary-type-assertion` (arrived with Gap 2's
`recommendedTypeChecked` and pinned explicitly in the rules block).

Note the two rules gate *different* things: `no-non-null-assertion` bans **every**
`!`, whereas `no-unnecessary-type-assertion` only removes casts the compiler can
prove are redundant. A cast that genuinely changes the type — e.g. `number as Emu`
(branding a plain number into the `Emu` nominal type) or `unknown as T` — is by
definition *not* redundant, so it is correctly left in place. The two adjacent
`Math.round(...)` image-sizing lines in `gen-xml.ts` illustrate this: `cy` is a
plain `number` so its cast was dropped, while `cx` is typed `Emu` (its initializer
`getSmartParseNumber()` returns `Emu`) so `as Emu` is load-bearing and stays. The
lint gate distinguishes the two automatically. `consistent-type-assertions` remains
an open option if we later want to constrain the *form* of the intentional casts
that survive.

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

Gaps 1, 2, 3, and 4 are done, and `no-unsafe-enum-comparison` (the last open
static-check item) is now enforced. All static-check gates are green with every
type-flow rule on. Nothing remains.

## `no-unsafe-enum-comparison` — ✅ resolved

The `CHART_NAME`/`CHART_TYPE` unification done earlier removed the chart half of
the findings (60 → 15). The remaining 15 non-chart comparisons were all
enum-member-vs-string-literal checks that happened to be value-safe but were
opaque to the rule; each was rewritten to compare against the enum member so the
rule is satisfied structurally rather than suppressed:

- **`gen-utils.ts` (10)** — the 10 chained `colorVal !== SchemeColor.<member>`
  comparisons in `createColorElement`'s validation collapsed to a single
  `!Object.values(SchemeColor).includes(colorVal as SchemeColor)`, matching the
  existing scheme-color idiom in `gen-objects.ts` (`SCHEME_COLOR_NAMES`).
- **`gen-objects.ts` (1) / `gen-xml.ts` (2)** — `_type === 'placeholder'` string
  literals replaced with `SLIDE_OBJECT_TYPES.placeholder` (already the dominant
  usage in both files).
- **`measure-fit.ts` (2)** — `anchor === 't'` / `'b'` replaced with
  `TEXT_VALIGN.t` / `TEXT_VALIGN.b` (the field is typed `TEXT_VALIGN`).

The rule now comes on by default from `recommendedTypeChecked`; its explicit
`off` override was removed from `eslint.config.mjs`. Build, typecheck, lint, and
the full unit suite (406) stay green.
