---
doc-schema-version: 1
title: "Development guide"
summary: "Setup, source layout, generated outputs, and contribution rules."
read_when:
  - Setting up the repository
  - Changing source layout or generated output policy
  - Updating development commands
doc_type: "guide"
---

# Development guide

## Prerequisites

- Node.js `>=24`.
- `pnpm` through Corepack. The `packageManager` field in `package.json` pins the version.

```bash
pnpm install
```

`pnpm install` runs `prepare`, which installs the git hooks and builds `dist/` when it is missing.
[CONTRIBUTING.md](https://github.com/shbernal/ts-pptx/blob/master/CONTRIBUTING.md#git-hooks) lists
the hooks.

## Repository layout

- `src/`: TypeScript source.
- `test/`: regression tests, schema fixtures and validator helpers.
- `dist/`: the built package. Tests import from it, not from `src/`. Do not hand-edit it unless
  the task asks to refresh release artifacts.
- `docs/`: documentation content. See [Site changes](#site-changes).
- `www/`: the site's theme and Vue components, including the demos page.
- `demos/`: the showcase decks (`demos/showcases`) and the Node streaming demo (`demos/node`).
- `scripts/`: build, gate, package and demo automation.
- `skills/`: the `ts-pptx-upstream` skill, which ships in the package. Skills for working on this
  repository are under `.agents/skills/`.
- `tools/api-docs/`: TypeDoc with a pinned TypeScript 6. The root compiler is TypeScript 7, which
  ships no JavaScript compiler API. TypeDoc needs that API, and so does
  `scripts/raw-xml-ratchet.mjs`, through the `typescript-6` alias in the root devDependencies.
  Keep both pins. `tools/api-docs/README.md` says when they can go.

Keep source changes in `src/` and tests in `test/`. Leave unrelated uncommitted changes in the
working tree as you found them, and do not revert changes you did not make.

Never add a file whose name differs from an existing one only by case. The repository is developed
on case-insensitive filesystems, where the second file overwrites the first.

`.gitattributes` stores and checks out every text file with LF line endings, whatever
`core.autocrlf` says. If `format:check` reports every file on Windows, the checkout has CRLF
endings. Check the files out again instead of running `format`, which would rewrite all of them.

## Source conventions

The OOXML generators live under `src/gen/`. `gen/define/*` normalizes user options onto the slide
model, and `gen/{drawingml,slide,pres,opc,chart,table,anim}/*` serialize the model to OOXML. Two
comment conventions let a reader find their way around a module:

- **File module map.** Each module opens with a TSDoc block that states its job. A larger file also
  lists its regions there.
- **Region banners.** A file-level section starts with a one-line banner. Its name matches the
  module-map entry, so grepping `===== <name> =====` finds it:

  ```ts
  // ===== Region Name =====
  ```

  Use one banner per group of related functions, not one per function, and never one inside a
  function body. Sequential steps inside a function are `// STEP N:` comments.

When you add a top-level region, add the banner and its module-map line together. Write banners in
ASCII `=====`, not box-drawing characters. The `// ── Name ──` sub-headers inside some enums, such
as `TableStyle` in `src/enums.ts`, group members of one construct. They are not file regions.

New emitter code builds XML with `el()` from `src/gen/oxml/el.ts`, not template strings.
[Emitting XML: the `el()` builder](ooxml.md#emitting-xml-the-el-builder) has its rules.

### Trailing `_` marks an escape hatch

A public member whose name ends in `_` (today that is `element_` across the read model) is a
deliberate escape hatch onto the internal representation. The underscore makes hatch use easy to
grep and easy to spot in review, in this repository and in consumer code.

Do not rename such a member to the bare name. That turns a flagged hatch into ordinary-looking API.
Every `element_` has a public `markDirty()` on the same object, because a live DOM node handed out
without that obligation is how an edit vanishes on save. The
[Escape hatches](scope-and-policy.md#escape-hatches) section of the scope policy says when a new
hatch is acceptable.

## Common commands

`pnpm run` lists every script. The table covers the ones you use while working. The
[gate matrix](testing.md#gate-matrix) shows what each aggregate, git hook and CI job runs.

| Command | What it is for |
| --- | --- |
| `pnpm run verify` | The per-change check: typechecks, source and docs checks, and every Vitest suite. |
| `pnpm run verify:full` | Before pushing, and for a package or release change: `verify` plus the site build, the script round trip, and the package and size gates. |
| `pnpm run check:static` | CI's static job: `lint`, `lint:chars` and `format:check`, then `check:core`. |
| `pnpm run check:core` | The cheap checks that `verify` and `check:static` share. |
| `pnpm run check:package` | A package boundary change: `package:lint`, `test:package` and both size gates. |
| `pnpm run test` | Every Vitest suite. `test:unit`, `test:read` and `test:schema` run one part of it. |
| `pnpm run watch:dev` and `pnpm run test:watch` | The edit-then-test loop, one per terminal. `watch:dev` rebuilds `dist/` without declarations. |
| `pnpm run typecheck` | The `src/` TypeScript project. `typecheck:scripts`, `typecheck:test` and `typecheck:site` cover `scripts/`, `test/` and the site. |
| `pnpm run lint`, `pnpm run lint:chars`, `pnpm run format:check` | The git hooks run these for you. `lint:chars:fix` and `format` apply fixes. |
| `pnpm run build` | Build `dist/` for its own sake. Gates build it when they need it. |
| `pnpm run test:browser` | The Playwright lane in Chromium. Run `pnpm exec playwright install chromium` once first. |
| `pnpm run docs:dev` | Serve the site with hot reload. |
| `pnpm demos:build` | Build the showcase decks. |

The [testing guide](testing.md) covers the single-purpose scripts, such as `coverage:probe`,
`test:com`, `byte-identity:check` and the freeze commands.

Pass a flag straight after the script name: `pnpm run lint --fix`. pnpm forwards a `--` to the tool
as a literal argument, so `pnpm run lint -- --fix` hands oxlint a `--` followed by `--fix`. oxlint
and oxfmt drop everything after the `--` without an error, so the command exits 0 and applies no
fix. If a flag seems to have had no effect, read the command line pnpm echoes.

## Building and typechecking

The aggregates, and every script that reads `dist/`, start with `scripts/ensure-dist.mjs`. It
rebuilds `dist/` only when `src/`, a build config, `package.json` or the lockfile is newer than the
output. Do not prefix a command with `pnpm run build &&`.

A green `build` is not evidence that the types are correct. tsdown's `.d.ts` pass does not
typecheck, so `const x: number = 'a-string'` still builds. Only `typecheck` catches it. Never
substitute one for the other.

## TypeScript strictness

Strictness is configured once in `tsconfig.base.json` and applies to all of `src/`. Beyond `strict`
and `strictNullChecks`, it enables `noUncheckedIndexedAccess`, `noPropertyAccessFromIndexSignature`,
`exactOptionalPropertyTypes`, `verbatimModuleSyntax`, `noImplicitReturns`,
`noFallthroughCasesInSwitch`, `noImplicitOverride`, `noUnusedLocals` and `noUnusedParameters`. Fix
a new error with real narrowing or a guard. A `!` assertion and a redundant `as` cast are both lint
errors (see [Lint policy](#lint-policy)).

### Absent versus present-but-`undefined`

`exactOptionalPropertyTypes` is the one strictness flag that asks a design question rather than
catching a slip, so the answer is written down here. A `foo?: T` declaration says the key is either
**missing** or holds a `T`. A key that is present and holds `undefined` is a third state.

**Option bags spell "unset" exactly one way: an absent key.** They get stored, spread and
enumerated: a layout placeholder's options onto a slide's, a column default under a cell's own, a
combo subchart's overrides onto the chart's. A spread decides on whether the key *exists*, not on
what it holds. So a normalizer that rejects a value removes it.

`src/options-internal.ts` has the two helpers for that. `setOrClear` is for a write-back.
`pickDefined` is for a literal that projects a key list off a bag that may not state them all. Its
module header is the long form of this rule. It is the write-side twin of `compact()` in
`src/script/from-read/values.ts`, which keeps the same invariant on the read side, for the same
reason: two IRs describing one deck must not compare unequal because one wrote `{ bold: undefined }`
and the other wrote `{}`.

**Three kinds of declaration say `| undefined`, and each says why where it is written.**

- A property backed by a **class accessor**, such as `Slide.background` or `Slide.transition`, is
  always *present* on a real slide, and writing `undefined` is how its setter is cleared. There is
  no absent state to describe.
- A **read-only argument bag**: a parameter its call sites assemble inline out of values they may
  not have, such as `relationshipEl`'s `opts` or `genXmlTitle`'s title settings. The reader consults
  it with `?.` and nothing spreads it, so the two states really are the same to it.
  `MaybeUndefined<T>` in `src/types/internal.ts` is the mapped type for this, and its doc comment
  says when *not* to reach for it.
- A **measurement record** built in one place and read in one place (`FitRun`, `FitParagraph`),
  where the builder produces every key on every record.

**One bag deliberately carries both states**, and it is the reason the rule is worth keeping
everywhere else: `ChartOptsOverrides`, a combo subchart's vetted option overrides.
`gen/chart/chart-xml.ts` merges it *over* the chart-level options, so a present `undefined`
suppresses a chart-level value where an absent key inherits it. Suppressing is the intent: the
subchart supplied an override the schema rejects, and the chart-level value is one it never asked
for. `resolveSubchartOptions` applies that distinction at the merge and drops what is left, so the
plot builders still receive an ordinary bag.

`tsconfig.test.json` sets the flag `false`, because TypeScript rejects it without
`strictNullChecks` and the test project turns that off (see the comment there).

## Lint policy

- `pnpm run lint` runs oxlint with `--deny-warnings`, so a warning fails the gate.
- `.oxlintrc.jsonc` sets `options.typeAware`, so the type-aware rules run, through
  `oxlint-tsgolint`. The option lives in the config rather than a CLI flag so an editor's oxlint
  integration reaches the same verdict as `pnpm run lint`.
- `src/` runs the full type-aware set. `scripts/` and `test/` turn it off except for
  `no-floating-promises` and `no-misused-promises`.
- `typescript/no-non-null-assertion` bans a bare `!`, and `typescript/no-unnecessary-type-assertion`
  bans a provably redundant `as`. They are pinned by name as a pair, so the ban on both ways around
  null safety survives a change to oxlint's presets. A deliberate `unknown as T` cast is not
  redundant.
- `eslint/no-console` is an error in `src/`. [Errors and diagnostics](#errors-and-diagnostics) says
  what to use instead.
- Every rule turned `off` carries its reason beside it in `.oxlintrc.jsonc`, and the file's header
  records why the baseline is the `correctness` category plus rules named one by one. Read both
  before enabling a category.
- oxfmt is the only formatter. oxlint ships no formatting rules, so there is nothing to disable and
  no compatibility package to add.
- `format:run` in `package.json` names the files oxfmt formats. The pre-commit oxfmt job in
  `lefthook.yml` selects staged files by extension. The two lists are maintained separately, so a
  file type added to one goes into the other in the same commit.
- `lint:chars` runs charcheck, which rejects the em dash and the horizontal bar in the prose that
  `charcheck.config.js` names. `lint:chars:fix` rewrites findings, where you can read the diff
  first. When the gate itself looks wrong, use the `charcheck-upstream` skill in `.agents/skills/`.

## OOXML changes

Before changing emitted OOXML, read [OOXML agent context](ooxml.md). It sets the order for looking
up schema and PowerPoint behavior, and says what counts as evidence for a change. A change to
emitted XML carries a focused fixture in `test/schema-cases.js`. The [testing guide](testing.md)
covers schema validation and `pnpm run test:schema`.

## Package boundary changes

The package ships one ESM build. A change to package exports, generated filenames or package
contents must keep the support contract in [where it runs](../getting-started/runtime.md). Run
`pnpm run check:package`. When `test:package` fails, [`test:package` fails](#testpackage-fails)
lists the usual causes.

## Demo changes

The demos are showcases. No gate builds them, and the [testing guide](testing.md) says what covers
the published package instead.

```bash
pnpm demos:build                        # both showcase decks
pnpm demos:build quarterly-review       # one, by slug
pnpm --dir demos/node run demo-stream   # streams a deck over HTTP
```

The browser version of the quarterly review deck is the site's demos page, which
`pnpm run docs:dev` serves.
[demos/README.md](https://github.com/shbernal/ts-pptx/blob/master/demos/README.md) says what each
demo is for.

## Site changes

One VitePress build publishes the project site at `https://shbernal.github.io/ts-pptx/`: the front
page, the docs and the demos page. It spans two trees:

- `docs/` is content: markdown under the frontmatter schema, navigated from `docs.json` and
  validated by `docs:check`.
- `www/` is the code that renders it: the VitePress theme, its stylesheet and the Vue components a
  page mounts. See [www/README.md](https://github.com/shbernal/ts-pptx/blob/master/www/README.md).

VitePress looks for a theme only at `<root>/.vitepress/theme`, so `docs/.vitepress/theme/index.ts`
is a one-line re-export of `www/theme`.

```bash
pnpm run docs:dev       # hot-reloaded, at http://localhost:5173/ts-pptx/
pnpm run docs:build     # the build CI publishes, with the docs check before and after it
pnpm run docs:preview   # serve the built output
pnpm run typecheck:site # tsc over www/**/*.ts and docs/.vitepress/**
```

### Repository-only pages

`repoOnly` in `docs/docs.json` names `docs/contributing/`. Its pages keep the frontmatter schema and
the docs gates, but the site does not build them. People read them on GitHub.

- A served page links to one by its GitHub URL,
  `https://github.com/shbernal/ts-pptx/blob/master/docs/contributing/<page>.md`. `docs:check`
  rejects a relative link, which would be dead on the site.
- A repository-only page links relatively to any page under `docs/`, and by GitHub URL to a file
  outside `docs/`. It never links a site route such as `/reading/`, which GitHub cannot resolve.

### Diagrams

A fenced block whose language is `mermaid` renders as a diagram on the site, through
`www/diagrams/`. GitHub renders the same fence, so a repository-only page can use one too. No gate
parses a graph. A graph that does not parse shows its parse error in place of the diagram, so look
at the page under `pnpm run docs:dev` before committing it.

### The demos page

- **It is a test fixture.** The Playwright `demo` project drives `/demos`. It is the only place
  `src/runtime/browser.ts`'s `writeFile` executes, and where the browser-built deck is compared byte
  for byte with the Node-built one. The specs find elements by ARIA role:
  `getByRole('group', { name: 'Download' })`, then a button matching `/^Build /`, then
  `role="status"` or `role="alert"` inside that group. Rearrange the markup freely, but keep those
  roles.
- **Nothing typechecks a `.vue` file.** `tsc` does not read single-file components, and this
  repository has no `vue-tsc`. So the page's logic lives in `www/demos/deck-preview.ts`, which
  `typecheck:site` reads and `test/regression/www/deck-preview.test.js` covers. The component is
  markup around it. Logic added to the component escapes both.

## Errors and diagnostics

[Errors and warnings](../errors-and-warnings.md) is the contract a consumer sees. This
section covers adding a throw or warning site.

### Adding one

- Name the condition in `src/codes.ts` first. `warn` and `warnOnce` take a `DiagnosticCode`,
  and each error class constructor takes its own code union, so a site whose condition has
  no code does not compile.
- Reuse an existing code when the condition is the same, even when the wording differs or
  another entry point raises it. A condition that can arrive under two error classes needs
  two codes, because the class is part of what a consumer branches on.
- Do not add an error class per throw site. The five classes are the whole taxonomy, and the
  code carries the specificity.
- Report a warning through `warn` or `warnOnce` from `src/diagnostics.ts`. Use `warnOnce`
  for a condition that can repeat across a deck, and keep the offending value in its message
  so that a different value still reports.
- Write a warning message without a `ts-pptx:` prefix, because the default handler adds one.
  Write an error message without an `ERROR:` prefix, because the class name labels it.
- Do not report a condition with `console.log`, `console.warn` or `console.error`. A console
  line cannot be captured, silenced or branched on, and oxlint rejects it under
  `eslint/no-console`. `.oxlintrc.jsonc` exempts three files: `src/diagnostics.ts`, which
  owns the default handler, and the two `verbose: true` table tracers,
  `src/gen/table/autopage.ts` and `src/gen/table/html-dom.ts`, whose output reports no
  condition and is opt-in.

### Warn or throw?

Ask what the library does next, not how bad the input looks.

- Warn when the library can carry on and still produce a deck the caller would recognise.
  Clamping the value, ignoring the option and falling back to a default all qualify.
- Throw when the request is dropped and the deck would come out without what was asked for.
  A deck that opens and is missing an image is worse than a failed build.

`addImage()` and a picture bullet sit on either side of this line. An image with neither
`data` nor `path` has nothing to place, so `addImage()` throws `image/missing-source`. A
bullet image whose `data` lacks a base64 header warns `bullet/image-missing-base64-header`
and draws a default bullet glyph.

Dropping a request and reporting only a warning is the combination the rule excludes. The
caller reads a warning and gets a deck that ignored the option.

### An out-of-range number

`clampRangedInput` in `src/units-internal.ts` applies this rule to a number whose OOXML
attribute has a range, such as a percentage, a size or an angle. Route a new ranged option
through it.

| Value | Answer | Why |
| --- | --- | --- |
| Finite and out of range | Clamp to the nearest bound and warn | A legal value sits next to the one asked for. `shadow.transparency: 120` paints at 100, `bullet.size: 500` draws the glyph at 400%, `fit.fontScale: 150` scales at 100. |
| `Infinity` or `-Infinity` | Clamp to the nearest bound and warn | It has a nearest bound like any other out-of-range number. |
| `NaN`, or a value that is not a number, such as `'50'` | Throw `InvalidOptionError` | No legal value is nearest, and clamping would write `val="NaN"` into the part. |

The throw carries `percent/non-finite` unless the caller passes its own code, such as
`shadow/angle-non-finite`. Rejecting a finite out-of-range value and emitting nothing is not
a third answer. It drops the request and reports only a warning.

### An empty colour string

A paint has three states, and each has one spelling: omit the option, `{ type: 'inherit' }`,
or `{ type: 'none' }`. `''` is not a fourth. It comes from the caller's own missing value,
such as an unset template field. The library reports it under `color/empty-string`, then
resolves it the way omitting the option resolves.

What omission resolves to depends on the option:

| Option | `''` and omission both resolve to |
| --- | --- |
| a text box's `fill` | `<a:noFill/>` |
| a shape's `line.color` | the shape line default |
| a chart's `dataLabelColor` | `DEF_FONT_COLOR` (`000000`) |
| a slide's `background` colour | the layout's background |
| `bullet.color` | no `<a:buClr>` element |

A gradient stop and a duotone half require a colour, so they have no omitted state. `''`
there paints `DEF_FONT_COLOR` and reports the same code.

`rejectEmptyColor` and `namedColorOr` in `src/gen/drawingml/color.ts` implement the rule.
Call one before the site's own fallback, so the diagnostic names the option that carried the
empty string. `createColorElement` handles the slots that require a colour. Painting a
default for `''` would put black on a shape whose caller expected the theme's paint.
Emitting nothing would make `''` differ from omission wherever omission means `<a:noFill/>`
or a stated default.

## When a check fails

Start from the check that failed and the layer it tests. These three are the ones whose
failures do not explain themselves. User-facing symptoms, such as a deck PowerPoint offers
to repair, are on [Troubleshooting](../troubleshooting.md).

### `test:package` fails

`test:package` packs the tarball, installs it with npm and pnpm, and resolves every export
subpath through the installed `exports` map, the way a consumer does. No suite under `test/`
does that, because those import from `dist/` by path. The [testing guide](testing.md) lists what
it asserts. The usual causes:

- A public export was renamed or removed. The TypeScript consumer that
  `scripts/package-smoke.mjs` generates still uses it, and no unit test notices.
- A new subpath is missing from `files`, from the `exports` map, or from `EXPORT_MATRIX` in
  `scripts/package-smoke.mjs`.
- A retired upstream artifact is back in the tarball. The `assertNoFile` calls in
  `scripts/package-smoke.mjs` name each one.
- A top-level `await` reached an entry's chunk graph, so `require()` of that entry throws.

A declaration that does not resolve fails `package:lint` first: it runs publint and
`@arethetypeswrong/cli` over the packed tarball.

### `test:schema` fails, or PowerPoint rejects a deck

The usual causes:

- The emitted XML is structurally invalid.
- PowerPoint treats a structure differently from the schema.
- A change moved package parts or relationships without a matching fixture.

To work it:

1. Read [OOXML agent context](ooxml.md) before changing the emitter.
2. Take the diagnostic's `id`, `description`, `partUri` and `xpath` to the `ooxml` MCP server's
   `ooxml_explain`, which answers what was legal at that position. The
   [testing guide](testing.md) describes what a validation error carries.
3. Add or update a focused fixture in `test/schema-cases.js`, and iterate with
   `pnpm run test:schema`.

A deck can pass the schema and still fail in PowerPoint, with `0x80070570` or a dropped
shape. `pnpm run test:com` opens decks in PowerPoint over COM to catch that. It runs only on
Windows with PowerPoint installed, and CI does not run it. The [testing guide](testing.md)
covers what counts as render evidence.

### `docs:check` fails

`docs:check` runs `docs:api` first, which regenerates `docs/reference/api/` with TypeDoc, then
`scripts/docs-check.mjs`. The second step reports:

- a missing or empty frontmatter field, or a `doc_type` outside the allowed list;
- an unbalanced code fence;
- a `docs.json` navigation entry with no page, or a page in no navigation group;
- a broken relative link or site route, or a relative link that leaves `docs/`;
- a relative link from a served page into `docs/contributing/`, which the site does not
  build. The message prints the GitHub URL to use instead;
- a link whose `#anchor` names no heading on the page it points at, including one pointing
  within the same page.

`docs/reference/api/` is gitignored and rebuilt on every run, so an edit there is lost. A
wrong API page comes from the TSDoc in `src/`: fix it there.

An anchor is resolved against the headings the target page actually defines, plus its
explicit `{#id}` and `<a id>`/`<a name>` targets, so a heading rename that leaves a link
behind is an error rather than a silent dead link. The slug is generated the way the page
that *links* it is read: the site's scheme for a served page, GitHub's for one under
`docs/contributing/`, which is read on GitHub and never built into the site. The two differ,
so the same heading can have two spellings and each link is checked under the one that will
be used.
