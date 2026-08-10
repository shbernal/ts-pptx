# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed

- **BREAKING (write): `TextPropsOptions.strike` admits `'noStrike'`, and `underline.style`
  matches `ST_TextUnderlineType`** (#14). `strike` was `boolean | 'sngStrike' | 'dblStrike'`,
  which had no spelling for the explicit off even though the serializer passes any truthy
  string straight to the attribute — so `strike: 'noStrike' as 'sngStrike'` already produced
  the right XML and only the published type refused it. It is now
  `boolean | 'noStrike' | 'sngStrike' | 'dblStrike'`. The breaking half is `underline.style`:
  `'dotDashHeave'` was a typo for `'dotDashHeavy'` and is corrected, and the missing
  `'words'` is added, so the union is the enumeration's full 18 members. Migration: replace
  `underline: { style: 'dotDashHeave' }` with `'dotDashHeavy'` — the old spelling was not a
  legal `ST_TextUnderlineType` value, so any deck written with it carried an invalid `@u`.

  The doc comments now also say which state silence is an alias for, on all three
  decorations. `false` and an omitted `strike` both write no attribute and therefore state
  *nothing*, leaving the run with whatever it inherits — the same as `bold`/`italic`, whose
  falsy arm is likewise an omission. `'noStrike'`, `underline: { style: 'none' }` and
  `caps: 'none'` are the spellings that state "off" and override an inherited decoration.

- **BREAKING (read): `Shape.slide` is now `Shape.host`, typed `ShapeHost`.** A shape proxy
  no longer belongs to a slide specifically — the same `p:sp` can sit in a slide's,
  a layout's, or a master's `p:spTree` — so the back-reference names what it actually
  is. `ShapeHost` is the small contract all three classes satisfy (`part`, `partName`,
  `opc`, `relationships`, `themeContext()`, `shapeByIdDeep()`) and is exported from
  `ts-pptx/read`. Migration: `shape.slide` → `shape.host`; where you genuinely need the
  `Slide`, narrow with `shape.host instanceof Slide`. `Slide` gains an `opc` getter
  (`=== presentation.opc`) and `SlideMaster`/`SlideLayout` gain public `opc` and
  `relationships` getters, all to satisfy that contract. Nothing else about a shape
  changed, and `Slide.shapes` is untouched.

- **The `ts-pptx-upstream` skill covers the far end of the cycle, not just the filing.**
  It ended at "write the workaround", which is the half that happens on its own — the
  half that rots is the release landing and nobody finding the stopgaps it retired. A
  new step 7 is that sweep: `npm view` for what is out, closed issues, `rg 'ts-pptx#'`
  for every marked stopgap here, then per stopgap bump, run the check the comment names,
  delete, leave a test behind, and comment upstream with the check that now passes. It
  rests on the marker actually saying something, so step 6 no longer prescribes
  `remove once fixed upstream` — a line that cannot be checked at bump time and sends
  the reader back to re-derive the reproduction — but the wrong output as an observable,
  the exact check that proves the fix, and the code the stopgap becomes. `ts-pptx#` is
  named as the literal token to grep for, which is what makes step 7 a command rather
  than a memory. Two other gaps a downstream consumer found by working through it: step
  4 now checks `npm view @shbernal/ts-pptx version` *before* the tracker, since a report
  against a stale pin costs a maintainer a full triage and ends in a close; and step 1
  carries the `FidelityNote.cause` triage — `unread` and `unwritable` are gaps worth
  filing, `unsupported` is the format's own limit and is not — which is our own
  machine-readable verdict on whose bug a silent loss is, and was being re-derived in
  each consumer's notes instead of read off the note.

- **The README installs the skill the way an agent will actually run it.** The documented
  line was the interactive one, in a section addressed to agents: it prompts for the
  skill and the runtimes, which unattended is a hang rather than an install. The
  non-interactive form is now the default shown, with the interactive one as the aside,
  and `--all` is called out — it installs into every runtime the CLI knows, around
  seventy, and leaves an `agent/` directory at the consumer's repository root. Also
  written down: the installed copy is a copy, so a version bump does not update it, and
  `skills experimental_install` restores the file from `skills-lock.json` without
  creating any runtime link, so it is a record rather than a restore command.

- **`docs/RELEASING.md` closes the loop the skill now depends on.** Issues here close
  when the fix merges, which is right for this repo and the wrong signal for a consumer:
  merged-and-unreleased lasts weeks, and a workaround deleted on the strength of a closed
  issue breaks against the version actually installed. Post-publish now comments the
  carrying version on each issue the release closed. `AGENTS.md` gains the rule #10 was:
  when a fix gives an option value a meaning it did not have, work out what its *absence*
  now means, because an emitter that defaults through a ternary has already assigned one
  state to silence.

### Added

- **`ts-pptx/script`'s standalone tier rebuilds a layout's decoration instead of dropping
  it.** A source layout became a `defineSlideMaster` call carrying a title and a
  background, and nothing else: the bands, rules, wordmarks, triangles and quote marks
  that make a deck recognisable as somebody's template were declared lost as
  `master.decoration` and thrown away, so a standalone script produced a deck that
  rendered its slides faithfully and wore a different suit. Each of a layout's
  non-placeholder shapes is now transcribed into that layout's
  `defineSlideMaster({ objects })` array, by **the same mapper the slides go through** —
  so a rectangle on a layout is decided by the code that decides one on a slide, and the
  two cannot drift. Fills, gradients, scheme tokens, custom geometry, adjust handles,
  rotation, effects and text all come across on the shape mapper's existing terms.

  The note claiming this was unwritable was wrong about which half was holding it, and
  measuring it is what showed that: `{ shape: { type } }` accepts any preset, `custGeom`
  included, and its `ShapeProps` carries fill, line, shadow and adjust handles, so most
  of the supposed write-side ceiling was already reachable. Three kinds genuinely are
  not, and each is handled rather than assumed away. A **group** has no variant, so it is
  flattened into its children — visually lossless, because `absoluteFrame` composes the
  group's offset, rotation, flips *and* child-space scaling into each child, and noted as
  `layout.group` because one selectable object becomes several. A **connector** has no
  variant either and is re-authored as a `line` preset: it paints the identical stroke,
  keeps its rotation (which the slide-side `addConnector` mapping loses), and matters
  more than it sounds — PowerPoint's line tool authors a `p:cxnSp`, so connectors are 18
  of the 45 shapes the fixture corpus's layouts actually draw. A **table** is dropped,
  with a `layout.decoration` note naming the shape.

  Layout **placeholders** are still not emitted, and that is unchanged and deliberate:
  the write path seeds every slide with each layout placeholder it did not populate, so
  re-declaring one would put an empty ghost shape on every slide bound to the layout.
  `master.decoration` survives for a **master's** own shapes and moves from a reading
  problem to a structural one — `defineSlideMaster` creates a *layout* under the single
  shared master, so a master's shape tree has no write-side counterpart to receive them.

  Fidelity notes recorded against a layout shape are namespaced under a new
  `LAYOUT_NOTE_PREFIX` (`layout.`), exported from `ts-pptx/script`, because the shared
  mapper speaks the slide vocabulary: `layout.line.width` is `line.width` seen from the
  chrome. The prefix is load-bearing twice over — the template-anchored tier suppresses
  every note under it (it rebuilds no layout, so none of them describes its output), and
  the round-trip check refuses to let one excuse a difference on a *slide*, which a
  shape name repeated between a layout and the slides bound to it would otherwise do.
  `isKnownNoteConstruct` is exported alongside it and resolves a prefixed construct to
  its slide entry in the coverage table.

- **`SlideMaster.shapes` / `SlideLayout.shapes` reach a template's own content, and
  `showMasterSp` says whether to draw it** (#12). Both classes exposed `placeholders`
  and nothing else, so the `p:sp`/`p:pic`/`p:graphicFrame` under a master's or layout's
  `p:cSld/p:spTree` — the bands, rules, logos and footer furniture a deck is recognized
  by — had no modeled path out of the read API at all; only `part.dom` reached them,
  which is the raw-XML hatch rather than the model. On a corpus of PowerPoint-authored
  decks this was the single largest read gap: a consumer walking `slide.shapes`
  reproduces the deck's content and none of its identity. Both getters return the same
  `AnyShape` union `Slide.shapes` does, from the same `buildShapes` dispatch, so
  shape-walking code applies unchanged — and the members carry the full paint surface
  (`resolvedFill`, `resolvedLine`, `presetGeometry`, `rotation`, `absoluteFrame`), which
  the smaller `Placeholder` class never had. Tokens resolve against the *owning* part's
  context: a master shape's `schemeClr accent2` goes through the master's own `p:clrMap`
  and theme, not a slide's. Groups recurse and compose to slide-absolute frames as at
  slide level, and both classes get `shapeByIdDeep`.

  `placeholders` is unchanged, and is now documented as the filtered view of the same
  tree — both hand out the same live `p:sp` elements. Read a placeholder there to
  *place* it, through `shapes` to *draw* it.

  Shipping the accessor alone would have traded one wrong answer for another, so
  `Slide.showMasterSp` and `SlideLayout.showMasterSp` ship with it: `@showMasterSp`
  (ECMA-376 attributeGroup `AG_ChildSlide`, `xsd:boolean` defaulting to `true`, so absent
  means shown) is how a slide or a layout suppresses the master's decorative shapes —
  PowerPoint writes it on section dividers and full-bleed layouts. Without it a consumer
  that gained access to master shapes would paint them onto slides that deliberately hid
  them. Both are read-only; the write API authors neither. The layout arm has a genuine
  oracle in `mixed.pptx` and `read-stress.pptx`, which each carry `showMasterSp="0"` on
  their title layout.

- **`fill: { type: 'inherit' }` authors a shape whose interior comes from the style
  reference or the placeholder** (#10). 3.1.0 gave `type: 'none'` its `<a:noFill/>` back
  (#9) and, in doing so, left the *inherit* state with no spelling on the shape and text-box
  path: that emitter defaults a **missing** `fill` to `<a:noFill/>`, so omitting the option
  is an explicit transparent interior, not silence. The two arms were therefore both
  no-fill, and 3.1.0's own migration advice — "omit `fill` instead" — pointed at the wrong
  one. `'inherit'` emits no `EG_FillProperties` member at all, which is the state that lets
  `p:style/a:fillRef` or the placeholder paint the shape. The default is untouched: every
  existing `addShape`/`addText` call without a `fill` keeps emitting `<a:noFill/>`, which is
  why this is additive rather than a fix to the ternary. `ShapeLineProps` inherits the new
  member, where it means the same thing the stroke side already got for free — omit the
  paint child, keep the `<a:ln>`. On a table cell and a slide background, where *omitting*
  the option already meant inherit, `'inherit'` is simply the explicit spelling of that.

- **Any commit is installable straight from GitHub: `npm i github:shbernal/ts-pptx#<sha>`.**
  It looked like this already worked, and it never did. `dist/` is gitignored and `prepare`
  only installed git hooks, so a git-URL install packed a tarball in which every `exports`
  entry named a file that had never been built — the install succeeded and the first import
  failed. `prepare` now also runs `scripts/ensure-dist.mjs --if-missing`, a new mode that
  builds an *absent* `dist/` and leaves a stale one alone. That distinction is the whole
  design: the existing freshness check would be wrong in `prepare`, where it would make
  `pnpm run build` and `pnpm run watch` build twice over, and every script that needs a
  current build already front-loads its own unconditional `ensure-dist`. Absent means
  build; stale is somebody else's question. Two things had to move with it, both only
  reachable once consumers could run this path at all: `ensure-dist` invokes the build
  through `npm_execpath` — the package manager actually running it — rather than a
  hardcoded `pnpm`, which is declared here as `packageManager` but never installed as a
  dependency and so resolved to a shim a plain-npm consumer does not have; and
  `install-hooks.mjs` now skips when `INIT_CWD` places the caller outside the checkout,
  which is what a consumer's install looks like. Without that it ran `lefthook install`
  inside the package manager's throwaway clone and propagated lefthook's exit status into
  someone else's `npm install`. This is for trying an unreleased fix, not for production
  dependencies: the install builds from source, so it pulls this package's
  `devDependencies` and takes minutes.

### Fixed

- **`readModelToIr` carries a text decoration's explicit off** (#14). A run stating
  `a:rPr/@u="none"`, `@strike="noStrike"` or `@cap="none"` was read correctly by
  `Run.underline` / `Run.strike` / `Run.caps` and then mapped to `undefined`, so the
  resulting `CallIr` carried no option and re-emitting wrote no attribute. All three now
  carry — `underline: { style: 'none' }`, `strike: 'noStrike'`, `caps: 'none'` — and only an
  *absent* attribute maps to an absent option.

  `none` / `noStrike` are not the same fact as stating nothing. Each is a member of its own
  enumeration (ECMA-376 §20.1.10.81, §20.1.10.78, `ST_TextCapsType`) and would be redundant
  with omission if omission were the only way to be off. It is not, because run properties
  resolve down the `a:lstStyle` → placeholder → layout → master chain: a run that would take
  `u="sng"` from its list style and states `u="none"` is not underlined, and the same run
  with the attribute dropped is. So the loss was invisible on a deck with no inherited
  decoration and a visibly wrong answer on one that has any.

  It was undeclared either way — the same shape as #13. `DeckIr.fidelity` named neither
  construct, and `canonicalDeckIr` did not carry the field, so `diffDeckIr` compared two
  models that were *both* missing it and reported the deck clean; a consumer's round-trip
  harness structurally could not see this. Carrying the tokens closes both at once, and
  neither state notes. Two PowerPoint-authored fixtures state these tokens: `mixed.pptx` and
  `table.pptx` hold 132 runs stating `u="none"` and `strike="noStrike"`, 100 of which also
  state `cap="none"`. No new note fires, so `script:census` is unmoved.

- **`readModelToIr` keeps a baked `normAutofit`'s `fontScale` and `lnSpcReduction`** (#13).
  The mapper flattened every `normAutofit` frame to `fit: 'shrink'`, so both numbers were
  gone by the time a consumer held the `DeckIr` and everything downstream re-emitted a bare
  `<a:normAutofit/>`. Neither end of the round trip was at fault — `TextFitShrinkProps`
  already carries both fields and `TextFrame.autofitFontScale` already reads them back — so
  this was a mapping omission with an exact fix available: a frame that bakes either
  attribute now emits the object form `fit: { type: 'shrink', fontScale, lnSpcReduction }`,
  and one with neither keeps `fit: 'shrink'`.

  The two spellings are two states, not one value at two precisions. ECMA-376 §21.1.2.1.3
  defaults each attribute to 100%/0% only when it is *omitted*; PowerPoint recomputes an
  unbaked scale on edit and draws a baked one exactly as written until then. A deck baked at
  `fontScale="40000"` therefore came back painting its text two and a half times too large,
  in a file valid either way, which is why nothing caught it.

  It was also silent, which is the worse half. `printScript` named no fidelity note, so a
  consumer following the documented rule — trust the tier's own notes — was told the tier
  lost nothing; and `canonicalDeckIr` did not carry the fields either, so `diffDeckIr`
  compared two models that were both missing them and reported the deck clean. A round-trip
  oracle built on `diffDeckIr`, which is what the docs recommend building, could not detect
  this class of loss at all. Carrying the numbers through the IR closes all three at once.

  One arm still loses something, and now declares it: the write path rejects a percentage
  outside 0–100 and drops the attribute with a warning, so a source outside that range falls
  back to bare `'shrink'` with a `text.autofit.fontScale` / `text.autofit.lnSpcReduction`
  note (`dropped`/`unwritable`) rather than passing through a number that would vanish. No
  corpus fixture is malformed, so both read 0/44 and the census is unmoved.

- **Chart area and plot area fills honour `type` instead of only `color`** (#11).
  `ChartPropsFillLine.fill` is typed `ShapeFillProps`, so both areas looked like they took
  every fill kind a shape does — and `c:spPr` really is `a:CT_ShapeProperties`, the same
  optional `EG_FillProperties` group, so nothing in OOXML said otherwise. But both emitters
  gated on `fill?.color`, so every spelling that carries no colour fell to the `<a:noFill/>`
  arm and did nothing: `type: 'gradient'`, `type: 'pattern'`, and `type: 'inherit'` the
  moment it was added above. They now go through the shared fill dispatch, so a chart area
  can take a gradient or a pattern, `'none'` states a transparent area explicitly, and
  `'inherit'` emits no fill child at all and leaves the area to the chart style.

  Two spellings deliberately still mean no-fill, because the gate is on a fill being
  *stated* rather than merely present. `normalizeChartOptions` defaults `plotArea.fill` to
  `{}`, so every chart ever authored arrives at the emitter carrying a fill object — a
  presence check would have painted all of them a default grey. And `{ transparency: 50 }`
  with no colour is not a fill: there is nothing for the alpha to apply to. That was a
  documented `@example` on the option and has never worked; it now reads
  `{ color: '696969', transparency: 50 }`.

  `type: 'image'` remains unavailable on a chart and now warns
  (`image-fill/unresolved-media`) instead of silently doing nothing: a blip fill needs a
  media relationship on the chart part, and only shape and slide-level fills register one.
  No existing deck changes — the only inputs whose output moved are ones that used to emit
  `<a:noFill/>`, an invisible area nobody asked for.

## [3.1.0] - 2026-08-09

### Added

- **The package ships a `ts-pptx-upstream` skill, and `InternalError` says where to send a
  report.** Most code that calls this library is written by an agent working in a repository
  that is not this one, and an agent that hits a library defect writes a workaround instead of
  filing — which is the rational move from where it stands, since the workaround unblocks its
  user today and the tracker belongs to a repo it is not in. So the report never happens and the
  next consumer rediscovers the same defect. Three layers answer three different reasons the
  report dies. `InternalError` now appends the tracker URL from its constructor rather than from
  its throw sites, so a site added later cannot forget it; it is the only class that does, because
  it is the only one that already declares whose bug it is, and a "report this" banner on every
  malformed package would train callers to skip the line that always means something. The
  `ErrorCode` TSDoc — and therefore the `.d.ts` an agent in `node_modules` actually reads — now
  states the test for the other four: the supported bar is *"the output opens cleanly in Microsoft
  PowerPoint"*, and it reads in both directions, so a `PackageReadError` on a file PowerPoint opens
  cleanly is our gap, not bad input. And `skills/ts-pptx-upstream/` is published in the tarball, so
  `npx skills add ./node_modules/@shbernal/ts-pptx` works offline and always matches the installed
  version. The skill's load-bearing instruction is the one about the deck: presentations carry
  client names, unreleased strategy and pricing, and the tracker is public, so it spends most of
  its length on reducing a failure to a script that builds its own deck — and passes
  `--repo shbernal/ts-pptx` on every `gh` call, since `gh` in a consumer repo would otherwise file
  our bug into theirs. A third issue form, `agent-report.yml`, is where the error message's URL
  lands; its attachment dropdown deliberately has no option for a file containing real data, and
  "what should have happened" asks for the reason — an ECMA-376 clause, PowerPoint's own behaviour,
  or the docs — since that is what separates an actionable report from a matter of opinion. No API
  changed. `InternalError.message` gained a trailing pointer, which is not a contract: the class and
  the `code` are API and the message never was.

- **`TableCell.fillNoFill` reads an explicit `<a:noFill/>` on a cell** (#7) — the cell-side
  counterpart of 3.0.0's `Shape.fillNoFill`, and what `TableCell.noFill()` has always been
  able to write. `hasOwnFill` is not this question: it is `true` for any
  `EG_FillProperties` child, so it cannot separate a suppressed fill from a gradient, a
  pattern, a picture or an `a:grpFill`, and every colour accessor (`resolvedFill`,
  `fillColor`, `fillSchemeColor`) reports `null` for a no-fill cell exactly as it does for
  one inheriting the style's shading. Deriving it as "has an own fill that no accessor
  recognises" also changes meaning silently the day a further fill kind gets an accessor.

### Changed

- **Development toolchain: ESLint + Prettier → oxlint + oxfmt, and TypeScript 6 → 7.**
  This is a contributor-facing change with **no runtime effect on the published package** —
  no API was added, removed or altered, and the generated OOXML is byte-identical across
  all 183 parts of the reference decks. Enforcement was held level rather than relaxed: of
  the 90 lint rules previously enabled on `src/`, 89 carry over (the missing one, `no-octal`,
  is already a syntax error in an ES module), type-aware linting stays on, and
  `no-floating-promises` / `no-misused-promises` now also cover `scripts/` and `test/`,
  where the old configuration could not afford them. See `docs/development.md`.

- **Published `.d.ts` files are textually different, though no type changed.** TypeScript 7
  prints declarations slightly differently: string literal types use single quotes
  (`readonly shapeType: 'autoShape'` rather than `"autoShape"`) and redundant parentheses
  are dropped (`fontRef?: (StyleFontRef | null) | undefined` becomes
  `fontRef?: StyleFontRef | null | undefined`). Chunk filename hashes shift as a
  consequence. No declaration, signature or export moved, and `publint`,
  `@arethetypeswrong/cli` and the packed-package smoke test all pass — but a consumer
  diffing shipped declarations between versions will see churn, which is why it is recorded
  here rather than left as an implementation detail.

  Note that documentation generation deliberately stays on TypeScript 6: TypeScript 7 is
  the native Go compiler and ships no JavaScript compiler API for TypeDoc to import. That
  copy is confined to the private `tools/api-docs` workspace package and reaches nothing
  that is published.

### Fixed

- **`fill: { type: 'none' }` emits `<a:noFill/>` instead of nothing at all** (#9). The
  option's own name states the shape is transparent, and it was the one call that did not
  produce that state: `genXmlColorSelection` had no `none` case, so the fill child was
  omitted entirely and the interior fell back to `p:style/a:fillRef` or the placeholder —
  a shape carrying a style reference rendered in the theme's accent colour rather than
  transparent. `line: { type: 'none' }` on the same options object has always emitted its
  `a:noFill`; the two now agree. This reaches every caller of the shared fill dispatch, so
  a table cell (`TableCellProps.fill`) can author a transparent cell the same way.
  Round-trip consequence: `addShape({ fill: { type: 'none' } })` → save → load now reports
  `Shape.fillNoFill === true`, where it reported `false` before. The one behaviour that
  changes for existing decks is that `type: 'none'` no longer produces the *inherit* state
  by accident — if you were relying on it to mean "leave the fill to the style", use
  `type: 'inherit'`, added in the next release. (This sentence originally said to omit the
  `fill` option, which is wrong on the shape and text-box path: a missing `fill` emits
  `<a:noFill/>` there. See #10.)

- **`readModelToIr` carries a line's `@cap`, and declares a dropped `@algn`** (#8). Both
  legs of the cap mapping already existed — `ShapeLineProps.cap` authors the attribute and
  3.0.0's `Shape.lineCap` reads it back — but the script tier's `lineOption` never consumed
  it, so a deck this library wrote could not survive its own converter. The loss was
  *undeclared*, which is the part that mattered: the round-trip gate excludes exactly what
  a fidelity note names, so it passed green while the deck changed. `@cap` extends every
  dash by the stroke width and decides whether each draws as a rectangle or a lozenge, so
  on a thick dashed rule the before and after are visibly different. `@algn` is readable and
  has no write option, so it now records a `line.align` note (`dropped`/`unwritable`) —
  for `algn="in"` only, since `ctr` is what an omitted `@algn` already renders as.

- **`latexToOmml` emits accents as `<m:acc>` rather than `<m:limUpp>`** (#6). `\hat`,
  `\bar`, `\vec`, `\dot`, `\ddot`, `\tilde`, `\acute`, `\grave`, `\check`, `\breve`,
  `\mathring`, `\H`, `\dddot` and their short-form aliases (`\^`, `` \` ``, `\'`, `\"`,
  `\.`, `\=`, `\u`, `\v`, `\r`) all landed as over-*limits*, with limit spacing and
  semantics, because temml emits a bare `<mover>` (correct for a browser — MathML renderers
  derive accent positioning from the operator dictionary) and mathml2omml has no dictionary
  and keys strictly off `accent="true"`. The pipeline now carries the small dictionary
  subset that closes the gap, and while it is there it swaps temml's *spacing* modifier for
  the combining mark ECMA-376 §22.1.2.20 says an `accPr` character should be — so `\vec{v}`
  gets an arrow accent instead of a full-size arrow hung over the base.

  Scoped to `latexToOmml`: `mathmlToOmml` passes hand-written MathML through unchanged,
  because there `accent` is the caller's to set. Constructs that were already mapping well
  are untouched (`\widehat`/`\overbrace` stay `m:groupChr`, `\overline`/`\underline` stay
  `m:borderBox`, `\stackrel` stays `m:limUpp`), and two stay limits by necessity: `\utilde`
  and other under-accents, since OMML has no under-accent object and the symmetric
  `accentunder="true"` makes mathml2omml emit an *over*-accent; and `\ddddot`, whose
  two-character operator has no single `m:chr`.

- **A table cell's explicit `a:noFill` survives read → script → write.** Previously a
  suppressed cell fell out of `cellFill` as "no fill option" and the copy took the table
  style's banding — the opposite of what the source showed. Enabled by the
  `TableCell.fillNoFill` reader above and the `fill: { type: 'none' }` writer fix above.

## [3.0.0] - 2026-08-09

### Added

- **`Shape.fillNoFill` reads an explicit `<a:noFill/>` on a shape's fill** — the fill-side
  counterpart of `lineNoFill`, and the only accessor that separates a deliberately
  transparent shape from one that inherits its fill through `p:style/a:fillRef`. Every other
  fill accessor (`fillColor`, `fillSchemeColor`, `resolvedFill`, `gradientFill`,
  `patternFill`, `pictureFill`) reports `null` for both, so a consumer honouring the read
  model painted an `a:noFill` rectangle in the theme's accent colour. The same class had
  `ChartFill.noFill`, `ChartLine.noFill` and `CellBorder.noFill` already; the shape fill was
  the omission. `Shape.noFill()` has always been able to *write* this state.

- **`Shape.lineCap` and `Shape.lineAlign` read `a:ln/@cap` and `a:ln/@algn`.** `@cap` was a
  write/read asymmetry inside this library: `ShapeLineProps.cap` authors it and nothing read
  it back, so a deck this writer produced could not round-trip through its own reader without
  losing an attribute the writer put there on purpose. Both report the raw OOXML token, the
  way `lineDash` reports `@val` — `'flat'` / `'rnd'` / `'sq'` and `'ctr'` / `'in'` — and
  `null` when unset rather than a defaulted value. On a thick dashed rule the cap decides
  whether each dash reads as a rectangle or a lozenge and changes the drawn length of every
  one; SVG's `stroke-linecap` is the exact equivalent.

- **`GroupShape.childFrame` reads a group's own child coordinate space**
  (`p:grpSpPr/a:xfrm/a:chOff` and `a:chExt`), as a `ChildFrame` of `offsetX` / `offsetY` /
  `extentX` / `extentY` in EMU, or `null` when the group has no transform. `absoluteFrame`
  reads these internally to compose slide-absolute geometry and remains the right answer for
  anything that *paints* — this is for a consumer that *rebuilds* a group as OOXML, which
  needs the source child space to reproduce its scaling and could otherwise only rebuild
  groups whose child space is the identity. Named after the OOXML attributes rather than the
  read model's `left`/`top`/`width`/`height`, because it is the source rectangle of a
  mapping, not a frame on the slide.

- **`clipPath()` names the clip silhouettes you would otherwise re-derive.** A `ClipShape`
  is data — a named silhouette plus its options — and `clipPath(shape, w, h)` resolves it to
  the freeform `points` path `addImage` emits as its `<a:custGeom>` clip mask. The first
  silhouette is the half-disc a cover-slide picture placeholder cuts. See
  [`docs/image-in-shape.md`](docs/image-in-shape.md).

  ```js
  import { clipPath } from '@shbernal/ts-pptx'

  const w = 5.22, h = 7.5
  slide.addImage({
    path: 'cover.jpg', x: 0, y: 0, w, h,
    points: clipPath({ kind: 'half-disc', flat: 'right' }, w, h),
    sizing: { type: 'cover', w, h },
  })
  ```

  `flat` names the edge the straight side sits on; `preset` picks the proportion, `'deep'`
  (the default) or `'shallow'`. Paired with `sizing: 'cover'` this is a standalone
  reproduction of what a picture placeholder does — a layout `custGeom` clipping an
  inherited blipFill — with no placeholder, and no layout, involved.

  **The box size is an argument for a reason.** A `custGeom` point written as `%` resolves
  against the *slide*, not the picture, so a box-relative silhouette has to be emitted in
  inches already scaled to its box. `clipPath` multiplies its fractions out at build time,
  which is what lets one silhouette scale to any region — and is why handing a path to a
  picture of a different size puts the clip somewhere else entirely. That trap is the whole
  reason this is worth shipping rather than leaving to each caller.

- **`slide.addModel3d()` embeds a 3D model** — PowerPoint's *Insert ▸ 3D Models*. A glTF
  binary (`.glb`) travels inside the package, and PowerPoint 2019+ renders it live and lets
  the viewer orbit it. See [`docs/3d-models.md`](docs/3d-models.md).

  ```js
  slide.addModel3d({
    path: 'assets/engine.glb',
    preview: { path: 'assets/engine-render.png' },
    meterPerModelUnit: 1 / 240, // the model's largest bounding-box dimension
    x: 1, y: 1, w: 6, h: 4,
  })
  ```

  Two things are worth knowing before using it:

  - **Supply `preview`.** Everything that is not PowerPoint 2019+ draws that picture instead
    of the model — including PowerPoint's own slide thumbnails, PDF export, and print. The
    library has no 3D renderer, so omitting it embeds a gray placeholder and emits a
    `model3d/preview-missing` warning. Same bargain as `addOleObject()`'s `cover`.
  - **Set `meterPerModelUnit`.** The `am3d` scene is measured in metres. PowerPoint reads the
    model's bounding box and normalizes its largest dimension to 1 metre; ts-pptx does not
    parse glTF, so it emits `0.5` (correct for a model 2 units across) and leaves the rest to
    you. Left at the default, a model 240 units across becomes a 120-metre object with the
    camera inside it. Set it to `1 / <largest bounding-box dimension>`.

  `camera` overrides the viewpoint (`pos`/`lookAt`/`up` in metres, `fov` in degrees); the
  defaults are the ones PowerPoint wrote for a 2×2×2 cube. Out-of-range and non-finite values
  throw rather than being coerced. Linked (non-embedded) models, animation scenes, and a typed
  read accessor are out of scope for now — a model read through `ts-pptx/read` surfaces as an
  inert `graphicFrame` and survives load → save and `importSlide` byte-intact.

- **Browser support is now proven in CI, not assumed.** A Playwright lane
  (`pnpm run test:browser`, CI job `browser`) runs the package in headless Chromium
  against `demos/vite-demo`. No library code changed — this converts an existing
  claim into evidence.

  The assertion worth naming is cross-runtime byte identity. The demo imports the same
  showcase module `pnpm demos:build quarterly-review` builds, and `src/zip.ts` pins
  `FIXED_MTIME`, so the two packages are directly comparable: all **113 parts** of the
  browser-built deck are byte-identical to the Node-built one. Every serializer, the zip
  writer, part ordering and relationship numbering are therefore runtime-invariant — by
  comparison, not by inspection. A second spec reads the object-URL download back with
  jszip (an implementation independent of the `fflate` the library writes with) to confirm
  it is a real OPC package.

  Two boundaries stay exactly where they were, and are now stated in
  [`docs/runtime-and-package-support.md`](docs/runtime-and-package-support.md) rather than
  left to inference:

  - **Runtime support is not layout fidelity.** Nothing in the lane depends on a rendered
    page. Real `offsetWidth`, the resolved cascade, and browser-chosen fonts remain out of
    active scope; a layout difference between two browsers is not a defect in this
    package's browser support, whereas a `.pptx` a browser builds differently from Node is.
  - **`tableToSlides` measurement is still unavailable without a layout engine**, exactly as
    documented — `offsetWidth` is `0`, widths fall back to computed CSS and then an equal
    split, and `data-pptx-width` / `data-pptx-min-width` pin them.

  The explode/normalize/diff machinery the byte-identity gate has always used moved to
  `scripts/pptx-parts.mjs` so both gates share one definition of "the same bytes"; the
  refactor was verified byte-identical against a baseline frozen with the pre-refactor
  script.

- **The whole `RuntimeAdapter` now runs in a real browser**, not just the download path.
  A second Playwright fixture serves the shipped `dist/browser.js` **unbundled** over a
  static server and drives decks written to reach the three loaders `demos/vite-demo`
  cannot, because its showcase draws every asset rather than loading one:

  - `loadMedia` — a fetched raster image lands in the package as the same bytes Node reads
    off disk, *and* as the same bytes as the source file. The two implementations return
    different strings for the same image (Node raw base64, the browser a `FileReader` data
    URI) and everything downstream reconciles them, including the image sizer.
  - `createSvgPngPreview` — the `<canvas>` rasterizer, whose branches nothing exercised
    before: a real PNG where Node can only stub a placeholder, plus the undecodable-SVG and
    zero-dimension arms, each of which must fail rather than ship a blank fallback.
  - `loadFontData` — a font fetched over HTTP bakes the same `fontScale` and embeds the
    same `/ppt/fonts/` bytes as one read off disk.

  The deck definitions are written once and built twice, once per runtime, so a divergence
  in the fixture cannot read as a divergence in the runtime.

  Two things fell out of loading the shipped file unbundled. `opentype.js` turns out to be
  a *dynamic* bare import inside the measure/fit chunk — invisible to every bundled
  consumer, and now documented for anyone loading the entry over a plain
  `<script type="module">`. And Node and the browser are *expected* to disagree on exactly
  one part: the SVG PNG fallback, where Node has no rasterizer. The lane asserts the shape
  of that disagreement so it cannot quietly become a different one.

- **`tableToSlides` is tested against a table a browser actually laid out** — a third
  Playwright project, `html-table`, on the same unbundled harness server.

  `pickColWidthBasis` chooses between three column-width bases, and its *first* arm — the
  rendered `offsetWidth` — had never executed anywhere. The Node suite drives happy-dom,
  where `offsetWidth` is `0` for every cell, so it always took a fallback arm; the unit
  suite reached the function by handing it numbers directly, which proves the `if` and not
  the pipeline behind it. So the primary path of the feature, including the `arrColSrc`
  arithmetic that fixed the spanning-`data-pptx-width` defect, was covered only at its own
  function boundary.

  The fixture is built so the measured basis and the computed-CSS basis **disagree**
  (`offsetWidth` is the border box, computed `width` the content box), because a test that
  only showed "the widths came out proportional" would be equally green if the measured arm
  never ran. The spec re-derives both bases from the live page and fails if they ever
  converge. Sensitivity-checked by disabling the arm: exactly the two arm-dependent
  assertions go red, reporting the CSS ratio.

  **This does not move the scope line, and the wording matters.** What is asserted is that a
  measurement is *taken and honoured* — proportionally, with `data-pptx-width` still winning
  outright. Nothing asserts that Chromium's numbers are the right numbers or that another
  engine would agree; that is live-DOM layout fidelity, it has no oracle, and it stays out
  of active scope. A layout difference between two browsers is still not a defect in this
  package; a `.pptx` a browser builds differently from Node still is — which the lane pins
  directly, by converting the same markup in both runtimes and asserting Node falls back to
  the CSS basis where the browser measures.

  The new project contributes its V8 coverage to the merge like every other browser spec, so
  the merged report moved up on all four axes — statements 93.91 → 94.03, branches
  83.71 → 84.16, functions 98.29 → 98.58, lines 96.11 → 96.20. Two notches in
  `scripts/coverage-gates.json` are ratcheted with it (statements 92 → 93, branches
  82 → 83), which is what keeps a gate from carrying two points of slack.

- **Coverage from both lanes is merged into one number** (`pnpm run coverage:gate`, CI job
  `coverage`). `scripts/coverage-merge.mjs` folds the browser lane's V8 coverage into the
  Node report on one rule — *the Node report defines the shape, the browser lane
  contributes hits* — so the merged denominator is identical to the Node report's and the
  two percentages are directly comparable. Merged: statements 93.91, branches 83.71,
  functions 98.29, lines 96.11.

  It also makes this repo's **point-of-slack rule fail a build** rather than live in prose:
  `scripts/coverage-gate.mjs` is red both when a number falls below its notch *and* when it
  clears it by less than a full point. Prose does not fail a build, which is how the
  exclusion drop left `functions` at 0.35 of slack while an acceptance criterion of
  "thresholds still pass" was satisfied.

  The four numbers in `vitest.config.ts` remain the Node suite's own floor, and the browser
  lane keeps its per-function gate — a percentage cannot say *which* adapter function
  stopped running.

- **The package is now bundled for Node and run, inside `pnpm run test:package`.** The
  browser lane put a real bundler in front of the `browser` condition; nothing asked the
  same of the `node` entry, and the two are different questions — Node's resolver finds a
  specifier on disk at call time, while a bundler must resolve every one of them, including
  dynamic ones, at build time. A package can be perfectly importable and still be
  unbundlable.

  `bundleForNode()` esbuild-bundles the installed tarball with `platform: 'node'` across
  every export subpath but `/browser`, then runs what it emitted. It fails if the build
  warns, if anything other than a Node builtin stayed external, or if the emitted bundle
  cannot write a `.pptx`. Both the npm and pnpm fixtures, since pnpm's symlinked store is a
  different shape for a bundler to walk.

- **A bundle-size budget for the browser entry** (`pnpm run bundle-size:check`, part of
  `check:package`). Nothing measured shipped size before; a size promise nobody measures is
  a promise that quietly stops being true. It fails only on a step change and asks for a
  re-freeze only when a win is worth banking, because bytes move on every commit. The
  figure is a growth detector, not a download size — `dist/` is unminified and every real
  browser consumer minifies it.

### Changed

- **BREAKING: `Paragraph.bullet` is replaced by `Paragraph.bulletDetail`.** The old accessor
  reported a *tagged string* — `'none'` / `'char:•'` / `'autoNum:arabicPeriod'` — which is
  ambiguous when the glyph is itself a colon and reads as a bare glyph if you do not know
  better. That is not hypothetical: this library's own script converter first consumed it as
  one, and `'none'.codePointAt(0)` put a literal `n` bullet on every converted deck, silently.

  `bulletDetail` returns a discriminated union with no parsing left to get wrong, and carries
  what the string could not — `a:buAutoNum/@startAt`, and the bullet's own `a:buFont` /
  `a:buSzPct` / `a:buSzPts` / `a:buClr`. It also reports a fourth kind the old accessor
  dropped to `null`: a picture bullet (`a:buBlip`), with its image part resolved.

  ```js
  // before
  para.bullet // 'autoNum:arabicPeriod' — startAt, font, size and colour unreachable

  // after
  para.bulletDetail
  // { kind: 'autoNum', scheme: 'arabicPeriod', startAt: 5,
  //   font: 'Wingdings', sizePct: 80, sizePt: null,
  //   color: 'C00000', schemeColor: null, resolvedColor: { … } }
  ```

  Migration: `bullet === 'none'` → `bulletDetail?.kind === 'none'`;
  `bullet?.startsWith('char:')` → `bulletDetail?.kind === 'char'`, with the glyph at
  `.char` rather than after the colon; `bullet.slice('autoNum:'.length)` → `.scheme`.
  A paragraph that inherits its bullet still reports `null`.

  Numbering is content rather than styling: a list continuing "5. Deploy" that came back as
  "1. Deploy" was a different slide, and `numberStartAt` was a pure write/read asymmetry —
  `addText` accepted it and nothing could produce it.

- **BREAKING (output): `<a:buSzPct/>` is emitted only when `bullet.size` is given.** It used
  to be written unconditionally, pinned to `val="100000"`, on every object-form bullet and on
  `bullet: true`. An explicit 100% is not the same as leaving it out — it *overrides*
  whatever bullet size the layout's or master's list style sets, so every bullet this path
  wrote silently forced its glyph back to full size. The same class of bug as the explicit
  `a:buNone` an omitted `bullet` emits, and invisible until `bulletDetail` gave the
  round-trip check something to see it with. An out-of-range `bullet.size` now warns and
  emits nothing rather than warning and pinning to 100%. Decks that want the old behaviour
  can pass `bullet: { size: 100 }` explicitly.

- **`sizing.w` / `sizing.h` are optional, and `sizing: { type: 'stretch' }` names what used to
  be nameless.** The emitter has always defaulted the fit box to the picture's own extent; the
  type demanded both anyway, so every `cover`/`contain` call restated `w`/`h` it had already
  supplied a line above. `sizing: { type: 'cover' }` is now the ordinary form, and passing them
  still means what it always did — a fit box deliberately different from the picture.

  `stretch` emits the plain `<a:stretch><a:fillRect/></a:stretch>` a raster already gets. It
  exists so the fill-the-box behaviour can be *asked for*, which is what makes the vector
  default below opt-out-able rather than a trap.

- **BREAKING: an SVG is placed at its own aspect ratio by default instead of being stretched
  to its box.** `addImage({ svg, w, h })` used to fill the box whatever the glyph's proportions
  were, so any icon with a non-square `viewBox` — a minority in every real icon set, and never
  the one you check — came out squashed. Every consumer's answer was the same wrapper that
  routes each call through `sizing: 'contain'`; that wrapper is now the library's default.

  ```js
  slide.addImage({ svg: icon, x: 1, y: 1, w: 3, h: 1 })                            // letterboxed, centered
  slide.addImage({ svg: band, x: 0, y: 0, w: 13.33, h: 0.4, sizing: { type: 'stretch' } }) // opt out
  ```

  **Scope, deliberately narrow.** Rasters are untouched: a photo's box is chosen for it, filling
  it is what PowerPoint does, and letterboxing every existing deck's pictures would be a change
  of a different order. A vector is different in kind — it *states* its ratio in a `viewBox`,
  and disagreeing with that statement is a defect rather than a layout choice. Nothing is
  emitted when the ratios already agree, so a square glyph in a square box produces the same
  bytes it always did, and an SVG carrying neither `viewBox` nor `width`/`height` stretches
  silently — no sizing was requested, so there is nothing to warn about.

  **Two related SVG fixes fall out of the same root cause** — the write path treated vector
  sources as unmeasurable long after `src/media/image-size.ts` learned to read a `viewBox`:

  - `{ svg, w: 4 }` derives its height from the intrinsic ratio (a 2:1 viewBox → 4in × 2in),
    where before an omitted dimension silently became 1 inch. Rasters have always done this.
  - `{ svg }` with **neither** dimension still falls back to 1 inch, and that is not an
    oversight. An SVG's user units are dependable relative to each other and merely
    conventional in absolute terms; treating them as 96-DPI pixels would insert a 24-unit icon
    as a quarter-inch object. The ratio is trusted, the magnitude is not.

  If you have a wrapper that adds `sizing: 'contain'` to every icon, delete it — the emitted
  XML is identical either way. If you were relying on a stretched vector, name it: `sizing: {
  type: 'stretch' }`.

- **`vitest.config.ts` no longer excludes anything of this repo's own from coverage.**
  The `dist/browser.js` / `dist/browser-*.js` entries are gone; the second never matched
  anything, because tsdown bundles the adapter *into* the entry. Dropping them took the
  measured functions figure 98.33 → 97.35 while actual tested-ness went up, which is the
  shape of an honest denominator: the Node suite cannot execute an adapter that needs
  `fetch`, `FileReader` and a canvas.

- **`@shbernal/ts-pptx/math` is Node-only by decision, not by accident.** It loads its
  optional peers through `createRequire`, which is what keeps `latexToOmml()` and
  `mathmlToOmml()` synchronous; the browser-compatible alternative is a dynamic `import()`
  that would make both async — a breaking change to every existing caller, for a use case
  nobody has raised. If a browser consumer turns up, the answer is an additional
  `/math/async` subpath, not a change to this one. Recorded in
  [`docs/runtime-and-package-support.md`](docs/runtime-and-package-support.md) so it is not
  re-litigated per release.

- **The docs called `tableToSlides`'s no-browser width path a *degradation*. It is a
  *fallback*, and the difference is not cosmetic.** Every user-facing statement of it — the
  README, `docs/project-target.md`, `docs/runtime-and-package-support.md`, `AGENTS.md`, the
  `/html` entry's own doc comment — said column widths "degrade to computed CSS", which
  reads as *the same answer, less precisely*. They do not measure the same box:
  `offsetWidth` is the border box and a computed `width` is the content box, so padding
  alone can put the two bases in different **proportions**. The `html-table` fixture is
  built to demonstrate exactly that — 1:1 measured against 2:1 from CSS on one table — so
  the repo had the fact and the docs contradicted it. Now stated wherever the fallback is
  described, with the remedy: pin the column with `data-pptx-width` where both runtimes
  have to agree.

  No behaviour changed, and deliberately so. Normalizing the CSS basis to the border box
  would need computed padding and border widths, which the DOMs that reach that arm need
  not resolve (a `%` padding computes to nothing usable without layout) — so it would
  converge the two only sometimes, and it would collapse the one discriminator the browser
  lane has, turning `table-widths.spec.mjs` back into a test that passes whether or not the
  measured arm ran. The reasoning is recorded on `pickColWidthBasis` itself.

- **The browser lane stays Chromium-only, deliberately.** The APIs in play (`fetch`,
  `FileReader`, canvas, object URLs, `<a download>`) are not where engines are known to
  disagree, and no divergence has been reported or observed. An engine gets added when
  there is something concrete to add it for — also written down rather than left as a
  default.

### Fixed

- **`readModelToIr` now emits its `table.rowAuto` note when *every* row is auto-height**, not
  only when some are. The guard excluded the all-auto case explicitly, and that case is both
  the more common one — a table authored with no explicit row heights has `a:tr/@h="0"` on
  every row — and a real loss: no `rowH` is emitted at all, so `addTable` divides the frame
  height evenly and three auto rows come back pinned to a third of the frame each. The table
  still looks the same; what is lost is the *implicitness*, which matters the moment someone
  edits a cell and expects the row to grow. A round-trip oracle gated on "nothing undeclared"
  is only as good as its note set, and this was a difference passing through undeclared.

- **`autoPage` let every continuation slide take one row more than fitted, so the last row
  hung off the bottom edge.** Affects `addTable(rows, { autoPage: true })` and
  `tableToSlides()` alike, and only tables whose cells carry top/bottom margins — which is
  every table converted from HTML, since cell padding becomes a cell margin.

  The pager charges each row its cells' top and bottom margins before deciding whether the
  row fits. On a page break it did that and then zeroed the accumulator, so the first row of
  each new page — and only that row — was placed for free. The page then filled to its budget
  as if it had that space, and it did not. The deeper the padding, the further the overflow:
  at 8px of cell padding a 60-row table paged `[10, 11, 11, 11, 11, 7]` where every full page
  had room for 10.

  The symptom is easy to miss because the error is *constant*: every generated page overflows
  by the same amount, so the pages agree with each other and disagree only with the first
  one. That is also why the existing continuation-slide regression stayed green — it compared
  continuation pages to each other.

  This closes upstream `gitbrent/PptxGenJS#1200`, filed against `tableToSlides` and long
  assumed to be a browser-layout question. It is not: the browser supplies column widths, and
  nothing the vertical arithmetic reads. `test/browser/table-autopage.spec.mjs` pages the
  same table in headless Chromium and on a DOM that renders nothing and asserts the two
  produce the same slides.

## [2.0.0] - 2026-08-05

### Removed

- **BREAKING: `Presentation.defineTableStyle()` and `TableProps.styleDrivenCells` are
  gone, along with the `TableStyleProps` / `TableStyleRegionProps` types.** A custom table
  style is unreachable markup in PowerPoint: it emitted well-formed, schema-valid XML that
  can never paint.

  PowerPoint resolves `<a:tableStyleId>` against its **own** table-style gallery and never
  reads a style definition out of the package. Measured by rendering in PowerPoint desktop
  16.0, not inferred from the schema:

  - a deck pointed at a **built-in** GUID that the package does not define renders
    correctly — so the gallery, not the part, is what is consulted;
  - a PowerPoint-authored deck with one style's GUID rewritten to a novel value in *both*
    `ppt/tableStyles.xml` and the slide, bytes otherwise identical, loses that table's
    styling entirely (black hairline grid on white) while its untouched neighbours keep
    theirs;
  - the same holds for a definition placed inline in `<a:tblPr>` as `<a:tableStyle>`, and
    for one nominated by the part's `def=`. Lifting a genuine PowerPoint-authored
    `<a:tblStyle>` block under a custom GUID does not help either, so the markup we emitted
    was never the problem.

  `styleDrivenCells` was actively harmful under that finding: it stood down the per-cell
  `border` and `color` defaults so a style region could take over, but no custom region ever
  paints, so it traded a correct grid for PowerPoint's default one.

  **Migration.** Style tables with direct formatting, which is what carried them all along:

  ```js
  // before
  const brand = pptx.defineTableStyle({
    name: 'Brand',
    wholeTbl: { border: { type: 'solid', color: 'D9D9D9', width: 0.5 } },
    firstRow: { fill: '1A2B3C', color: 'FFFFFF', bold: true },
  })
  slide.addTable(rows, { tableStyle: brand, hasHeader: true, styleDrivenCells: true })

  // after
  slide.addTable(rows, {
    hasHeader: true,
    border: { type: 'solid', color: 'D9D9D9', width: 0.5 },
    headerRow: { fill: { color: '1A2B3C' }, color: 'FFFFFF', bold: true },
  })
  ```

  Banded rows move into the row data (set each row's `fill` as you build it) or into
  `columns[i]` for vertical banding. `TableProps.tableStyle` still exists and still works —
  it is now typed as `TableStyle` alone, since only a built-in GUID renders.

  `ppt/tableStyles.xml` is still emitted, as a bare stub naming a default style id, because
  PowerPoint expects the relationship and content-type override. The diagnostics
  `table-style/region-overridden`, `table-style/missing-argument`, `table-style/missing-name`
  and `table/style-driven-cells-inert` are removed with the API.

  **The read side is unaffected.** `Table.resolvedStyle`, `TableCell.resolvedFill` and
  `importSlideMasters({ tableStyles })` still resolve style graphs out of imported decks —
  those definitions were written by PowerPoint, and PowerPoint honours its own.

### Fixed

- **`tableToSlides` cell padding is now converted from px to inches.** A cell's
  computed CSS `padding-*` was read in px and assigned straight to
  `TableCellProps.margin`, which is **inches**. The magnitude therefore passed
  through unscaled: a perfectly ordinary 4px padded cell emitted
  `marL="3657600"` — a **4 inch** text inset, wider than most columns — and any
  cell padded 1px or more also tripped the `margin/legacy-points` warning. The
  stale `px->pt 1:1` note it was written under had been true when cell margins
  were points; it was not true after margins became inches.

  Padding now resolves at 96px/in, so `padding: 4px` is `4/96in` (`marL="38100"`).
  96 rather than 72 because CSS defines the reference pixel as 1/96in and this
  conversion exists to mirror what the browser laid out; it is also the density
  the `"<n>px"` coordinate unit already uses, so the two px sites agree. The
  whole-px rounding is gone with it — a fractional computed padding keeps its
  precision, and the rounding happens once, in EMU.

  **This changes emitted bytes for any padded HTML table**, on the browser path
  as well as `@shbernal/ts-pptx/html`. Cells will look substantially tighter,
  because they are now inset by what was asked for. Nothing else about the
  conversion changed, and a table whose cells set no padding is unaffected.

- **`fitColumns: 'shrink'` measured against the wrong slide margin.** The space to the
  right of a table is bounded by the **right** margin; the calculation subtracted
  index 3 of the TRBL margin tuple, which is the **left** one. Invisible with the
  default symmetric `[0.5, 0.5, 0.5, 0.5]` — and wrong by the difference for any master
  whose `margin` is asymmetric, which is exactly the layout someone sets a left gutter
  on. The existing tests all used symmetric margins, so none of them could fail; the
  new one uses `[0.5, 0.25, 0.5, 2]`, where the two answers differ by 1.75in.

- **A styled table cell's own fill is no longer dropped by `pptx-to-script`.** A cell
  with an explicit `a:solidFill` inside a table that also has a `tableStyle` replicated
  as *unfilled*, because the mapper could not tell that colour apart from one the cell
  merely inherited from the style's header/banding rules — and baking an inherited
  colour in would freeze the banding, so it dropped both.

  `TableCell.hasOwnFill` (new, read side) tells them apart: an `a:tcPr` either carries
  an `EG_FillProperties` child or it does not. `TableCell.resolvedFill` already
  branched on exactly that internally; the flag simply was not exposed. A cell's own
  fill is now carried and an inherited one is still left to the style GUID, which
  reproduces the banding exactly. Neither case loses anything, so the `table.cell.fill`
  fidelity note is gone rather than merely narrowed.

  Measured on PowerPoint's own `table-cell-image-fill.pptx`: the red cell keeps its red.

### Added

- **Table editing on `ts-pptx/read`.** The read proxies were read-plus-text-edit only:
  `TableCell.text` was the sole setter, and every other change needed the `element_`
  escape hatch plus a manual `markDirty()`.

  Cell properties: `setAnchor`, `setVerticalText`, `setHorzOverflow`, `setAnchorCtr`,
  `setMarginsEmu`, `setBorder(edge, …)` (the four edges and both diagonals),
  `setFillColor`, `setFillSchemeColor`, and `noFill()`. Structure on `Table`:
  `addRow`, `removeRow`, `addColumn`, `removeColumn`, `mergeCells`, `unmergeCell`.
  Each mutates in place and marks the part dirty, matching the `text` setter.

  Every insertion respects the `CT_TableCellProperties` **sequence**. That is the
  whole hazard: an append-only setter produces an out-of-order `a:tcPr`, which
  PowerPoint reports as a corrupt file rather than as a bad edit, and which no getter
  would notice. A schema-validation case now authors a deck, edits it through these
  setters, saves it, and validates the result — the only shape of test that catches it.

  Structural edits keep the grid rectangular and every merge's continuations in step.
  Inserting a row or column *through* a merge extends it rather than splitting it;
  removing a merge origin promotes its first continuation, so the region survives one
  row shorter; removing a column inside a merge drops a covered cell rather than the
  origin, so the region keeps its content. `mergeCells` **rejects** a rectangle that
  cuts through an existing merge instead of silently widening it to fit.

  Unlike the write path, an invalid value here **throws** rather than warning and
  dropping — new codes `table/invalid-cell-anchor`, `table/invalid-cell-vert`,
  `table/invalid-cell-overflow`, `table/invalid-cell-margin`,
  `table/invalid-cell-border`, `table/row-index-out-of-range`,
  `table/column-index-out-of-range`, `table/merge-range-invalid`. On the write path a
  bad option comes from a deck being built, and dropping one value beats failing the
  build; here it comes from a caller editing one attribute, and doing nothing silently
  would leave them looking at an unchanged deck with nothing to explain it.

- **`TableProps.tableFill`** — the table's own background, written as a real `a:tblPr`
  fill that the cells sit on top of. The existing `TableProps.fill` is *stamped onto
  every cell* instead, so nothing ever reached `a:tblPr`. The two usually render alike,
  which is why the difference is worth stating: with `fill` there is no such thing as
  an unfilled cell, so a cell can never fall back to a table background — and a deck
  read back from PowerPoint carries the `a:tblPr` shape, not the flattened one.
  `fill` is unchanged (changing it would repaint every existing deck); both JSDocs now
  say which is which.

  Takes the same `ShapeFillProps` a cell does — solid, gradient, pattern or picture.
  Read back via `Table.resolvedFill`, `.pictureFill`, `.gradientFill`, `.patternFill`
  and `.fillSchemeColor`, the same five a cell has.

  No table-level **effect** surface: PowerPoint's UI exposes none, so a source deck
  will not contain one and there would be nothing to reproduce.

- **`TableCell.gradientFill` / `TableCell.patternFill`** (read) — a cell's `a:gradFill`
  and `a:pattFill`. `TableCell.resolvedFill` reports `null` for every non-solid choice
  by design, and it was the only fill accessor a cell had besides `pictureFill`, so a
  gradient- or pattern-filled cell was indistinguishable from an unfilled one.

  That was not only a reading gap. `pptx-to-script` had nothing to fall back on, so a
  gradient cell **replicated as an unfilled cell** — or, when the table had a style, as
  whatever banding colour the style graph resolved to. Both now round-trip. Writing
  them always worked; it is now documented and pinned by tests, and the output is
  confirmed to open in desktop PowerPoint.

- **`TableCellProps.diagonal`** — a cell's corner-to-corner rules
  (`a:tcPr/a:lnTlToBr` / `a:lnBlToTr`), PowerPoint's "Diagonal Down/Up Border". The
  read model has always decoded them; there was no way to write one, and
  `pptx-to-script` dropped them with a note. Kept off `border`'s tuple deliberately:
  widening that to six entries would break every existing caller for a rare feature,
  and the diagonals are not edges. A merged region draws its diagonal **once**, on the
  span origin — covered cells inherit the origin's edges but never its diagonals,
  because a diagonal is one corner-to-corner stroke and repeating it per covered cell
  would draw a sawtooth.

- **`TableCellProps.anchorCtr`** — centres a cell's whole text *block* horizontally
  (`a:tcPr/@anchorCtr`), independent of each paragraph's `align`. The two compose:
  `align` places each line inside the text block, `anchorCtr` places that block inside
  the cell. Read back via `TableCell.anchorCtr`. `false` is the schema default and
  emits nothing.

- **`TableCellProps.cell3D`** — a 3-D bevel on a cell (`a:tcPr/a:cell3D`): preset,
  width/height in points, `prstMaterial`, and an optional light rig. Niche —
  PowerPoint's table UI has no control for it, so it reaches a deck from a theme or
  another producer — but it round-trips through PowerPoint verbatim, so authoring and
  replicating one now works. Read back via `TableCell.cell3D`.

  Two schema constraints show through the API. `a:bevel` is required, so `cell3D: {}`
  still emits a bevel rather than an empty (invalid) `a:cell3D`. `a:lightRig` requires
  **both** `rig` and `dir`, so a half-specified rig is reported and dropped whole. Any
  value outside its `ST_` union is reported as the new **`table/invalid-cell3d`** and
  dropped. The four enums are exported as `BevelPresetType`, `PresetMaterialType`,
  `LightRigType` and `LightRigDirection`.

- **`TableCell.id` / `TableCell.headerIds`** (read only) — a cell's `a:tc/@id` and the
  header cells associated with it (`a:tcPr/a:headers/a:header/@val`), which is how a
  complex table tells a screen reader what a value means.

  **There is deliberately no write-API counterpart.** PowerPoint opens a deck carrying
  both without complaint and then strips them on the first save, so an emitter would
  ship a feature that dies as soon as anyone edits the deck. The measurement is
  `test/read/fixtures/authoring/probe-table-cell-a11y-and-3d.ps1`, and it is a
  controlled one: `a:cell3D` and `a:headers` were injected into the *same* `a:tcPr`,
  and PowerPoint kept the first and discarded the second. `TableProps.hasHeader`
  (`a:tblPr/@firstRow`) remains the header marker PowerPoint keeps — and the one its
  own accessibility checker reads. The accessors exist because a deck from another
  producer may still carry the association; `pptx-to-script` records its loss as
  `table.cell.headers`.

- **`BorderProps.dashType`** — the exact `a:prstDash` preset for a border, using the
  same vocabulary as `ShapeLineProps.dashType`. `BorderProps.type` is only a coarse
  three-way switch, so every dashed border it can express — dotted, long-dash,
  dash-dot — collapsed onto the single `sysDash` preset on write *and* on read-back.
  A deck whose table borders are `dot` or `lgDashDot` could not be authored or
  replicated. `dashType` names the preset directly and wins over `type` when both are
  set; `type: 'none'` still suppresses the border before any dash is chosen. Honored
  by table cell borders.

  `ShapeLineProps.dashType` (and therefore `BorderProps.dashType` and
  `ConnectorProps.dashType`) now spans the **whole** `ST_PresetLineDashVal` set: the
  three values it was missing — `dot`, `sysDashDot`, `sysDashDotDot` — are accepted.
  This is a widening, so no existing value changes meaning. `pptx-to-script` maps a
  read dash straight through instead of flattening it, and a dash outside the enum is
  recorded as `table.cell.borders.dash` rather than silently approximated.

  An unrecognized `dashType` is reported as the new **`border/invalid-dash-type`**
  diagnostic and falls back to what `type` implies, rather than being written — a
  value outside `ST_PresetLineDashVal` would make the part schema-invalid, which
  PowerPoint reports as a corrupt file.

- **`TableProps.outerBorder`** — a border for the table's **perimeter** only: the top
  edge of the first row, the bottom edge of the last row, the left edge of the first
  column and the right edge of the last column. A single `BorderProps` boxes the
  table; a TRBL array with holes rules only the sides it names, leaving the others to
  whatever `border` already drew. Applied after every other border source, so
  "outline the table, no interior grid" is `outerBorder` with no `border` at all.

  The perimeter is decided by **grid position**, not by authored cell, so merges
  work: PowerPoint defines a merged region's outer edges on the *covered* cells, and
  a colspan reaching the last column gets that column's rule on its `hMerge` dummy.
  Leaving the option unset emits nothing — existing decks are byte-identical.

- **Vertical table cell text now survives `pptx-to-script`.** `TableCellProps`
  inherits `textDirection` from `TextBaseProps` and the emitter has always written it
  to `a:tcPr/@vert`, but the replication mapper recorded the attribute as unwritable
  and dropped it, so a replicated deck lost every vertical cell label for no reason.
  The four directions the option spells (`horz`/`vert`/`vert270`/`wordArtVert`) now
  round-trip; the `table.cell.vert` note is narrowed to the East-Asian
  `ST_TextVerticalType` modes that genuinely have no spelling.

- **`TableCellProps.horzOverflow`** (`'clip' | 'overflow'`) — controls what a table
  cell does with a **single glyph** wider than its text width, emitted as
  `a:tcPr/@horzOverflow`. `'clip'` (PowerPoint's default) cuts the glyph at the cell
  edge; `'overflow'` lets it draw past. It matters for oversized display type, wide
  CJK/emoji glyphs, and icon fonts in a narrow column. Read back via the new
  `TableCell.horzOverflow` accessor on `ts-pptx/read`, and carried through
  `pptx-to-script`.

  **It is not a text-wrap switch, despite where it sits.** That distinction is the
  reason this landed: the attribute had long been filed as the route to per-cell
  no-wrap, and it is not. PowerPoint has no per-cell no-wrap at all — `wrap="none"`
  on a cell's `a:bodyPr` renders inert and is stripped on the next save, and
  `TextFrame.WordWrap` is read-only on a cell over COM. Cell text always wraps to
  the column width. `test/read/fixtures/authoring/probe-table-cell-wrap.ps1`
  reproduces every part of that, and `table-cell-horzoverflow.pptx` is the
  PowerPoint-authored oracle for what does work.

  Writing `'clip'` explicitly is honored but redundant: it is the schema default, so
  PowerPoint drops the attribute the first time it saves the deck. Leaving the option
  unset emits nothing, so no existing deck's bytes change. An unrecognized value is
  reported as **`table/invalid-horz-overflow`** and dropped rather than written —
  `ST_TextHorzOverflowType` admits only those two, and PowerPoint reports a
  schema-invalid slide part as a corrupt file rather than as a mis-set option.

- **`border/unknown-key`** — a new diagnostic reporting a key that is not part of
  `BorderProps`, which was previously discarded without a sound. The thickness
  field is `width`, in points; a border authored with any other name for it lost
  the value and rendered at the 1pt default, and a `.pptx` gives no second signal
  — nothing throws and the deck opens, only heavier than asked for.

  TypeScript's excess-property check already rejects a stray key on a border
  written inline at the call site. It deliberately does not fire when the border
  is built as a variable first (`const b = {...}` then `{ border: b }`), because
  a variable may legitimately be a supertype — and that is precisely the reuse
  pattern a shared grid style encourages. The check closes that gap at runtime.

  It sits in `resolveBorderWidth`, the one function every emitted border resolves
  its width through, so it covers table cell borders, table-style regions, and
  chart borders alike. One exception: `chartArea.border` is rebuilt from a fixed
  key list during normalization, so an unknown key on it (and `cap` with it) is
  stripped before generation and cannot be reported. Emitted bytes are unchanged
  in every case — this only adds a report where a value was already being lost.

### Changed

- **`docs/tables.md`** — the table guide, which did not exist. Tables were the
  most-used object after text and the only major one with no prose doc. Covers the
  cell model, the styling precedence chain, borders (per-cell default vs. perimeter,
  dash styles, diagonals), fills, merges, sizing and auto-paging, the read/edit
  surface, and a "not authorable" section for the constructs PowerPoint discards.

- **`TableProps.border` is documented as what it is: a per-cell default, not the
  table's perimeter.** The old wording ("single value applied to all 4 sides / array
  in TRBL for individual sides") read like the outside of the table. It never was:
  `normalizeTableRows` broadcasts it to *every* cell, so
  `border: [solid, none, solid, none]` gives each cell a top and bottom rule — a full
  set of horizontal grid lines — rather than a rule above and below the table.
  Nothing about the behaviour changed; the doc now says so plainly and points at the
  new `TableProps.outerBorder` for the perimeter case.

- **A table-level `border` side with `width: 0` is now emitted as the hairline it
  asks for**, instead of being replaced by the 1pt default. Per-*cell* borders
  already treated `0` as a real width; only the table-level path used a truthiness
  test, so the two disagreed for exactly one value. Both now share one helper. A
  border that sets no `width` at all is unaffected.

## [1.0.0] - 2026-07-29

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
- Explicit ESM package boundary with typed subpath exports for `inspect`,
  `measure`, `read`, `math`, `node`, and `browser`.

- **`Run.charSpacingPt`** on the read model — character spacing (tracking) in
  points from `a:rPr/@spc`, the read counterpart of the write-side `charSpacing`
  option. It sat beside `caps`, `strike`, and `baselinePct` as the one run
  property the flat inspect surface could read and the deep model could not.

- **An error taxonomy: `TsPptxError` and five subclasses.** The library threw
  ~160 bare `Error`s and shipped no error classes at all, so a consumer wanting to
  tell *"you passed a bad coordinate"* from *"this font file is corrupt"* from
  *"these bytes are not a package"* had to match on message substrings — text the
  project is explicitly free to reword in any release.

  Every failure is now a `TsPptxError` carrying a stable `code`:

  ```ts
  import { MediaError, PackageReadError, InvalidOptionError } from '@shbernal/ts-pptx'

  try {
  	await buildDeck(spec)
  } catch (err) {
  	if (err instanceof MediaError) return retryWithPlaceholderAsset(spec)
  	if (err instanceof PackageReadError) return rejectUpload(err.code)
  	if (err instanceof InvalidOptionError) throw err // our bug — fail loudly
  	throw err
  }
  ```

  The classes are a deliberately flat set of five — `InvalidOptionError`,
  `UnsupportedFeatureError`, `PackageReadError`, `MediaError`, `InternalError` —
  answering *whose problem is this?*; the `code` carries the specificity.

  **The class and the `code` are API; the `message` is not.** Codes draw on the
  same vocabulary as diagnostics, so a condition reads the same whichever way it
  surfaces: `coord/non-finite` means the same thing thrown as an
  `InvalidOptionError` or warned as a `Diagnostic`. Each code belongs to exactly
  one class and the pairing is type-enforced —
  `new MediaError('coord/non-finite', …)` does not compile.

  Everything remains `instanceof Error`, so existing `catch` blocks are unaffected.
  The classes are re-exported from every entry point and resolve to one shared
  module, so `instanceof` works regardless of which subpath you imported from and
  which subpath threw. Where the library wraps a lower-level failure, the original
  is preserved on `cause` rather than flattened into the message. See
  [docs/errors.md](docs/errors.md).

- **A diagnostics seam: `setDiagnosticHandler`.** Library warnings were hardwired
  to `console.warn` across ~100 call sites, so a consumer generating decks in a
  batch job could neither silence nor route them, and nothing downstream could
  react to a *specific* condition without matching on message substrings.

  Every warning is now a structured `Diagnostic { code, message, detail? }`
  delivered to a handler you can install:

  ```ts
  import { setDiagnosticHandler } from '@shbernal/ts-pptx'

  setDiagnosticHandler((d) => logger.warn({ code: d.code }, d.message))
  setDiagnosticHandler(() => {}) // silence
  setDiagnosticHandler(null) // restore the console default
  ```

  **The `code` is API; the `message` is not.** A code identifies a *condition* in
  `area/condition` form (`'chart/non-finite-value'`, `'coord/bare-number-is-inches'`)
  and is stable: adding one is back-compatible, removing or renaming one is
  breaking. The wording behind it is free to improve in any release — do not parse
  it. `DiagnosticCode` is a closed union, so codes complete in an editor and a
  typo is a compile error.

  There is no separate strict mode: a handler that throws is one, and it composes
  with whatever policy you want.

  ```ts
  setDiagnosticHandler((d) => {
  	if (d.code === 'coord/bare-number-is-inches') throw new Error(d.message)
  })
  ```

  The handler is process-global rather than per-presentation. That is deliberate
  and documented — the emitting code is a tree of free functions with no
  presentation in scope — with the trade-off (concurrent builds cannot be told
  apart) written down in [docs/diagnostics.md](docs/diagnostics.md).

### Changed

- **Breaking (non-Node, non-browser runtimes): the bare `@shbernal/ts-pptx`
  import no longer resolves to the browser build.** `exports["."]` carries
  `node`, `browser`, and `default` conditions. The first two were right; the
  third pointed at `dist/index.js`, which did nothing but re-export
  `./browser.js`. So a runtime that sets neither condition — Deno, Bun, an edge
  worker — got the DOM adapter, and `writeFile()` tried to create an anchor
  element and click it in an environment with no `document`. For a documented
  Node-first project, the neutral fallback should not have been the browser.

  `dist/index.js` is now its own runtime-agnostic entry. Authoring is unchanged
  and everything that hands bytes back to you still works — `write()`,
  `stream()`, `toParts()` — which is the shape a worker wants anyway. Remote
  media and fonts still load, over `fetch`/`btoa`/`TextEncoder`, which every one
  of these runtimes has. Two things differ:

  - **`writeFile()` throws** an `UnsupportedFeatureError`
    (`runtime/file-output-unavailable`) naming `@shbernal/ts-pptx/node` and
    `@shbernal/ts-pptx/browser`, rather than failing on a missing `document`
    deep inside the call. There is no filesystem and no DOM here, so there is no
    destination to write to; take the bytes from `write()` and place them
    yourself.
  - **`tableToSlides()` is absent.** It resolves an element id against the
    global `document`, so it is defined on the browser entry alone. The
    DOM-agnostic form is the free `tableToSlides` on `@shbernal/ts-pptx/html`,
    which takes the element directly.

  Nothing changes for Node or browser consumers at runtime: those conditions
  already resolved to `dist/node.js` and `dist/browser.js` and still do.

  **Types now resolve through the same condition as the code.** `.` previously
  served one `dist/index.d.ts` to every condition, so a Node consumer was typed
  against the browser class and TypeScript accepted `pptx.tableToSlides(…)` on a
  build where it was `undefined` at runtime. Each condition now carries its own
  `types`, so what the compiler shows matches what the runtime has.

- **`ts-pptx/inspect` is now a projection over `ts-pptx/read`, and
  `fast-xml-parser` is no longer a dependency.** The library shipped *two*
  independent readers of a `.pptx` over two different XML parsers: the deep,
  navigable `read` model on `@xmldom/xmldom`, and the flat `inspect` snapshot on
  `fast-xml-parser` with its own hand-rolled, JSZip-shaped package facade. The
  overlap was near-total — boxes, rotation, group composition, runs, paragraphs,
  font size, colour, fill, line, shape type, wrap, autofit, body insets, part
  listing — so every read-side fix had to be made twice, and a divergence between
  the two was invisible to every test in the repo.

  The *shape* of `inspect` is unchanged: a cheap, flat, allocation-light snapshot
  is a genuinely different use case from a navigable model, and every exported
  type and function keeps its name and meaning. The *implementation* is gone.
  `inspect` now reaches the package through `OpcPackage` and every field through
  the read model's own getters, so the two surfaces cannot disagree about what a
  deck says.

  **`fast-xml-parser` is dropped from `dependencies`** (it stays a devDependency
  for one maintenance script), which is 1.4 MB less installed for every consumer
  of the package, whichever entry they import.

  Four behaviour changes come with it:

  - **Slides are reported in presentation order (`p:sldIdLst`), not part-name
    order.** Dragging a slide in PowerPoint rewrites that list and leaves
    `slideN.xml` named as it was, so the old directory-order enumeration reported
    the authoring history rather than the deck. `slides[].index` is now the deck
    position; `slides[].path` still names the part.
  - **`textRuns[].text` is verbatim.** `fast-xml-parser` trims text nodes by
    default, so every run came back stripped of the whitespace an
    `xml:space="preserve"` run carries. That both lost the leading/trailing space
    that widens a line and welded adjacent runs together — a slide reading
    `"This is test content."` inspected as `"Thisis testcontent."`, with a word
    count to match. Element-level `text` (runs joined, whitespace collapsed,
    trimmed) is unaffected in shape but now has the right word boundaries.
  - **`loadPptxPackage()` returns an `OpcPackage`**, not the JSZip-shaped
    `{ files, file(path) }` facade; the `PptxPackage` / `PptxPackageFile` types are
    removed. `listPptxParts()` / `readPptxTextPart()` / `readPptxBinaryPart()` keep
    their signatures and still speak zip paths. Migration: a caller that reached
    into `pptxPackage.file(path).async('string')` calls `readPptxTextPart(pkg,
    path)`, and one that wants more can now use the `OpcPackage` directly or hand
    it to `Presentation.fromPackage()` without re-reading the bytes.
  - **The input must be a real OPC package.** A zip holding slide XML but no
    `[Content_Types].xml` used to inspect fine; it now throws a `PackageReadError`
    (`package/not-an-opc-package`), the same bar `ts-pptx/read` applies.

  A run highlight authored as a theme token now resolves to a literal hex against
  the slide's theme instead of reading `null`.

  The cost is bundle size: a consumer importing only `/inspect` now pulls the read
  model's chunks (~440 KB of library code before their own tree-shaking) where it
  used to pull ~20 KB plus `fast-xml-parser`. This project is Node-first, and one
  reader that is right beats two that drift.

  Every other field of every element across all 43 fixture decks is byte-identical
  to the old implementation, pinned by a new characterization snapshot
  (`test/read/fixtures/inspect-surface.snapshot.json`).

- **`Shape.presetGeometry` moved from `AutoShape` to the `Shape` base**, so a
  picture or connector reports its preset geometry too. PowerPoint gives both one
  — a `p:pic` is `rect` unless cropped to a shape — and `Shape.adjustValues` was
  already on the base documenting itself as the companion to a member only
  auto-shapes had.

- **Nine reality-checks that wrote to the console now throw or warn.** A handful of
  validation sites in `gen/define/` reported a problem with a direct
  `console.error` / `console.log` and then carried on. That output could not be
  captured, silenced, or branched on — it predated the diagnostics seam and was
  never a `warn()` call, so the migration to `setDiagnosticHandler` did not see
  it. Each site now takes whichever surface fits what the library actually does
  next:

  - **`addImage()` throws** `InvalidOptionError` when the source is unusable —
    `image/missing-source`, `image/path-not-a-string`, `image/data-not-a-string`,
    `image/missing-base64-header`. **This is a behaviour change:**
    `addImage({})` used to print a line and silently omit the image, leaving a
    deck that opened fine and was missing content. There is nothing to draw, so
    it now rejects, matching `addMedia()` and the `hyperlink` check in the same
    function.
  - **`addTable()` throws** `InvalidOptionError` (`table/rows-not-nested`) for a
    row that is not an array of cells. It already rejected exactly this for row
    0; later rows only logged and pushed an empty row, quietly dropping content.
    **This is a behaviour change** for the later-row case.
  - **A picture bullet whose `data` lacks a base64 header warns**
    (`bullet/image-missing-base64-header`) rather than throwing: the run emitter
    falls back to a default glyph, so the deck is still valid. Behaviour is
    unchanged apart from the output now being routable.
  - **A malformed `hyperlink` reports once instead of twice.** Registration
    logged a line and declined to mint a relationship; the emitter then threw
    `hyperlink/not-an-object` / `hyperlink/missing-target` for the same input.
    The log is gone — the throw was always the real report.

  The `verbose: true` table tracer still prints to the console. It is a DEV-ONLY
  flag whose output reports no condition, and not passing it silences it.

- **Error messages no longer label themselves.** Fourteen messages carried an
  `ERROR: ` / `ERROR! ` prefix, four an `addMedia() error: ` one, and
  `coordToEmu`'s carried a literal `ts-pptx: `. The class name already labels the
  failure in every stack trace and console rendering, so all of them are gone.
  `presentation.layout = 'nope'` threw the literal string `UNKNOWN-LAYOUT`; it now
  names the value and says what to pass instead. The conditions and their codes
  are unchanged — only code matching on message text is affected, which the
  contract has never supported.

- **Warning output is now the default handler's job, not the message's.** Two
  messages carried their own prefix (`[WARNING] `, `Warning: `) and one carried a
  literal `ts-pptx: ` inside the text, which the console handler then doubled.
  All three are gone; the prefix is applied in exactly one place. Only code that
  scrapes stderr for those exact strings is affected — the conditions, and now
  their codes, are unchanged.

- **`@shbernal/ts-pptx/html`: HTML `<table>` → slides, anywhere there is a DOM.**
  `tableToSlides` was reachable only as a method on the browser build, which made
  converting an existing HTML table a browser-only capability. It is now also a
  free function on a new `/html` subpath that runs under Node with any DOM
  implementation and in the browser, from one artifact — deliberately no
  `browser`/`node` condition split.

  ```ts
  import { tableToSlides } from '@shbernal/ts-pptx/html'

  tableToSlides(pptx, win.document.getElementById('report'))
  tableToSlides(pptx, 'report', { document: win.document }) // by id
  ```

  Pass the element and no global DOM is consulted at all — the document and its
  view come from the element's own `ownerDocument`/`defaultView`. Pass a string
  id and it resolves against the new `TableToSlidesProps.document`, defaulting to
  the global `document` as before. `pptx` is structural (`addSlide` +
  `presLayout`), so any presentation instance works, including the Node build;
  `masterTitle` resolves on both forms.

  Strictly additive: `TsPptx.prototype.tableToSlides(eleId, options)` keeps its
  exact signature and behavior, and its body is now a delegation to the shared
  implementation, so the two forms cannot drift.

  **Column widths degrade without a layout engine.** In a browser the columns are
  sized from each cell's rendered `offsetWidth`. Nothing outside a browser lays a
  table out, so `offsetWidth` is `0` there — which previously made the
  proportional calc a `0/0` divide and emitted a table with zero-width columns.
  The basis now falls back in two steps: the computed CSS `width`s when the
  stylesheet states them for every column in one unit (all `px` or all `%`), then
  an equal split. `data-pptx-width` / `data-pptx-min-width` on the `<thead>`
  cells still win outright on every path, and are the way to pin widths
  regardless of runtime.

  Cell text is read the same way: `innerText` where the DOM genuinely renders,
  and otherwise a `childNodes` walk that keeps `<br>` as a line break. jsdom does
  not implement `innerText` at all (every cell would have come out empty), and
  happy-dom implements it as `textContent` rather than rendered text (every cell
  would have come out on one line).

  New public types: `TableToSlidesElement`, `TableToSlidesDocument`,
  `TableToSlidesHost`. These are structural rather than `lib.dom`'s
  `Element`/`Document`, because a non-browser DOM's types do not satisfy those and
  demanding them would reject from TypeScript exactly the implementations the
  entry exists to accept.

- **`pptx-to-script` re-embeds a surface's picture fill.** An image-filled shape
  or table cell (`p:spPr/a:blipFill`, `a:tcPr/a:blipFill`) converted to a script
  and came back unfilled: the read model saw the blip, the write API could author
  one (`fill: { type: 'image', image: { data } }`), and only the converter's fill
  mappers were missing an `AssetResolver` to join them with. They have one now,
  so the bytes are carried as an asset exactly as an `addImage`'s are — deduped
  against every other reference to the same part — and the blip's
  `a:alphaModFix` opacity carries as `transparency`.

  The write path emits every picture fill as a plain stretched blip, so a tiled,
  cropped or inset fill is re-embedded and then *flattened*: that is the new
  `fill.picture.geometry` / `table.cell.fill.picture.geometry` note, recorded
  only when the source uses one of them. The older `fill.picture` /
  `table.cell.fill.picture` notes narrow to the surfaces whose bytes cannot be
  carried at all — a blip embedding no part, a part missing from the package, and
  an SVG, which `addImage` accepts but a fill does not.

- **The read model can now see a picture fill.** `TableCell.pictureFill` and
  `AutoShape.pictureFill` decode an `a:blipFill` on a *surface* — the cell's
  `a:tcPr`, the shape's `p:spPr` — into `{ relId, partName, mode, srcRect,
  fillRect, tile, alpha, dpi, rotWithShape }`. Rect edges are per-edge fractions
  (`÷ 100000`, so `0.1` is 10 %, and a negative `fillRect` edge stays negative);
  tile offsets stay in EMU and tile scales become fractions. All three call sites
  share one reader (`readPictureFill`), so a slide background's `image` variant
  now carries the same decoded `picture` alongside the `relId`/`partName` it
  always had.

  This closes the asymmetry the cell picture-fill writer opened: the library could
  author an image-filled cell and then read it back as unfilled, because
  `resolvedFill` decodes solid colours only. It also settles what a `Picture` is
  and is not — a `p:pic` whose image is its sibling `p:blipFill` — as against a
  shape or cell whose *surface* happens to be an image.

  Gated on `test/read/fixtures/table-cell-image-fill.pptx` for the cell half
  (stretched, tiled, bordered and merged picture cells, plus a solid control) and
  on `math-omml.pptx` for the shape half, whose `p:spPr/a:blipFill` carries a
  negative `<a:fillRect b="-6667"/>` bleed. That shape sits in an `mc:Fallback`
  branch, which the read model does not walk, so the test unwraps the
  `mc:AlternateContent` — the shape XML and its relationships are PowerPoint's own
  bytes either way.

### Fixed

- **`tableToSlides` cell padding was parsed with a regex that deleted the decimal
  point.** Computed `padding-*` went through `.replace(/\D/gi, '')`, which strips
  every non-digit — including the `.` — so a `1.5px` padding became the number
  `15`, a ten-fold inset, and `0.5px` became `5`. Fractional computed paddings are
  ordinary: any `em`/`%` padding, or a `rem` on a non-integer root size, resolves
  to one. It is parsed as a CSS length now, keeping the fraction and rounding
  once. A value that is not an absolute px length (a `%` padding, a keyword) insets
  by nothing rather than by whatever digits it contained.

  This is visible to existing browser callers: a table whose cells have
  fractionally-computed padding gets a correct inset where it previously got a
  roughly 10× one.

- **`tableToSlides` produced an empty table for an id that is not a valid CSS
  identifier.** The table id was interpolated raw into a CSS selector
  (`#${id} tr:first-child th`), so an id starting with a digit, or containing `.`
  or `:`, matched nothing — *after* passing the `getElementById` reality-check,
  which made it look like the table had simply been read as empty. Every query is
  now scoped to the element, so the id is never parsed as a selector. Selector
  semantics are otherwise unchanged.

- **`tableToSlides` emitted the literal color `NANNANNAN` for a computed color
  that was not `rgb()`.** Computed colors were parsed by stripping `rgb(`/`rgba(`
  and splitting on commas. A browser always normalizes to `rgb()`, so this held
  there — but nothing outside a browser normalizes, and a DOM that returns the
  authored `#ff0000` produced `Number('#ff0000')` → `NaN` → that string, emitted
  without complaint. Colors now go through a CSS parser that handles
  `rgb()`/`rgba()` and `#rgb`/`#rrggbb`, clamps and rounds channels, and falls
  back to the caller's default rather than guessing for anything else. Browser
  output is unchanged: `rgb()` parses exactly as before, and a fully transparent
  background still becomes white.

  A dead condition next to it was removed:
  `getComputedStyle(cell).getPropertyValue('transparent')` tested a CSS *keyword*
  as if it were a property, so it returned `''` for every cell and never fired.

- **`tableToSlides` dropped every row that was not inside a `<thead>`/`<tbody>`/
  `<tfoot>`, then threw `addTable: Array expected!`.** Rows were collected with
  three descendant queries (`thead tr`, `tbody tr`, `tfoot tr`), which see only
  sectioned rows. `<table><tr>…` is valid authored markup, and a table assembled
  with `createElement`/`appendChild` has no `<tbody>` at all — the HTML *parser*
  inserts one, the DOM API does not. Such a table lost all of its rows and
  reached `addTable` empty, which is the reported failure in upstream
  gitbrent/PptxGenJS#1005. Rows now come from the table's own row list, and a row
  in no section is treated as a body row. The same change stops a table nested
  inside a cell from having its rows folded into the outer table — the descendant
  queries matched those too (as does happy-dom's non-conformant `rows`, so each
  row's ownership is checked rather than assumed).

- **`tableToSlides` applied `data-pptx-width` to the wrong column when a header
  cell spanned.** The width overrides were looked up by *column* index
  (`thead tr:first-child th:nth-child(n)`) against a row indexed by *cell* — the
  two part ways the moment a `colspan` is involved. A 2-span header with
  `data-pptx-width="4"` followed by a header with `data-pptx-width="2"` sized the
  span's second column with the *next* header's 2in and left the last column with
  no override at all. This is upstream gitbrent/PptxGenJS#1244. A spanning cell's
  `data-pptx-width` / `data-pptx-min-width` now divides across the columns it
  covers, matching what its `offsetWidth` and its computed CSS width already did.

  Two related mismatches went with it. The overrides were read from `<thead>`
  `<th>` cells while the widths themselves came from the first row *anywhere*, so
  a table with no `<thead>` measured one row and took overrides from a row that
  did not exist — both now come from the same cells. And the width-source query
  was a descendant selector, so a table with `<th>` in both `<thead>` and
  `<tfoot>` derived twice as many columns as it had; `addTable` then rejected the
  column count and discarded `colW` wholesale, taking every override with it.

- **`tableToSlides` emitted ragged tables, and crashed on an empty one.** An HTML
  row states only the cells it starts, so a short row is ordinary markup; pptx has
  no such model — `<a:tblGrid>` declares a column count and a row carrying fewer
  `<a:tc>` is a table PowerPoint has to repair. Rows are now measured against the
  grid the table actually occupies (`colspan` widening a row, a `rowspan` from
  above filling one it never mentions) and padded with blank cells to the width of
  the widest row. A table with no cells at all no longer fails deep in the
  auto-pager with `Reduce of empty array with no initial value`; it throws a
  `tableToSlides:` error naming what is missing.

- **A line's `cap`, and a stroke's `pattern`/`image` paint, were dropped before
  reaching the emitter.** `ShapeLineProps extends ShapeFillProps`, so a stroke
  accepts `gradient`/`pattern`/`image` as well as a solid `color`, plus its own
  `cap` — and `drawingml/line.ts` reads all of them. But the define pass rebuilt
  the caller's `line` object from a fixed list of keys, so any key added to the
  type without also being added to that list never survived normalization.

  `cap` was silently ignored everywhere: `line: { cap: 'round' }` on a shape and
  `border: { cap: 'round' }` on a table both emitted `cap="flat"`. `pattern` was
  worse than ignored — `line: { type: 'pattern', pattern: {…} }` reached
  `genXmlPatternFill` with no pattern object and threw *"Pattern fill requires a
  pattern object."* A gradient stroke was dropped on the `addText({ shape:
  'line' })` path specifically, which carried its own near-duplicate rebuild.

  All four rebuilds (two for shape/text lines, two for table borders) now spread
  the caller's object and override only the keys they actually default, so they
  cannot fall out of sync with the type again. Output for any deck that did not
  set the dropped options is byte-identical. `addBackground()` / `defineSlideMaster({ background })` derived
  the media extension from `path` only. With no path it substituted the
  `preencoded.png` placeholder, so `background: { data:
  'data:image/svg+xml;base64,…' }` embedded SVG bytes in a `.png` part that
  `[Content_Types].xml` announced as `image/png` — the Default/payload mismatch
  PowerPoint offers to "repair" — and the same held for any non-PNG format. The
  `data:` mime is now read first and wins over `path`, as it always has for
  `addImage()`.

  The sniff itself moves to one shared `imageExtensionForSource(path, data)` in
  `src/media/content-type.ts`, replacing the four near-copies that had drifted
  apart (image objects, image fills, picture bullets, OLE preview covers) —
  which is how the background copy came to be the one missing the mime branch.
  Two small consequences of the single implementation: a mime is now lower-cased
  the way a path extension already was (`data:image/PNG;` names its part `.png`,
  not `.PNG`), and an OLE object's `cover` follows the same bytes-win precedence
  as everything else.

- **A combo chart's per-subchart options are validated like the chart-level
  ones.** `addChart` normalizes its options once, but a `ChartMulti` entry's own
  `options` were merged over them only at emit time — after every clamp and enum
  correction had run — so they reached the part verbatim. A subchart
  `barOverlapPct: 250` emitted `<c:overlap val="250"/>` where `ST_Overlap` is
  -100..100, `barGapWidthPct: 9999` blew past `ST_GapAmount`'s 500 and
  `barGrouping: 'sideways'` failed the `ST_Grouping` enumeration: three
  PowerPoint-repair prompts reachable only through the combo API, and silent —
  the same values passed at chart level have always been clamped with a warning.

  Each subchart's options now go through the same pass, keyed to that subchart's
  own plot type, covering `barDir`, `barGrouping`, `barGapWidthPct`,
  `barGapDepthPct`, `barOverlapPct`, `bar3DShape`, `holeSize`, `firstSliceAng`,
  `lineDataSymbol`, `lineDataSymbolSize`, `lineDataSymbolLineSize` and
  `dataLabelPosition`. What it validates is the merged value the emitter actually
  reads, writing back only what a correction changed, so per-subchart options stay
  a sparse override of the chart-level ones.

  That also closes the wider half of the same hole: a combo chart's internal
  `_type` is a `ChartMulti[]`, so the chart-level corrections that key off the
  chart *type* — `barGrouping`, `dataLabelPosition` — previously matched no branch
  and never ran either. They now resolve per subchart, which is why one bad
  chart-level `barGrouping` correctly lands as `clustered` for a bar group and
  `standard` for a line group.

  One behaviour change comes with it: a **stacked bar subchart now emits
  `gapWidth 50`**, the narrower default a chart-level stacked bar already got,
  where it previously inherited the clustered default of 150. An explicit
  `barGapWidthPct` on either the chart or the subchart still wins.

- **A table cell with a non-solid fill no longer reports the table style's
  colour.** `TableCell.resolvedFill` fell through to the style graph whenever the
  cell's own fill was not a solid one, so an image-, gradient- or pattern-filled
  cell under a shading style reported a colour PowerPoint never paints, and an
  explicit `a:noFill` cell reported the shading it was suppressing. A cell that
  declares any `EG_FillProperties` choice of its own now overrides the style, the
  same guard `AutoShape.resolvedFill` already applied to the theme style matrix.
  A cell with no fill choice at all still inherits its banding/header shading as
  before.

### Added

- **A table cell can now be filled with a picture.** `addTable` accepts an image
  fill on a cell exactly as a shape or text box already did — `fill: { type:
  'image', image: { path } }` (or `{ data }`, or a bare `image:` with no `type`)
  — and emits `<a:blipFill>` inside `<a:tcPr>`, after the cell's borders. It
  works through every route a fill reaches a cell: per-cell `options.fill`, the
  `headerRow` and `columns[i]` sugar, and the table-level `fill`. Merged regions
  fill uniformly across the whole span, and an auto-paged table registers each
  overflow slide's media on that slide rather than piling every relationship onto
  the first one.

  This is a cell *fill*, not a picture nested in a cell — `CT_TableCell` accepts
  only `a:txBody`, `a:tcPr` and `a:extLst`, so no `<p:pic>` can live inside an
  `<a:tc>` at any effort. To float a real picture over a cell instead, use
  `pptx.tableLayout()` to get that cell's computed rect and `addImage()` at those
  coordinates — which is also how PowerPoint itself fakes it.

  No new API surface: `TableCellProps.fill` was already `ShapeFillProps`, so an
  image fill always type-checked. What was missing was the registration step —
  nothing walked a table's cells to allocate the media relationship, which is why
  it degraded at write time (see Fixed, below). Registration is keyed on fill
  *object identity*, so the `headerRow`/`columns` sugar — which shares one fill
  object across every cell it styles — mints a single relationship rather than one
  per cell.

  Gated on `test/read/fixtures/table-cell-image-fill.pptx`, authored by desktop
  PowerPoint: stretched, tiled, bordered and merged picture-fill cells in one
  table. Two findings from it are worth keeping. PowerPoint writes a bare
  `<a:tcPr/>` on a merged region's *covered* cells and repeats no fill there,
  while this library copies the origin's fill onto them — kept deliberately, since
  a covered cell is never rendered and the copy keeps image fills uniform with the
  solid case. And PowerPoint's *stretched* cell fill omits `dpi="0"
  rotWithShape="1"` and `<a:srcRect/>`, which our shared `genXmlImageFill` always
  writes; that shared emitter was left alone, because both attributes are optional
  with no schema default and PowerPoint authors exactly that attribute set for its
  own *tiled* cell in the same table.

- **A deck IR (`@shbernal/ts-pptx/script`), the read half of turning an existing
  `.pptx` back into source.** `readModelToIr(presentation)` walks a deck read
  through `ts-pptx/read` and returns a serializable description of the write-API
  calls that would rebuild it — `{ slideSize, props, slides, assets, fidelity }`,
  where each slide holds `{ method, args }` calls whose `args` are literal
  write-API option objects. Geometry is carried as exact `"<n>emu"` strings
  wherever the option is `Coord`-typed, and as six-decimal inches for the three
  that are not (`colW`, `rowH`, `margin`) — the proven minimum for an EMU-exact
  round-trip.

  It is a new subsystem rather than part of `ts-pptx/read` because it needs both
  the read model and the write option types, and because the read subpath is
  documented as isomorphic (bytes in, bytes out); a converter whose output is
  source text would break that guarantee for its consumers.

  Every construct that cannot survive is a `FidelityNote` on the IR rather than a
  log line, carrying `{ slideNumber, shapeName, construct, disposition, cause,
  detail }`. `cause` distinguishes a missing read accessor (`unread`) from a
  missing write option (`unwritable`) from a structural limit (`unsupported`),
  which is what makes a note actionable. Because the notes are data, a round-trip
  check can exclude exactly the declared losses and treat any other difference as
  a defect — an undeclared loss fails, and a declared loss that actually survives
  is a stale note. Read `DeckIr.fidelity` before trusting a conversion: notable
  entries include theme-referenced outline width (`p:style/a:lnRef` resolves a
  colour but not a width or dash), embedded audio/video (only the poster frame is
  readable), and OMML equations.

- **`printScript(ir, options)` turns that IR into a runnable TypeScript module.**
  Returns `{ code, assets, notes }`: the module source, the image bytes it
  expects beside it, and the losses that apply to *this* output. The emitted
  script anchors on a template, and the template is the **source deck itself,
  unmodified** — `Presentation.fromTemplate` already strips a package's slides
  while leaving its masters, layouts, theme, and document properties
  byte-identical, so only slide content is regenerated and the deck's whole
  design survives untouched. Slides are emitted in source order; contiguous
  slides sharing a layout share one generator, since `appendSlides` binds one
  layout per call.

  `notes` is not simply `ir.fidelity`. A template-anchored output *rescues* some
  declared losses — all twelve document properties ride in the template, so the
  IR's `deck.docProps` note does not apply and is omitted rather than left to
  teach readers to skim. It also *adds* losses that belong to the tier rather
  than to the conversion: a slide's `p:cSld@name` reads fine and would survive a
  byte copy, but has no public write-API setter. The applicable set is
  reproduced as a header comment in `code`, so the artifact carries its own
  caveats, and it is the set a round-trip check should exclude from its diff.

  Binding is by layout name where that is unambiguous, because a name survives
  being re-pointed at a different template; a deck whose masters repeat a layout
  name falls back to gallery position, since `appendSlides` throws on an
  ambiguous name rather than choosing.

- **`printStandaloneScript(ir, options)` prints the same IR with no template at
  all.** The emitted module depends on nothing but this package: the theme, one
  `defineSlideMaster` per source layout, and every slide, all re-authored through
  the public write API and therefore all editable. It is the second printer over
  one IR, and the only thing the two differ in is where the deck's chrome comes
  from.

  The trade is worth stating before choosing a tier. Template-anchored output
  gets the deck's entire design back byte for byte, at the cost of shipping the
  source deck alongside the script and of leaving that design uneditable.
  Standalone output is one file, and the parts of the original design the read
  model cannot see are gone. Three of them are unreachable from *both*
  directions, so no amount of converter work recovers them: `a:fmtScheme` (the
  fill, line and effect style lists a shape's `p:style` indexes into — no reader,
  and the write path emits a hardcoded Office one), `p:txStyles` (the master's
  per-level text styles — no reader, though `SlideMasterProps.textStyles` could
  author them), and master/layout decoration. A fourth, `p:clrMap`, is readable
  with no setter. Each is a fidelity note in this tier and rides across untouched
  in the other, which is why the other shipped first.

  A `defineSlideMaster` here carries a title and a background and nothing else,
  and that is a write-path constraint rather than a shortcut:
  `addPlaceholdersToSlideLayouts` seeds every slide with each layout placeholder
  the slide did not populate, as an empty text shape. Since this converter
  authors every source shape as concrete absolute-positioned content and binds
  none of them to a placeholder, re-declaring a layout's placeholders would add a
  ghost shape to every slide for each one.

  The second consumer moved the IR twice, both still within this unreleased
  window. `DeckIr` gained `chrome` (`{ theme, masters[] }`), which the
  template-anchored printer ignores entirely since the source deck *is* its
  chrome. And `SlideIr.calls` is now populated for a `carried` slide too: marking
  a slide carried was an erasure and is now a recommendation, because a printer
  with no source package to copy from had the whole slide erased rather than only
  the unwritable construct that made it uncarryable — and that construct already
  declares its own loss.

  Two notes describe the write path rather than the source deck, because both are
  permanent properties of any standalone output: a presentation always carries a
  blank layout named `DEFAULT`, seeded in the constructor with no way to remove
  it (`master.default`), and all five settable document properties are stamped in
  the constructor and cannot be unset — assigning `''` writes an empty element
  rather than removing it (`deck.docPropsDefault`).

- **A guide for the whole subsystem: [PPTX To Script](docs/reference/pptx-to-script.md).**
  The two tiers and how to choose, the chrome cliff that forces the split, the
  fidelity-note contract, the measured loss list across the fixture corpus, and
  what a clean round-trip run does and does not prove.

- **A round-trip check for generated scripts: `canonicalDeckIr` + `diffDeckIr`,
  and `pnpm run script:roundtrip`.** Reads a deck, prints a script, runs it,
  reads the deck that came out, and diffs the two IRs using the printer's
  fidelity notes as the exclusion list — so a difference no note predicted is a
  defect. `diffDeckIr` returns `{ differences, undeclared, declared, added,
  unmatchedNotes }`; `undeclared` is the number to gate on.

  The comparison is on IRs rather than packages because the output can never be
  byte-identical (fresh rel ids, regenerated shape ids), so comparing bytes would
  fail for every deck and measure nothing. `canonicalDeckIr` removes the noise
  that is not loss — a value spelled out that means what its absence means
  (`bold: false`, `wrap: true`, the default `a:bodyPr` insets), and asset
  identity by content digest rather than by generated filename — and each such
  rule cites the OOXML default that makes it an equivalence rather than a
  convenience.

  Its reach is bounded, and knowing how bounded is the point. Both IRs come from
  the same reader, so a construct the read path cannot see is absent from both
  and compares equal; and the converter need not be injective, so two source
  constructs mapping to one call also compare equal. It detects **asymmetry**.
  Mutation testing says so concretely: of twelve deliberate converter defects it
  catches six, and the six that survive are exactly the symmetric ones. Read a
  clean run as "nothing the converter can distinguish was lost", and pair it with
  `pnpm run read:census` and the IR unit tests, whose expectations come from
  `src/types/*.ts` rather than from the converter.

- **`ts-pptx/script` now transcribes a slide's show transition, in both tiers.**
  A `SlideIr` gained a `transition` field and the printers emit it as
  `slideN.transition = { … }` — a property assignment rather than a call, which
  is how the write API models it. Speed bucket, exact `p14:dur` duration,
  `advClick`/`advTm` advance behaviour and the type-specific variant attributes
  (`{ dir: 'd' }`, `{ spokes: '2' }`) all carry across, each omitted from the
  emitted literal when the source left it at its OOXML default.

  The type is filtered against the write API's closed vocabulary. The read model
  reports `TransitionInfo.type` as an *open string*, because it also decodes
  PowerPoint's modern effects (Morph, Vortex, Ripple, …) and tells them apart by
  namespace, while `TransitionType` names the 21 base ECMA-376 transitions and
  nothing else. A name that does not survive the filter files a
  `slide.transition` note instead of producing a script that does not compile.
  PowerPoint's own probed effect table lists exactly those 21 base effects, so the
  filter is checked against ground truth rather than against a transcribed list.

  Transition **sounds** map in both OOXML forms: the stop-previous `p:endSnd`, and
  an embedded start sound whose WAV is resolved through the slide's own `r:embed`
  and carried as an asset. The embedded form survives the standalone tier only —
  `extractSlides` does not surface a transition's audio part, so the append path
  the template-anchored tier rides has nothing to wire and drops the sound
  (silently, with no dangling reference). That tier declares it as
  `slide.transitionSound`.

  The round-trip oracle was widened to match: `CanonicalSlide` now carries
  `transition`, compared structurally rather than as one opaque value, so a note
  can declare a lost *sound* without also excusing a wrong transition *type*.
  Without that, a printer that stopped emitting transitions entirely would have
  produced identical calls and reported a clean round trip.

  Assets are now numbered per media kind (`image1.png`, `audio1.wav`) instead of
  over one shared counter, so a generated script does not bind a transition sound
  to a `const image7`.

### Fixed

- **`addChart` mutated the arrays and options object it was handed.** Series
  normalization was applied in place, so `pptx.addChart(data, opts)` rewrote
  `data[0].labels` from `['A','B','C']` to `[['A','B','C']]` — the nested form
  the multi-level category serializer wants — and stamped an internal
  `_dataIndex` onto every series. Any caller that reused its own data afterwards
  (to build a legend, a table, or a second chart) silently got one nested array
  where it had passed three strings, and the failure surfaced far from the chart
  call. `addChart` now normalizes into copies; the caller's series objects are
  never written back to.

  The options object had the same problem and is fixed with it: defaults
  (`chartColors`, `barGapWidthPct`, `plotArea`, …) were written onto the caller's
  object, and invalid entries were *deleted* from it — an out-of-range
  `layout.x`, a bad `catGridLine.size`, a `dataLabelPosition` illegal for the
  chart type. Sharing one options object across two charts therefore meant the
  second chart saw the first chart's normalization. Options are now copied before
  anything is applied, so `addChart` treats both of its arguments as read-only
  inputs.

  Two consequences worth noting. `_dataIndex` is gone from the public
  `OptsChartData` type — it was only ever there because the normalization wrote
  it onto the caller's object, and it remains on the internal series shape the
  emitters read. And code that *relied* on reading the normalized values back off
  its own options object after the call (the filled-in defaults, the clamped
  percentages) no longer sees them; pass the values explicitly instead. Emitted
  OOXML is byte-identical.

- **A theme color on a chart gridline or series line emitted invalid XML.**
  `valGridLine`, `catGridLine`, `serGridLine` and `barSeriesLine` built their
  color by hand as `<a:srgbClr val="…"/>`, bypassing the shared color emitter
  every other color in the library goes through. A scheme token therefore landed
  verbatim in the `val` attribute — `<a:srgbClr val="accent1"/>`, where the
  attribute is `ST_HexColorRGB` — so gridlines could not follow the deck's theme.
  Both emitters now route through `createColorElement`, which picks
  `<a:schemeClr val="accent1"/>` for a scheme token and `<a:srgbClr>` for hex.
  `OptsChartGridLine.color` widens from `HexColor` to `Color` accordingly, so
  `valGridLine: { color: SchemeColor.accent1 }` now type-checks and renders.

  Sharing that emitter also brings the gridline path in line with every other
  color site: a leading `#` is stripped, an 8-digit RGBA value splits its alpha
  byte into a sibling `<a:alpha>`, and an unparseable color warns and falls back
  to the default instead of being written out as-is. One cosmetic consequence:
  hex is now normalized to uppercase, so a gridline authored as `'d9d9d9'` emits
  `val="D9D9D9"`. Rendering is identical and the built-in defaults are unchanged,
  but a test that pins chart XML bytes for a lowercase gridline color will see a
  one-time diff.

- **A negative `w`/`h` produced a deck PowerPoint refused to open at all.** The
  signed value went straight into `<a:ext cx=… cy=…>`, and both attributes are
  `ST_PositiveCoordinate` — so one negative extent anywhere cost the *whole*
  presentation: *"The file or directory is corrupted and unreadable"* (0x80070570),
  naming no shape, no part, and no slide. LibreOffice rendered the same package
  happily, so a pipeline that previews with LibreOffice saw nothing wrong until
  the deck reached PowerPoint.

  It was easy to hit, because a signed delta is the natural way to write "draw
  from A to B": `addShape('line', { x: x0, y: y0, w: x1 - x0, h: y1 - y0 })`
  works fine until the line happens to run leftward or upward. A negative
  extent is now normalized to the box PowerPoint itself would write for that
  geometry — origin at the min corner, absolute extent, and a flip on the
  mirrored axis — which is the encoding `addConnector` has always derived from
  its endpoints. `{ x: 1, y: 3, w: 1.5, h: -2 }` emits
  `<a:xfrm flipV="1"><a:off x="914400" y="914400"/><a:ext cx="1371600" cy="1828800"/></a:xfrm>`.

  It applies to every object kind, not just `addShape`, because it happens at the
  one point where all of them share a placement path — after each `Coord` form
  has resolved to EMU, so `'-25%'` and `'-2in'` normalize alongside a plain
  negative number. A derived flip XOR-composes onto an explicit one, so
  `{ w: -2, flipH: true }` is mirrored twice and therefore not at all. Group
  auto-bounds normalize each child before taking the bounding box; previously a
  child with a negative extent reported a `maxX` left of the group's `minX` and
  collapsed the group frame.

  No warning is emitted, and `Math.min`/`Math.abs` at the call site is no longer
  needed — the signed form is now a supported spelling rather than a trap.

- **`addAnimation()` and `groupObjects()` could not find a shape whose
  `objectName` contained `&`, `<`, `>`, `"`, `'`, a tab or a newline.** A shape
  added as `objectName: 'Q&A'` is stored attribute-escaped (`Q&amp;A`) so it can
  be written into `<p:cNvPr name>` as-is, and the two lookups compared the
  caller's raw string against that escaped text. Neither ever matched:
  `addAnimation({ preset: 'fadeIn', objectName: 'Q&A' })` warned *"no object
  named "Q&A" on the slide"* and dropped the effect — leaving the deck with no
  `<p:timing>` at all — and `groupObjects(['Q&A'])` threw *"no top-level object
  on this slide has that objectName"*, both naming a shape that was plainly
  there. Renaming to `QandA` was the only way through. Any `objectName` now
  works as a lookup key, including for a shape inside a group and for the
  "already inside a group" hint that tells a grouped name apart from a typo.
  Escaping is now done once, where the comparison happens
  (`resolveObjectNameToId`), rather than re-derived at each call site — which is
  how the animation lookup came to disagree with the connector one, the only
  caller that had it right.

- **An image fill on a table cell type-checked, then silently rendered as no
  fill.** `TableCellProps.fill` has always been `ShapeFillProps`, so `fill: {
  type: 'image', image: { path } }` on a cell compiled and flowed all the way to
  the emitter — but nothing on the table path ever called
  `registerImageFillMedia`, so no media relationship was allocated. The fill
  reached `genXmlColorSelection`'s `case 'image'` with no resolved `_imgRid`,
  warned *"image fill is missing its resolved media reference"*, and emitted
  `<a:noFill/>`. The package contained no `ppt/media/` entry and the cell rendered
  blank. Shapes (`gen/define/shape.ts`) and text boxes (`gen/define/text.ts`) had
  been wired since image fills were introduced; tables were the one major object
  kind left out. Cell image fills now resolve properly — see Added, above.

- **A transition sound supplied as `data:audio/x-wav;…` was written to a media
  part named `.x-wav`.** The media filename was taken from the data URI's mime
  *subtype* verbatim, so the package ended up with `ppt/media/audio-1-1.x-wav`
  and a `<Default Extension="x-wav"/>` declaring a file type that exists nowhere
  else. The content type was right and PowerPoint opened it, but nothing else
  would recognise the file. The subtype now maps to a real extension
  (`audioExtensionForSubtype`), so the same bytes land on `audio-1-1.wav` under
  `<Default Extension="wav" ContentType="audio/x-wav"/>` — what PowerPoint itself
  authors. `audio/x-wav` is not an exotic input: it is exactly the content type
  PowerPoint writes for an embedded transition sound, so it arrives on every deck
  read back in and handed to `ts-pptx/script`.

- **A tab, carriage return or line feed inside an XML attribute value was emitted
  literally, so it read back as a space** (`dn-xml-attr-whitespace`). XML 1.0
  §3.3.3 requires a parser to normalise those three characters to a single space
  inside an attribute value *before any consumer sees them*; carrying one across
  needs a character reference. Every caller-supplied string that lands in an
  attribute was affected — `objectName` (`p:cNvPr/@name`), alt text
  (`p:cNvPr/@descr`), layout and slide titles (`p:cSld/@name`), section titles
  (`p14:section/@name`) and hyperlink tooltips — so a two-line layout title came
  back as one line. This is not theoretical: PowerPoint's built-in German layout
  set ships a layout named across two lines ("Abschnitts-\<LF\>überschrift").

  The fix is a new `encodeXmlAttrValue` used by the attribute-emitting paths only
  (the element builder in `src/gen/oxml/el.ts`, plus the few emitters that write
  attributes with template strings). It is deliberately *not* a widening of
  `encodeXmlEntities`, which also escapes element text — there a literal newline
  is meaningful content, and escaping it would change bytes across every
  text-bearing part in the package.

  **This changes emitted bytes** for any deck whose attribute values contain a
  tab, CR or LF; every other deck is byte-identical (verified against the full
  1637-part demo deck). Consumers that string-matched the emitted XML for such an
  attribute must now match `&#9;`/`&#10;`/`&#13;`. Reading is unaffected: any
  conforming parser resolves the references back to the original characters.

- **Nine converter defects that no static check and no execution check could
  catch**, all found by the round-trip comparison above and all producing a
  script that typechecked, ran, and wrote a plausible deck.
  - **Every paragraph bullet was wrong.** `Paragraph.bullet` reports a *tagged*
    string (`'none'`, `'char:<glyph>'`, `'autoNum:<type>'`) and the mapper read
    it as a bare glyph — so an explicit `a:buNone` became a literal `n` bullet
    (`'none'.codePointAt(0)`), a real character bullet became `c`, and a numbered
    list became `a`. The numbering scheme was also passed as `style` rather than
    `numberType`, so every numbered list fell back to the default scheme, and the
    glyph was emitted with fewer than the four hex digits the write path requires,
    which made it substitute its own default.
  - **Placeholders were emitted with no geometry at all.** A shape with no
    transform of its own reported no frame, and omitting `x`/`y`/`w`/`h` does not
    leave the geometry to be inherited — an appended slide inherits nothing — it
    produces a zero-height box in the corner. Now resolved through the
    layout/master chain via `resolvedFrame`.
  - **A group's rotation and flips were applied twice**, because the children were
    already emitted in composed slide-absolute coordinates *and* the transform was
    repeated on the group. A 30° group came back at 60°; a flipped group came back
    unflipped, the double flip having cancelled.
  - **Image crops were sent to `sizing`**, which crops in displayed inches against
    the image's measured natural size, instead of to `crop`, which is `a:srcRect`
    emitted verbatim. Fractions read as inches shrank every cropped picture.
  - **Every text body was re-anchored to centre.** An unset `a:bodyPr/@anchor`
    means top in OOXML and centre in `addText`, so the converter now spells the
    anchor out rather than leaving it off.
  - **Every uncoloured run was repainted black**, since `addText` fills a run with
    no colour using `DEF_FONT_COLOR`. The inherited colour is now resolved and
    emitted (with a note that it is frozen against later theme edits), or, where
    nothing resolves it, declared as a colour that may be wrong.
  - **PowerPoint text boxes were demoted to auto shapes.** Text-box-ness is
    `p:cNvSpPr/@txBox`, not "has no preset geometry" — PowerPoint gives every text
    box an explicit `prstGeom rect`, so the old test misclassified all of them.
  - **An SVG picture was reduced to its raster fallback.** The vector part is now
    preferred, since `addImage` accepts SVG bytes and regenerates a fallback.
  - **A bulleted paragraph with more than one run was split in two**, because the
    write path treats a bullet on a run that does not open a line as a request for
    a new paragraph. Paragraph-level `bullet` now rides on the first run only.
  - Also declared rather than fixed, each being a real limit rather than a
    mistake: an automatic field (`a:fld` — slide number, date, footer) has no
    accessor and no `addText` expression; an explicit zero `a:spcBef`/`a:spcAft`
    is indistinguishable from unset on the write side; an auto-height table row
    comes back pinned to its content's height; and a paragraph that inherits its
    bullet cannot say "inherit" through the write API.

- **`line: { type: 'none' }` did nothing.** The value is documented on
  `ShapeFillProps.type` and accepted by the type checker, but `genXmlLineFill` had
  no branch for it and emitted no paint child — which is how a stroke says
  "inherit", not "none". A shape authored with an explicitly suppressed outline
  therefore *grew* the theme's border instead of losing it. It now emits
  `<a:noFill/>`.

- **`readModelToIr` mapped three constructs onto write-API shapes that do not
  exist.** All three produced an IR that typechecked and a script that failed at
  run time, because `IrValue` is deliberately loose enough that `tsc` cannot
  check an argument against the signature it is meant to satisfy.
  - `addChart` received the chart type as a third positional argument; the
    signature is `addChart(data, options & { type })`, so the type now rides in
    the options object.
  - `addConnector` received a bounding box (`x`/`y`/`w`/`h`) and a nested `line`
    object. It takes two endpoints and flat stroke options, so connectors now
    emit `x1`/`y1`/`x2`/`y2` derived from the box **and its `a:flipH`/`a:flipV`
    flags** — without the flips every up- or leftward connector is silently
    mirrored — plus `color`/`width`/`dashType`/arrowheads. A gradient or
    translucent connector stroke has no flat spelling and is now noted.
  - A slide bound to a layout whose name is shared by another layout in the deck
    made `appendSlides` throw. `SlideIr.layout` now carries the gallery index and
    whether the name is unique.

- **Three shapes could vanish from a conversion with no fidelity note**, against
  the contract that a dropped shape is never silent: an auto shape with neither
  text nor geometry of its own (an unfilled placeholder, whose outline comes from
  the layout), and a group whose every child was dropped. Both are now declared
  (`shape.empty`, `group.empty`). The third path — an unrecognised shape kind —
  turned out to be unreachable, which the type checker proves.

- **`Presentation.appendSlides` now carries speaker notes.** A generator slide
  authored with `addNotes` previously lost its notes entirely when spliced onto a
  loaded deck — `extractSlides` emitted no notes part, so the append path had
  nothing to wire. It now emits a `notesSlide` per notes-bearing slide, wired back
  to the slide it annotates, with the notes body's own hyperlink relationships
  preserved (`rId1` = notes master, `rId2` = slide, hyperlinks from `rId3`).

  A notes slide must bind to a notes master, and a template usually has none — a
  deck authored without speaker notes carries no `notesMaster` part at all — so the
  generator's own notes master (and the theme its `.rels` requires) rides along in
  `ExtractedSlides.notesMaster` and is installed **only** when the destination deck
  has none. A destination that already has a notes master keeps it, so its notes
  styling wins; this matches the existing `importNotes` policy. `ExtractedSlide`
  gains an optional `notes` field for the same reason.

- **`markDirty()` on every read-model class that exposes `element_`.** `element_`
  hands out the live DOM node, but an edit through it was silently discarded on
  `save()` unless the caller reached the owning part themselves
  (`shape.slide.part.markDirty()` — three hops, undocumented). The obligation now
  sits on the same object as the hatch: `Slide`, `Shape`, `TextFrame`,
  `Paragraph`, `Run`, `Table`, `TableRow`, `TableCell`, `Placeholder`,
  `NotesPlaceholder`, `Theme`, `Chart`, `ChartAxis`, `ChartSeries`, `ChartEx`,
  `ChartExAxis`, `ChartExSeries`, and `ResolvedTableStyle`. `Shape.markDirty()`
  was `protected` and is now public.
- **`Slide.element_`** — the `p:sld` root, filling the one missing rung in the
  `element_` ladder. Slide-level DOM access was previously only reachable as
  `slide.part.dom`.
- Guide formulas passed to a `custGeom` shape's `guides` option now have their
  leading operation checked against the 17 operations ECMA-376 §20.1.9.11
  defines. An unknown operation (e.g. `{ name: 'w2', formula: 'bogus 1 2' }`)
  previously emitted schema-shaped but semantically dead geometry whose first
  feedback was a PowerPoint repair prompt; it now warns and skips the guide,
  matching the existing degenerate-entry behaviour. Operands are still passed
  through uninterpreted.
- The project's escape-hatch policy is now written down in
  [project target](docs/project-target.md) — the convenience-vs-guarantee rule
  and why the read path gets a deep raw hatch while the write path does not.
- **Four aggregate checks**, replacing the practice of hand-composing four or
  five scripts per iteration. `verify` (~45s) is the three typechecks +
  `backlog:validate` + the whole test suite; `verify:full` (~65s) adds
  `package:lint` and `test:package`. `check:static` and
  `check:package` are the two halves CI runs as separate jobs. `verify` and
  `verify:full` omit `lint`/`format:check` by design — the git hooks own those.
- **A `dist/` freshness guard** (`scripts/ensure-dist.mjs`) that every test,
  typecheck, and package script now starts with. It rebuilds only when `src/` or
  a build config is newer than `dist/`, and is a ~0.1s no-op otherwise. This
  replaces both halves of the old pattern — the unconditional
  `pnpm run build &&` prefix and the `:fast` twins that skipped it — so there is
  no longer a stale-`dist/` footgun to reason about.

### Changed

- **`demos/` are showcases now, and nothing there is a test.** Two flagship decks
  live in `demos/showcases/` and build with `pnpm demos:build [slug]` — a
  corporate quarterly review (themed colour scheme, five masters, native
  gradients, grouped KPI cards, three chart types, a styled table, speaker notes)
  and an image-led photo essay (full-bleed photography, gradient scrims, a
  duotone picture effect, an embedded video, a live hyperlink). They replace
  `demos/modules/`, 7,100 lines of feature-enumerating builders that made a deck
  nobody would show anyone.

  The demo smoke test (`test:demos`, `test:demo:node`, `test:demo:vite`,
  `scripts/demo-smoke.mjs`) is gone with it, and `check:package` and
  `verify:full` no longer chain it. Its actual job — proving the built package
  works for a consumer — belongs to `test:package`, which now imports all nine
  export subpaths out of an installed tarball under both npm and pnpm and forces
  the `browser` condition. One signal did not survive and is recorded as an
  accepted gap in [testing](docs/testing.md#demos-are-not-tests): the Vite build
  was the only check that put a real bundler in front of the package.

  The review deck imports nothing from `node:`, so `demos/vite-demo` imports that
  same module and builds the identical deck in a browser rather than keeping a
  second copy of demo code.

- **Breaking (internal constructors):** `ChartAxis`, `ChartSeries`, and
  `ChartExAxis` now take the owning chart part (respectively the owning `ChartEx`)
  as a second constructor argument, so `markDirty()` can reach it. These are
  obtained from `Chart.axes` / `Chart.series` / `ChartEx.axes`; only code that
  hand-constructed them is affected.
- **Breaking (structural type):** `ResolvedTableStyle` gained a required
  `markDirty(): void` member. Only code that builds the object literal itself
  (rather than reading it from `table.resolvedStyle`) is affected.
- **Breaking (internal type):** the internal `ShadowPropsInternal.opacity`
  field (the derived shadow alpha) is renamed to `_alpha`, clearing it of the
  removed public `opacity` shadow input. A stray `opacity` from an untyped/legacy
  caller is now inert (it lands on a field nothing reads) instead of being
  actively stripped — no behaviour change for supported inputs; use
  `transparency` (0–100). Only code reading the internal shape off a corrected
  shadow is affected.
- `pnpm run byte-identity:baseline` now refuses to run when `src/gen/` has
  uncommitted changes (override: `--allow-dirty`). A baseline frozen after the
  refactor has begun records the very bytes it exists to detect, so every later
  `check` passes trivially. The error names the workaround it is closing —
  `git stash` on a dirty tree, which risks unrelated work to a pop conflict.
- **The byte-identity gate builds the showcase decks.** It had built its corpus
  by importing `demos/modules/demos.mjs`, removed in the demos-to-showcases move
  above, so both subcommands died on `MODULE_NOT_FOUND` — worse than no gate,
  because a harness nobody can run still gets cited as one. It now builds every
  deck registered in `demos/showcases/lib/showcases.mjs` (a registry `build.mjs`
  reads too, so a new showcase is gated the day it lands), writes them under
  `.tmp/byte-identity/decks/` rather than over the artifacts `pnpm demos:build`
  leaves for a human, and explodes each under its own slug so a diff names the
  deck that moved. `Math.random` is reseeded per deck: on a single stream,
  editing the first deck shifts every GUID in the second, and the gate reports a
  diff in a deck nobody touched.

  The corpus is 177 parts against the old deck's 1637. Every part *kind* survives
  — charts with their embedded workbooks, media, notes slides, masters, layouts,
  themes — but it is a narrower slice of the emitters, so AGENTS.md now says to
  confirm the part you touched is in the baseline before reading a PASS as proof
  of anything.

- The three `tsc` projects are now `incremental`, keeping their build state in
  the gitignored `.tmp/` (a distinct `tsBuildInfoFile` per project). Warm
  `typecheck` drops from ~3.4s to ~1.3s; cold runs are no slower, so CI is
  unaffected.
- The OOXML schema fixtures now run concurrently, taking that suite from ~50s to
  ~10s — cheap enough that `verify` runs it on every iteration instead of
  reserving it for `verify:full`.
- The root build configs (`eslint.config.mjs`, `vitest.config.ts`,
  `tsdown.config.ts`, `tsdown.dev.config.ts`) are now linted and typechecked.
  They previously matched no ESLint `files` block and no tsconfig `include`, so
  nothing checked them at all.
- A missing OOXML validator now **fails** the read suite under `CI` instead of
  silently skipping a few hundred schema assertions. Locally it still skips, but
  prints a notice — a green local run no longer reads as a complete one.
- CI runs the static checks once rather than once per Node version, and adds a
  Windows leg for the package and demo scripts, which are the only exercise the
  Windows-specific subprocess handling in `scripts/script-utils.mjs` gets. The
  publish workflow now reuses the CI workflow instead of keeping its own copy of
  the gate, which had already drifted out of sync.

### Fixed

- **A theme-indexed picture background reported the wrong part.** When a
  `p:bgRef`'s `fmtScheme` entry is an `a:blipFill` — the third `bgFillStyleLst`
  slot in several stock Office themes (Ion, Facet, …) — the fill element comes out
  of the *theme* part, so its `r:embed` is scoped to the theme's relationships.
  `SlideBackground.resolvedFill` resolved it against the owning
  slide/layout/master's relationships instead, which does not fail loudly: the
  same id usually exists there and points at something else entirely. On an
  Ion-themed deck `resolvedFill.partName` read `/ppt/slideLayouts/slideLayout1.xml`
  where the image is `/ppt/media/image1.jpeg`. Affects `Slide.background`,
  `SlideMaster.background`, and `SlideLayout.background`.

[3.1.0]: https://github.com/shbernal/ts-pptx/releases/tag/v3.1.0
[3.0.0]: https://github.com/shbernal/ts-pptx/releases/tag/v3.0.0
[2.0.0]: https://github.com/shbernal/ts-pptx/releases/tag/v2.0.0
[1.0.0]: https://github.com/shbernal/ts-pptx/releases/tag/v1.0.0
