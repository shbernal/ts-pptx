---
doc-schema-version: 1
title: "Backlog Workflow"
summary: "How to record and classify this project's own work in docs/backlog.yml: anonymous downstream consumer needs plus the retained upstream-derived entries, without reintroducing dropped package targets."
read_when:
  - Recording a downstream need raised by a consumer
  - Updating backlog classifications
  - Closing a backlog item after a fix lands
  - Deciding whether a behavior belongs in this project
doc_type: "guide"
---

# Backlog Workflow

This workflow maintains the project's backlog ledger, `docs/backlog.yml`. It is a
**project backlog**: a record of work this project intends to do and decisions it
has made. It holds two kinds of entry:

1. **Downstream needs** — generic PPTX behavior a downstream consumer needs that
   belongs in this package (`type: downstream-need`, `source: downstream`). These
   are the primary, ongoing source of new work, and they are recorded
   **anonymously**: describe the missing PPTX behavior and its generic
   reproduction, never a private consumer's name, file paths, or content.
2. **Retained upstream-derived signals** — a set of gitbrent/PptxGenJS issues and
   PRs that were judged relevant to this project before upstream tracking was
   retired. They carry a github reference (`owner/repo#N`) in `source` and remain
   as historical context and standing feature ideas.

> **Upstream tracking is retired.** This project no longer fetches upstream issues
> or reconciles the ledger against GitHub, and the `backlog:check:upstream`
> tooling has been removed. Do not re-add a sync step. New entries should be
> project needs (`downstream-need`); the github-sourced entries already in the
> file are kept as-is unless a local change closes one.

The `source` field still discriminates the two kinds, and the validator enforces
it: github references for the retained legacy entries, the bare token `downstream`
for downstream needs (never a consumer path — that would leak private structure).

## Out-Of-Target Work

Do not use the backlog to reintroduce dropped package targets:

- CommonJS package support;
- IIFE or global browser bundles;
- direct CDN script-tag workflows;
- old generated artifact names under `dist/`;
- upstream release matrix or build-system compatibility work.

## Source Of Truth

The decision ledger is [backlog.yml](backlog.yml). It records what has been
dismissed, what is under consideration, and what should be implemented locally.

The ledger is intentionally metadata-first for the retained github entries: do
not copy full upstream issue or PR bodies into this repository. For
`downstream-need` items the full design rationale is welcome (see **Downstream
Needs** below).

For the shape of an individual entry — every field, in write order, with
realistic annotated values for both an implemented `downstream-need` and a
retained github entry — see the reference
[backlog-item-template.yml](backlog-item-template.yml). It is not validated and
not part of the ledger; it exists so the structure survives the `items:` list
being emptied. The authoritative field shape is still
`buildItemSkeleton` in [scripts/backlog-ledger.mjs](../scripts/backlog-ledger.mjs)
(what `pnpm run backlog -- add` emits), and the legal values are the
`vocabulary:` block in [backlog.yml](backlog.yml).

## Review Inputs

Before classifying or promoting a candidate, read the local project boundaries:

- [Project target](project-target.md)
- [Runtime and package support](runtime-and-package-support.md)
- [Agent development guide](agent-development.md)
- [OOXML agent context](ooxml-agent-context.md)
- [Testing guide](testing.md)

When a candidate affects emitted OOXML, follow the OOXML workflow before making
source changes. Use the `ooxml` MCP server for ECMA-376 structure and the
`microsoft_learn` MCP server for Microsoft or PowerPoint-specific behavior.

## Ledger Tooling

Use the local ledger command to inspect and maintain entries in
`docs/backlog.yml`:

```bash
pnpm run backlog -- list
pnpm run backlog -- list --status needs-repro --type downstream-need
pnpm run backlog -- show dn-some-slug other-slug
pnpm run backlog -- show --status non-target --json
pnpm run backlog -- values status
pnpm run backlog -- validate
```

The default list output is intentionally compact: item id, status, priority,
current-project applicability, and summary. It is *not* the full record — to
review rationale fields (`non_target_reasons`, `current_project_notes`,
`evidence`, …) use one of the full-detail paths instead of hand-parsing the
YAML:

- `list --json` prints the complete items (lossless), not a compact projection,
  for any other tool or agent that needs structured output.
- `show` prints full items as readable text. It takes one or more ids
  (`show a b c`) or, with no id, every item matching the list filters
  (`show --status non-target`). Add `--json` for the structured form.
- `--print-limit 0` makes a filtered `list` print every matching row.

Use `values status` to see which status values are currently used in the ledger
and how many entries use each one.

The command also supports exact-ID maintenance operations:

```bash
pnpm run backlog -- set-status dn-some-slug implemented
pnpm run backlog -- remove dn-some-slug
```

Mutation commands validate the ledger before writing and refuse ambiguous or
duplicate ids. Use `--dry-run` to check the intended mutation without writing.

## Classification

Classify each item with one status:

- `needs-repro`: plausible, but no current-project reproduction exists yet.
- `target`: likely relevant to this project; worth scoping.
- `accepted`: worth implementing or opening a local task for.
- `interesting-with-tweaks`: useful signal, but the original fix or framing does
  not fit this project as-is.
- `non-target`: dismissed because it conflicts with current goals.
- `watch`: incomplete signal; revisit only when new evidence appears.
- `deferred`: relevant, but intentionally not scheduled now.
- `implemented`: fixed locally with test or fixture evidence. Transient — the
  entry is pruned in a follow-up commit; see "Closing Implemented Items".
- `partially-implemented`: part of the work has landed; the remainder is tracked
  in the entry's `next_action`.
- `superseded`: covered by another local fix or decision.

Use these priority values:

- `p0`: corrupt or unopenable deck, PowerPoint repair prompt, data loss, invalid
  package relationship, or invalid content type.
- `p1`: Open XML validator failure, broken chart/table/media/shape
  serialization, or a high-value missing PPTX feature with a clear OOXML path.
- `p2`: feature request that maps cleanly to the current API and package target.
- `p3`: docs, examples, ergonomics, or compatibility polish.
- `none`: outside the current project target.

## Target Areas

Use one or more target areas:

- `powerpoint-repair`
- `invalid-ooxml`
- `schema-order`
- `relationship-or-content-type`
- `chart-ooxml`
- `table-ooxml`
- `image-svg-media`
- `speaker-notes-masters-layouts`
- `missing-pptx-feature`
- `typescript-api`
- `browser-modern-esm`
- `package-boundary`

Use one or more non-target reasons when dismissing an item:

- `commonjs`
- `iife-global`
- `cdn-script-tag`
- `legacy-dist-artifact`
- `unsupported-runtime`
- `upstream-build-system`
- `release-matrix`
- `docs-only-for-legacy-workflow`
- `existing-pptx-import`
- `not-reproducible-in-current-project`
- `insufficient-evidence`
- `out-of-project-scope`
- `escape-hatch-footgun`
- `composition-belongs-downstream`
- `not-supported-by-powerpoint`

The `vocabulary:` block in [backlog.yml](backlog.yml) is authoritative and carries a
comment on each of the last two explaining what distinguishes it from its neighbours.
`not-supported-by-powerpoint` in particular may only be applied on **render evidence** —
schema reasoning and a COM read-back both fail to establish it (see
[testing.md → The object model is not a render oracle](testing.md#the-object-model-is-not-a-render-oracle)).

## Decision Questions

For each item, answer these in the ledger note:

1. Is this about generated `.pptx` correctness, current package behavior, or real
   feature coverage?
2. Does it still apply under the TS-first, ESM-only package shape?
3. Is any proposed fix tied to legacy architecture this project dropped?
4. Can this checkout reproduce the behavior?
5. Would a local fix live in `src/` with focused tests in `test/`?
6. Does the item require OOXML schema lookup, Microsoft implementation docs, the
   validator, or a PowerPoint-authored comparison?

## Evidence Requirements

Do not mark an item `accepted` without at least one current-project evidence
path:

- a minimal ts-pptx reproduction;
- generated `.pptx` output;
- extracted package XML path and observed problem;
- `pnpm run test:schema` result or planned fixture;
- PowerPoint repair/open result when available;
- Open XML SDK or Microsoft documentation reference when PowerPoint behavior is
  not obvious from schema alone.

For emitted OOXML changes, the implementation handoff should require a focused
fixture in `test/schema-cases.js` and `pnpm run test:schema` when practical.

## Fixture-Gated Work: Ask For The Fixture, Don't Guess

When a feature can only be tested against OOXML that must be **genuine
PowerPoint output** — a read-model accessor validated against real Office XML, or
a write-side behaviour whose target XML is "what PowerPoint authors" (preset IDs,
part wiring, namespaces, inheritance) — and that fixture/oracle does **not** yet
exist, do not implement against synthetic, hand-typed, or write→read
round-tripped XML. Guessing the target XML produces circular or wrong evidence.

Instead, record the fixture as the blocking precondition in the backlog and stop:

- If a backlog entry already gates the feature, set its `next_action` to authoring
  the fixture (e.g. `await-reader-then-author-<construct>-fixture`) and describe
  the exact construct the oracle must contain in `current_project_notes`.
- If none exists, add a `downstream-need` entry whose `current_project_notes`
  states the "do not implement without it" fixture dependency and what the oracle
  must capture, then leave the feature unimplemented until the fixture lands.
- Tag the entry with the relevant `constructs:` key(s) (see the
  `vocabulary.constructs` list in `backlog.yml`, e.g. `custom-geom`,
  `style-ref-color`, `group-rot-flip`) so a downstream replication audit can join
  the detected construct to this gating entry.

Author the fixture itself with the `powerpoint-fixture-authoring` skill, verify it
with that skill's own
`.agents/skills/powerpoint-fixture-authoring/scripts/verify-powerpoint-fixture.ps1`
(there are no `.ps1` files under `scripts/`), record provenance + SHA-256 in
`test/read/fixtures/README.md`, then wire the test to the fixture (read harness
for read accessors; a `test/schema-cases.js` comparison/inspection check for
write-side oracles). Only then implement and close the entry.

## Reopening Dismissed Items

Reopen a `non-target` or `superseded` item only when one of these changes:

- the documented project target changes;
- a reproduction that applies to this checkout appears;
- a local bug proves the same root cause;
- PowerPoint or Open XML SDK behavior shows that the previous dismissal was too
  narrow.

Update `last_reviewed`, `status`, and `current_project_notes` when reopening.

## Implementation Handoff

When a ledger item moves to `accepted`, create a local implementation task with:

- the source reference;
- local reproduction steps;
- expected generated package or XML behavior;
- relevant `src/` and `test/` files;
- OOXML and Microsoft references, if applicable;
- verification commands to run.

Keep the fix local to the current package target.

## Closing Implemented Items

After fixing an item, update [backlog.yml](backlog.yml) in the same work session.
Do not leave the entry at `accepted`, `target`, or `needs-repro` after
the local fix has landed.

For each fixed item:

- set `status` to `implemented`;
- update `current_project_notes` with the local commit or fix summary;
- add the source and test files to `evidence.local_files`;
- set `schema_fixture` and `validator_result` when a schema fixture was added;
- add relevant OOXML or Microsoft references used for the fix;
- set `next_action` to `none`;
- remove the downstream workaround the entry gated (tracked downstream by the
  in-code comment that references this entry's id, not by any field here).

Validate the ledger before finishing:

```bash
pnpm run backlog -- validate
pnpm run backlog -- show <id>
```

### Then Prune The Entry

`implemented` is a transient state, not a resting place. The ledger records what
is still open — what has been dismissed, what is under consideration, what should
be built — and a closed entry has stopped answering that question. It starts
duplicating one the repository already answers better: the commit, its tests, and
`CHANGELOG.md`. That is the same rule AGENTS.md states for work implemented on
the spot ("its record is the project's own commit history, tests, and
CHANGELOG.md — do not also add a backlog entry"), applied at the other end of an
entry's life. Left in place, closed entries accumulate until the file reads half
as a changelog and buries the entries still doing its actual job.

Prune in a **separate, later commit** than the one that closes the entry — never
by deleting it in the fix commit itself:

```bash
pnpm run backlog -- remove <id>
```

The order is what makes the detail recoverable. The fix commit carries the entry
at `implemented` with its full closing notes, so `git log -- docs/backlog.yml`
still leads a future reader to the reasoning; deleting it in the same commit
means those notes never exist anywhere. A prune commit should name each removed
id and the commit that closed it, so the message itself is the index back into
history.

Two checks before removing:

- **Grep the repo for the id.** Other docs, source comments, or test names may
  reference it, and a removed entry must not leave a dangling reference behind.
  (The downstream workaround comment is a separate matter — it lives in the
  consumer's repository and is removed there when the fix lands.)
- **Only `implemented` prunes.** `partially-implemented` is open work whose
  `next_action` still names something unclaimed, and `superseded` points at the
  decision that replaced it; both stay. So does anything at `non-target`,
  `deferred` or `watch` — a recorded dismissal is a decision the ledger exists to
  keep, not a closed item.

  **One narrow exception: a `non-target` entry closed on
  `not-supported-by-powerpoint`.** That reason does not record a decision about this
  project's scope, which is the kind of thing the ledger is for and which can be
  revisited when the target changes. It records a property of *PowerPoint* — the
  construct is valid OOXML we can emit, and it will never paint. The finding needs to
  survive, but it needs to survive **in front of whoever is about to re-attempt the
  construct**, and nobody reads the backlog before writing an emitter. Distil it into
  the doc that feature's own workflow already sends them to (`docs/tables.md` for the
  custom-table-style case) and put any reusable method note in `docs/testing.md`, then
  prune, naming the destination doc in the prune commit. If the finding has no such
  home, it is not distilled — leave the entry in place.

`remove` validates the ledger before writing and touches nothing but the named
entry, so the resulting diff should be a pure deletion. If it is not, something
else moved and is worth looking at before committing.

## Downstream Needs

A downstream consumer of this package (not part of its source) is the main source
of new backlog work. When a consumer task exposes a generic PPTX gap that belongs
here — an OOXML serialization fix, an API/typing gap, a repeated layout primitive,
media/SVG handling, post-processing that patches generated XML — record it as a
`downstream-need` instead of leaving a one-off workaround undocumented.

**Record it anonymously.** This ledger is public; the consumer is not. Describe
the missing PPTX behavior and how *any* consumer would reproduce it. Do NOT name
the consumer project, quote its file paths, deck/client names, or content. Frame
the rationale generically ("a consumer building an assessment grid needs …"), and
let `evidence.local_files` point only at files in THIS repo.

A downstream need is something we already believe is valuable, so the full generic
design rationale and any long-form analysis are welcome in `current_project_notes`.

Add one with the ledger CLI, then write the rationale into the file:

```bash
pnpm run backlog -- add --id dn-<slug> --type downstream-need \
  --source downstream \
  --summary "<one generic line>" --priority p2
```

`add` writes a valid skeleton (defaults: `status: target`, `priority:
p2`, `applies_to_current_project: yes`, today's dates) and validates the result.
Then edit the entry to add `target_area`, evidence, and the (generic) design essay
under `current_project_notes` (a `|` block scalar).

There is no `stopgap` field: the temporary workaround lives downstream and is
tracked there by an in-code comment referencing this entry's id, keeping the
consumer's paths out of this public ledger. When the fix lands here, flip `status`
to `implemented` and remove that downstream workaround.

`id` uses a `dn-<slug>` prefix.

## Promotion Checklist (before moving a candidate into the project)

1. Prove the need with a minimal, consumer-agnostic reproduction.
2. Reduce the behavior to a minimal ts-pptx fixture.
3. Add a ts-pptx regression or schema test.
4. Pack or link the project into the downstream consumer to verify.
5. Run the consumer's build/render/lint/eval path against the linked project.
6. Keep only generic code in ts-pptx; keep project policy downstream.

## Keep Downstream (not project candidates)

These encode a specific consumer's brand, content, or deck workflow and stay
downstream — do not raise them as backlog items:

- Brand guidance, workflow-specific scripts, and consumer content.
- A consumer's default font choice.
- Icon-set policy, imports, aliases, and provenance manifests.
- External stock-asset sourcing helpers.
- Lint quality thresholds, annotated screenshots, and human-review artifacts.
- Slide semantics manifests as agent-facing design-intent contracts.
- Greenfield deck eval prompts, scorecards, and generator-adapter behavior.
- LibreOffice/ImageMagick rendering orchestration for local visual QA.
