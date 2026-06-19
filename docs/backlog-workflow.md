---
doc-schema-version: 1
title: "Backlog Workflow"
summary: "How to record and classify this fork's own project work in docs/backlog.yml: downstream downstream needs plus the retained upstream-derived entries, without reintroducing dropped package targets."
read_when:
  - Recording a downstream need raised by downstream
  - Updating backlog classifications
  - Closing a backlog item after a fix lands
  - Deciding whether a behavior belongs in this fork
doc_type: "guide"
---

# Backlog Workflow

This workflow maintains the fork's backlog ledger, `docs/backlog.yml`. It is a
**project backlog**: a record of work this fork intends to do and decisions it
has made. It holds two kinds of entry:

1. **Downstream needs** — generic PPTX behavior the downstream consumer needs
   that belongs in this package (`type: downstream-need`, `source:
   downstream[:path]`). These are the primary, ongoing source of new work.
2. **Retained upstream-derived signals** — a set of gitbrent/PptxGenJS issues and
   PRs that were judged relevant to this fork before upstream tracking was
   retired. They carry a github reference (`owner/repo#N`) in `source` and remain
   as historical context and standing feature ideas.

> **Upstream tracking is retired.** This fork no longer fetches upstream issues
> or reconciles the ledger against GitHub, and the `backlog:check:upstream`
> tooling has been removed. Do not re-add a sync step. New entries should be
> project needs (`downstream-need`); the github-sourced entries already in the
> file are kept as-is unless a local change closes one.

The `source` field still discriminates the two kinds, and the validator enforces
it: github references for the retained legacy entries, `downstream[:path]` for
downstream needs.

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
pnpm run backlog -- show dn-some-slug
pnpm run backlog -- values status
pnpm run backlog -- validate
```

The default list output is intentionally compact: item id, status, priority,
current-project applicability, and summary. Use `--json` when another tool or
agent needs structured output, and `--print-limit 0` when a filtered listing
should print every matching entry.

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
- `target-candidate`: likely relevant to this fork; worth scoping.
- `accepted`: worth implementing or opening a local task for.
- `interesting-with-tweaks`: useful signal, but the original fix or framing does
  not fit this project as-is.
- `non-target`: dismissed because it conflicts with current goals.
- `watch`: incomplete signal; revisit only when new evidence appears.
- `deferred`: relevant, but intentionally not scheduled now.
- `implemented`: fixed locally with test or fixture evidence.
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
- `not-reproducible-in-current-project`
- `insufficient-evidence`

## Decision Questions

For each item, answer these in the ledger note:

1. Is this about generated `.pptx` correctness, current package behavior, or real
   feature coverage?
2. Does it still apply under the TS-first, ESM-only package shape?
3. Is any proposed fix tied to legacy architecture this fork dropped?
4. Can this checkout reproduce the behavior?
5. Would a local fix live in `src/` with focused tests in `test/`?
6. Does the item require OOXML schema lookup, Microsoft implementation docs, the
   validator, or a PowerPoint-authored comparison?

## Evidence Requirements

Do not mark an item `accepted` without at least one current-project evidence
path:

- a minimal PptxGenJS reproduction;
- generated `.pptx` output;
- extracted package XML path and observed problem;
- `pnpm run test:schema` result or planned fixture;
- PowerPoint repair/open result when available;
- Open XML SDK or Microsoft documentation reference when PowerPoint behavior is
  not obvious from schema alone.

For emitted OOXML changes, the implementation handoff should require a focused
fixture in `test/schema.test.js` and `pnpm run test:schema` when practical.

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
Do not leave the entry at `accepted`, `target-candidate`, or `needs-repro` after
the local fix has landed.

For each fixed item:

- set `status` to `implemented`;
- update `current_project_notes` with the local commit or fix summary;
- add the source and test files to `evidence.local_files`;
- set `schema_fixture` and `validator_result` when a schema fixture was added;
- add relevant OOXML or Microsoft references used for the fix;
- set `next_action` to `none`;
- delete any downstream `stopgap` the entry referenced.

Validate the ledger before finishing:

```bash
pnpm run backlog -- validate
pnpm run backlog -- show <id>
```

## Downstream Needs (downstream)

downstream is a consumer of this package, not part of its source. When a
downstream task exposes a generic PPTX gap that belongs here — an OOXML
serialization fix, an API/typing gap, a repeated layout primitive, media/SVG
handling, post-processing that patches generated XML — record it as a
`downstream-need` instead of leaving a one-off workaround undocumented. These are
now the main source of new backlog work.

A downstream need is something we already believe is valuable, so the full design
rationale and any long-form analysis are welcome in `current_project_notes`.

Add one with the ledger CLI, then write the rationale into the file:

```bash
pnpm run backlog -- add --id sf-<slug> --type downstream-need \
  --source downstream:<path/that/needs/it> \
  --summary "<one line>" --priority p2 \
  --stopgap <downstream path the gap forces a workaround in>
```

`add` writes a valid skeleton (defaults: `status: target-candidate`, `priority:
p2`, `applies_to_current_project: yes`, today's dates) and validates the result.
Then edit the entry to add `target_area`, evidence, and the design essay under
`current_project_notes` (a `|` block scalar). The `stopgap` field records the
downstream file carrying the temporary workaround, so the loop is closeable:
when the fix lands here, flip `status` to `implemented` and delete the stopgap
downstream.

`id` uses an `sf-<slug>` prefix.

## Promotion Checklist (before moving a candidate into the fork)

1. Prove the need with a downstream deck or eval.
2. Reduce the behavior to a minimal PptxGenJS fixture.
3. Add a PptxGenJS regression or schema test.
4. Pack or link the fork into downstream.
5. Run the relevant downstream build/render/lint/eval command.
6. Keep only generic code in PptxGenJS; keep project policy in downstream.

## Keep In downstream (not fork candidates)

These encode downstream/downstream specifics or deck workflow and stay downstream —
do not raise them as backlog items:

- downstream brand guidance, CV workflow scripts, and workflow-specific content.
- Aptos as a project default font.
- Lucide and Dashboard Icons policy, imports, aliases, and provenance manifests.
- Pexels or other external asset sourcing helpers.
- `slide-lint` quality thresholds, annotated screenshots, and human-review artifacts.
- Slide semantics manifests as agent-facing design-intent contracts.
- Greenfield deck eval prompts, scorecards, and Codex adapter behavior.
- LibreOffice/ImageMagick rendering orchestration for local visual QA.
