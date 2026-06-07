# Upstream Signal Workflow

This workflow tracks upstream PptxGenJS issues and pull requests as signals for
this fork. Upstream is useful for finding real `.pptx` generation gaps, but it
is not the local project target.

Use this workflow when reviewing upstream reports or PRs that may point to:

- PowerPoint repair prompts or corrupt output;
- invalid OOXML or Open XML SDK validation failures;
- broken package relationships, content types, or part layout;
- missing PPTX generation features that fit the current API;
- browser or TypeScript issues that apply to the current ESM package boundary.

Do not use this workflow to reintroduce dropped upstream targets:

- CommonJS package support;
- IIFE or global browser bundles;
- direct CDN script-tag workflows;
- old generated artifact names under `dist/`;
- upstream release matrix or build-system compatibility work.

## Source Of Truth

The decision ledger is [upstream-signals.yml](upstream-signals.yml). It records
what has been dismissed, what is under consideration, and what should be
implemented locally.

The ledger is intentionally metadata-first. Do not copy full upstream issue or
PR bodies into this repository. Store the upstream URL, a short summary, current
decision, and evidence gathered against this checkout.

## Review Inputs

Before reviewing a candidate, read the local project boundaries:

- [Project target](project-target.md)
- [Runtime and package support](runtime-and-package-support.md)
- [Agent development guide](agent-development.md)
- [OOXML agent context](ooxml-agent-context.md)
- [Testing guide](testing.md)

When a candidate affects emitted OOXML, follow the OOXML workflow before making
source changes. Use the `ooxml` MCP server for ECMA-376 structure and the
`microsoft_learn` MCP server for Microsoft or PowerPoint-specific behavior.

## Candidate Collection

Agents may fetch upstream metadata from `gitbrent/PptxGenJS`, but should avoid
reviewing full threads until a candidate passes the local relevance filter.

Useful metadata fields:

- issue or PR number;
- title;
- state;
- labels;
- created and updated dates;
- URL;
- whether the item is an issue or pull request.

Example metadata-only collection command:

```bash
gh issue list --repo gitbrent/PptxGenJS --state all \
  --json number,title,state,labels,createdAt,updatedAt,url
```

Use `gh pr list` with the same repo and similar fields for pull requests.

## Untreated Item Check

After an initial pass, use the read-only checker to find upstream issue and PR
metadata that is not represented in the ledger:

```bash
pnpm run upstream:signals:check
```

The checker requires the GitHub CLI (`gh`). It uses metadata from `gh` and
compares upstream numbers against entries in `docs/upstream-signals.yml`. It
accepts filters for routine follow-up checks:

```bash
pnpm run upstream:signals:check -- --state open --type issue
pnpm run upstream:signals:check -- --created-since 2026-06-07
pnpm run upstream:signals:check:json -- --updated-since 2026-06-07
```

The checker never edits the ledger. Record reviewed decisions manually so
`seen` does not become confused with `reviewed`.

## Ledger Tooling

Use the local ledger command to inspect and maintain entries already recorded in
`docs/upstream-signals.yml`:

```bash
pnpm run upstream:signals -- list
pnpm run upstream:signals -- list --status needs-repro --type issue
pnpm run upstream:signals -- show upstream-issue-1440
pnpm run upstream:signals -- values status
pnpm run upstream:signals -- validate
```

The default list output is intentionally compact: item id, status, priority,
current-project applicability, and `upstream_summary`. Use `--json` when another
tool or agent needs structured output, and `--print-limit 0` when a filtered
listing should print every matching entry.

Use `values status` to see which status values are currently used in the ledger
and how many entries use each one.

The command also supports exact-ID maintenance operations:

```bash
pnpm run upstream:signals -- set-status upstream-issue-1440 implemented
pnpm run upstream:signals -- remove upstream-issue-1440
```

Mutation commands validate the ledger before writing and refuse ambiguous or
duplicate ids. Use `--dry-run` to check the intended mutation without writing.

## Classification

Classify each reviewed item with one status:

- `unreviewed`: fetched, but not judged.
- `needs-repro`: plausible, but no current-project reproduction exists yet.
- `target-candidate`: likely relevant to this fork.
- `accepted`: worth implementing or opening a local issue for.
- `interesting-with-tweaks`: useful upstream signal, but the upstream fix or
  framing does not fit this project.
- `non-target`: dismissed because it conflicts with current goals.
- `watch`: incomplete upstream signal; revisit only when new evidence appears.
- `implemented`: fixed locally with test or fixture evidence.
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

For each reviewed item, answer these in the ledger note:

1. Is this about generated `.pptx` correctness, current package behavior, or real
   feature coverage?
2. Does it still apply after the TS-first, ESM-only package changes?
3. Is the upstream proposed fix tied to legacy architecture?
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
- upstream adds a reproduction that applies to this checkout;
- a local bug proves the same root cause;
- PowerPoint or Open XML SDK behavior shows that the previous dismissal was too
  narrow.

Update `last_reviewed`, `status`, and `current_project_notes` when reopening.

## Implementation Handoff

When a ledger item moves to `accepted`, create a local implementation task with:

- upstream source URL;
- local reproduction steps;
- expected generated package or XML behavior;
- relevant `src/` and `test/` files;
- OOXML and Microsoft references, if applicable;
- verification commands to run.

Keep the fix local to the current package target. If the upstream PR includes a
useful idea but also reintroduces non-target package behavior, mark the signal
`interesting-with-tweaks` and implement only the applicable behavior.

## Closing Implemented Signals

After fixing an upstream signal, update [upstream-signals.yml](upstream-signals.yml)
in the same work session. Do not leave the entry at `accepted`,
`target-candidate`, or `needs-repro` after the local fix has landed.

For each fixed upstream issue or PR:

- set `status` to `implemented`;
- update `current_project_notes` with the local commit or fix summary;
- add the source and test files to `evidence.local_files`;
- set `schema_fixture` and `validator_result` when a schema fixture was added;
- add relevant OOXML or Microsoft references used for the fix;
- set `next_action` to `none`.

If one source fix closes both a PR and its linked issue, update both ledger
entries so future reviews do not reopen already-fixed work.

Validate the ledger before finishing:

```bash
pnpm run upstream:signals -- validate
pnpm run upstream:signals -- show upstream-issue-1234
```
