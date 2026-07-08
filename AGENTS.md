# AGENTS.md

## Repository Expectations

- This repository builds PptxGenJS, a JavaScript/TypeScript library that emits PowerPoint `.pptx` packages using OOXML.
- Use `pnpm` for repository scripts. The package declares Node `>=24`.
- Keep source changes focused in `src/` and tests in `test/`. Treat `dist/` as generated package artifacts unless the task explicitly requires refreshing release outputs.
- Preserve unrelated dirty state. Do not revert user changes.

## Scope: Node-First (Two Out-Of-Active-Scope Domains)

- This project is **Node-first**: it runs and is tested without a browser or any
  office application. Two domains are out of *active* maintenance scope. Do not
  proactively build features or hunt for fixes in them, and do not block other
  work on them; when a task lands in one, say so and treat it as out of scope
  unless the user explicitly opts in. They are not rejected on merit — outside
  contributors are welcome to submit PRs — but the maintainer is not driving them.
  See `docs/project-target.md` ("Out Of Active Scope") for the full statement.
  - **Live-DOM / browser-layout features** — anything reading a *rendered* page
    rather than in-memory data, notably `tableToSlides()` (`offsetWidth`,
    `window.getComputedStyle`). These are browser-only and cannot be reproduced in
    the Node test suite. The in-memory `addTable(rows, opts)` path is the supported
    one. If logic here genuinely needs covering, extract the DOM-independent part
    into a pure helper and unit-test it (pattern: `resolveHtmlColWidth`).
  - **Third-party office-suite interop quirks** — breakage that only appears after
    a round-trip through another app (e.g. WPS copy/paste, then PowerPoint) when
    the generated package is itself valid OOXML. The supported bar is that output
    opens cleanly in Microsoft PowerPoint. Such an item only becomes in-scope with
    a repro pinning the defect to invalid OOXML the library itself emits.

## API Evolution Policy

- This fork is maintained for our own use; there is no external backward-compat
  obligation. Prefer fixing root causes here rather than asking a downstream
  consumer to work around them — a fix in this public package helps every
  consumer.
- Breaking changes are acceptable and encouraged when they make the API clearer
  or safer. Do not block an improvement on reverse compatibility. When you make
  one, record it (with migration guidance and downstream impact) in `CHANGELOG.md`.
  When you only *propose* one, or want to track a not-yet-implemented candidate,
  record it in the backlog ledger `docs/backlog.yml` (see the Backlog Workflow
  below).
- Silent coercion of invalid input is a footgun, not a feature: prefer warning or
  failing on `NaN` / `undefined` / out-of-range values over emitting a degenerate
  result (e.g. a zero-size object).

## OOXML And PowerPoint Work

- Before changing emitted OOXML, read `docs/ooxml-agent-context.md`.
- Do not vendor full standards PDFs or large extracted specification text into this repository as agent context. Store small, repo-specific notes with section references instead.
- Prefer executable evidence over prose alone: inspect minimal PowerPoint-authored `.pptx` packages when needed, compare package XML, and add focused regression or schema fixtures.
- If a feature can only be tested against genuine PowerPoint output (a read accessor validated against real Office XML, or a write-side behaviour whose target XML is "what PowerPoint authors") and that fixture/oracle does not exist yet, do not implement against synthetic or round-tripped XML. Record the fixture as the blocking precondition in `docs/backlog.yml` (tagging the relevant `constructs:` key) and stop until it is authored — see the "Fixture-Gated Work" section in `docs/backlog-workflow.md`.

### MCP Tool Selection

Two MCP servers cover complementary parts of the OOXML/PowerPoint space. Work
through them in order before falling back to web search.

**Step 1 — `ooxml` MCP** (source: ECMA-376 / ISO 29500 parsed XSDs and spec PDFs)

Use this for questions whose answer lives in the *standard itself*:
- Element and complexType definitions (`ooxml_element`, `ooxml_type`)
- Legal child elements in document order (`ooxml_children`)
- Attribute names, types, defaults, and required/optional (`ooxml_attributes`)
- Enum values for a type (`ooxml_enum`)
- Namespace URIs (`ooxml_namespace`)
- OPC package parts, content types, and relationships (`ooxml_package_part`, `ooxml_parts`)
- Free-text search across the spec PDFs (`ooxml_search`, `ooxml_section`)

The `ooxml` MCP does **not** cover Microsoft-proprietary details such as built-in
style GUIDs, behavior differences between Office versions, [MS-OE376] / [MS-PPTX]
deviations from the standard, or Open XML SDK usage. If the answer requires any of
those, move to Step 2 rather than falling back to web search.

**Step 2 — `microsoft_learn` MCP** (source: Microsoft Learn / Open Specifications)

If the `ooxml` MCP returns incomplete or no answer, always try this before web search.
Use it for:
- Microsoft Open Specifications ([MS-OE376], [MS-PPTX], [MS-OFFCRYPTO], …) —
  these document how Office *implements* or *deviates from* ECMA-376 and contain
  Microsoft-proprietary enumerations such as built-in table style GUIDs, preset
  shape adjustment ranges, and behavior flags not in the standard.
- PowerPoint-specific rendering behavior, repair heuristics, and version-gated features.
- Open XML SDK (`DocumentFormat.OpenXml`) API usage and samples.
- Azure / Microsoft 365 platform documentation.

Use `microsoft_docs_search` for a broad query first, then `microsoft_docs_fetch` on
a returned URL when you need the full page content.

**Step 3 — web search (`WebSearch` / `WebFetch`)**

Only after both MCPs have been tried and the information is still missing or
ambiguous. Useful for community discoveries (e.g. undocumented GUIDs found by
reverse-engineering), third-party library behaviour, and content that postdates the
MCPs' corpora.

## Backlog Workflow

- `docs/backlog.yml` is the fork's project backlog ledger. New work is recorded as `downstream-need` items — generic PPTX behavior a downstream consumer needs (`source: downstream`); a set of retained gitbrent/PptxGenJS issues/PRs remain from before upstream tracking was retired (`source: owner/repo#N`). The validator enforces that the source matches the type. The full process lives in `docs/backlog-workflow.md`. Upstream tracking is retired — do not re-add a GitHub sync step.
- **Downstream needs are ANONYMOUS.** This ledger is public; the consumer is not. Describe the generic PPTX gap and how any consumer reproduces it — never the consumer's name, file paths, deck/client names, or content.
- Record a not-yet-implemented candidate here only; if you implement a change immediately, its record is the fork's own commit history, tests, and `CHANGELOG.md` — do not also add a backlog entry.
- To add a downstream need, use `pnpm run backlog -- add --id dn-<slug> --type downstream-need --source downstream --summary "…"`, then write the generic design rationale into `current_project_notes`. For these we DO want full design detail (they are believed-valuable).
- When you implement a fix or feature derived from a backlog item, update the corresponding entry: set `status` to `implemented`, update `last_reviewed` to today's date, update `current_project_notes` with where the fix landed, update `evidence.local_files`, and set `next_action` to `none`. The temporary workaround lives downstream (tracked there by an in-code comment referencing the entry id); remove it when the fix lands.
- Also update any companion items that share the same root cause.
- Every field constrained by `vocabulary` (`status`, `priority`, `target_area`, `applies_to_current_project`, `non_target_reasons`, `evidence.kinds`) MUST use a value already listed under that file's top-level `vocabulary:` block. Before writing a value, scan the `vocabulary:` lists and reuse the closest existing term — do not invent synonyms (`validator-pass` for `validator-result`, `repro-confirmed` for `minimal-repro`, etc.), as the validator rejects them.
- If no listed value genuinely fits the situation, do not force an approximation: add the new value to the appropriate `vocabulary:` list (with a one-line rationale in your message) in the same change, then use it. Extending the controlled vocabulary deliberately is fine; drifting away from it by typo is not.
- ALWAYS run `pnpm run backlog:validate` after editing `docs/backlog.yml` (it is fast and offline) and fix every reported error before committing. A clean ledger is a precondition for the edit being considered done. `pnpm run test:tools` exercises the ledger tooling itself.

## Verification

- For source changes, run `pnpm run build` and `pnpm run typecheck` when practical.
- For behavior changes, run `pnpm run test:unit`.
- For OOXML serialization changes, add or update a fixture in `test/schema.test.js` and run `pnpm run test:schema`.
- `pnpm run test:schema` requires the validator installed with `./tools/ooxml-validator/install.sh`.
- For release/package boundary changes, consult `docs/testing.md` and run the relevant package or demo smoke commands.
