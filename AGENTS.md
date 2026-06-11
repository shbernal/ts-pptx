# AGENTS.md

## Repository Expectations

- This repository builds PptxGenJS, a JavaScript/TypeScript library that emits PowerPoint `.pptx` packages using OOXML.
- Use `pnpm` for repository scripts. The package declares Node `>=24`.
- Keep source changes focused in `src/` and tests in `test/`. Treat `dist/` as generated package artifacts unless the task explicitly requires refreshing release outputs.
- Preserve unrelated dirty state. Do not revert user changes.

## API Evolution Policy

- This fork is maintained for our own use; there is no external backward-compat
  obligation. Prefer fixing root causes here rather than asking consumers (e.g.
  `downstream`) to work around them — a fix in this public package helps every
  consumer.
- Breaking changes are acceptable and encouraged when they make the API clearer
  or safer. Do not block an improvement on reverse compatibility. When you make
  or propose one, record it (with migration guidance) in the
  `UPSTREAMING_CANDIDATES.md` tracker at `../UPSTREAMING_CANDIDATES.md`.
- Silent coercion of invalid input is a footgun, not a feature: prefer warning or
  failing on `NaN` / `undefined` / out-of-range values over emitting a degenerate
  result (e.g. a zero-size object).

## OOXML And PowerPoint Work

- Before changing emitted OOXML, read `docs/ooxml-agent-context.md`.
- Do not vendor full standards PDFs or large extracted specification text into this repository as agent context. Store small, repo-specific notes with section references instead.
- Prefer executable evidence over prose alone: inspect minimal PowerPoint-authored `.pptx` packages when needed, compare package XML, and add focused regression or schema fixtures.

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

## Upstream Signals Workflow

- `docs/upstream-signals.yml` tracks upstream issues and PRs as signals for local work.
- When you implement a fix or feature derived from an upstream signal, update the corresponding item in that file: set `status` to `implemented`, update `last_reviewed` to today's date, update `current_project_notes` with where the fix landed, update `evidence.local_files` to reflect the current source location, and set `next_action` to `none`.
- Also update any companion items that share the same root cause (e.g. an issue whose `next_action` was `handle-with-upstream-pr-NNNN` when that PR is now implemented).
- Every field constrained by `vocabulary` (`status`, `priority`, `target_area`, `applies_to_current_project`, `non_target_reasons`, `evidence.kinds`) MUST use a value already listed under that file's top-level `vocabulary:` block. Before writing a value, scan the `vocabulary:` lists and reuse the closest existing term — do not invent synonyms (`validator-pass` for `validator-result`, `repro-confirmed` for `minimal-repro`, etc.), as the validator rejects them.
- If no listed value genuinely fits the situation, do not force an approximation: add the new value to the appropriate `vocabulary:` list (with a one-line rationale in your message) in the same change, then use it. Extending the controlled vocabulary deliberately is fine; drifting away from it by typo is not.
- ALWAYS run `pnpm run upstream:signals:validate` after editing `docs/upstream-signals.yml` (it is fast and offline) and fix every reported error before committing. A clean ledger is a precondition for the edit being considered done. `pnpm run test:tools` exercises the ledger tooling itself.

## Verification

- For source changes, run `pnpm run build` and `pnpm run typecheck` when practical.
- For behavior changes, run `pnpm run test:unit`.
- For OOXML serialization changes, add or update a fixture in `test/schema.test.js` and run `pnpm run test:schema`.
- `pnpm run test:schema` requires the validator installed with `./tools/ooxml-validator/install.sh`.
- For release/package boundary changes, consult `docs/testing.md` and run the relevant package or demo smoke commands.
