---
doc-schema-version: 1
title: "Evidence And Fixtures"
summary: "What counts as evidence before implementing a change, and the rule that a feature testable only against genuine PowerPoint output waits for a real fixture instead of being guessed at."
read_when:
  - Deciding whether there is enough evidence to start implementing
  - Implementing a read accessor or write-side behaviour with no fixture yet
  - Authoring a PowerPoint-produced fixture or oracle
doc_type: "guide"
---

# Evidence And Fixtures

Two rules govern how a change here gets grounded: what counts as evidence at all,
and what to do when the only evidence that *would* count does not exist yet.

## Evidence Requirements

Do not start implementing without at least one current-project evidence path:

- a minimal ts-pptx reproduction;
- generated `.pptx` output;
- extracted package XML path and observed problem;
- `pnpm run test:schema` result or a planned fixture;
- PowerPoint repair/open result when available;
- Open XML SDK or Microsoft documentation reference when PowerPoint behavior is
  not obvious from schema alone.

For emitted OOXML changes, the implementation should carry a focused fixture in
`test/schema-cases.js` and a `pnpm run test:schema` run when practical.

## Fixture-Gated Work: Ask For The Fixture, Don't Guess

When a feature can only be tested against OOXML that must be **genuine
PowerPoint output**, a read-model accessor validated against real Office XML, or
a write-side behaviour whose target XML is "what PowerPoint authors" (preset IDs,
part wiring, namespaces, inheritance), and that fixture/oracle does **not** yet
exist, do not implement against synthetic, hand-typed, or write→read
round-tripped XML. Guessing the target XML produces circular or wrong evidence.

Instead, treat the fixture as a blocking precondition and stop: open a GitHub
issue naming the exact construct the oracle must contain, and leave the feature
unimplemented until the fixture lands.

Author the fixture itself with the `powerpoint-fixture-authoring` skill, verify it
with that skill's own
`.agents/skills/powerpoint-fixture-authoring/scripts/verify-powerpoint-fixture.ps1`
(there are no `.ps1` files under `scripts/`), record provenance + SHA-256 in
[test/read/fixtures/README.md](https://github.com/shbernal/ts-pptx/blob/master/test/read/fixtures/README.md),
then wire the test to the fixture (the read harness for read accessors; a
`test/schema-cases.js` comparison/inspection check for write-side oracles). Only
then implement.

## Related

- [Testing guide](testing.md): the verification commands, and
  [The object model is not a render oracle](testing.md#the-object-model-is-not-a-render-oracle)
  for the one claim that needs *render* evidence specifically.
- [OOXML agent context](ooxml-agent-context.md): where to look up the structure
  a fixture is supposed to demonstrate.
- [Project target](project-target.md): whether the behaviour belongs here at all.
