---
doc-schema-version: 1
title: "Evidence and fixtures"
summary: "What counts as evidence before implementing a change, and the rule that a feature testable only against genuine PowerPoint output waits for a real fixture instead of being guessed at."
read_when:
  - Deciding whether there is enough evidence to start implementing
  - Implementing a read accessor or write-side behaviour with no fixture yet
  - Authoring a PowerPoint-produced fixture or oracle
doc_type: "guide"
---

# Evidence and fixtures

Two rules ground a change here. One says what counts as evidence. The other says
what to do when the only evidence that *would* count does not exist yet.

## Evidence requirements

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

## Fixture-gated work, and why guessing fails

Some work can only be judged against XML PowerPoint itself wrote. A read accessor
checked against real Office XML. A write path whose target is "whatever PowerPoint
authors" for preset IDs, part wiring, namespaces, inheritance. If that fixture does
not exist yet, do not substitute one. Hand-typed XML, synthetic XML, and write→read
round trips all agree with whatever the code already does, which is exactly the
question being asked.

So stop. The fixture is a blocking precondition, not a follow-up. Open a GitHub
issue naming the exact construct the oracle has to contain, and leave the feature
unimplemented until it lands.

Then author it with the `powerpoint-fixture-authoring` skill and verify it with that
skill's own
`.agents/skills/powerpoint-fixture-authoring/scripts/verify-powerpoint-fixture.ps1`
(there are no `.ps1` files under `scripts/`). Record provenance and SHA-256 in
[test/read/fixtures/README.md](https://github.com/shbernal/ts-pptx/blob/master/test/read/fixtures/README.md).
Wire the test to the fixture, through the read harness for read accessors or a
`test/schema-cases.js` comparison for write-side oracles. Implement last.

## Related

- [Testing guide](testing.md): the verification commands, and
  [The object model is not a render oracle](testing.md#the-object-model-is-not-a-render-oracle)
  for the one claim that needs *render* evidence specifically.
- [OOXML agent context](ooxml-agent-context.md): where to look up the structure
  a fixture is supposed to demonstrate.
- [Project target](project-target.md): whether the behaviour belongs here at all.
