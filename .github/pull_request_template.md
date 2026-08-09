<!--
  This file must stay at `.github/pull_request_template.md`. GitHub only
  auto-populates a PR body from `pull_request_template.md` at the repo root, in
  `docs/`, or directly in `.github/`. Inside a `.github/PULL_REQUEST_TEMPLATE/`
  subdirectory it becomes a *multi-template* choice that applies only when a
  `?template=` query parameter is passed — i.e. never, in practice. It lived
  there once and silently applied to nothing. Do not "tidy" it back.
-->

## Summary

<!-- Required: what changed, in a sentence or two. -->

## Motivation

<!-- Required: why. What problem does this solve, or what does it enable? -->

## Related issue

<!--
  Use `Closes #N` so the issue closes when this merges. If the change fixes an
  issue but you do not want it closed, say why.
-->

## Change type

- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change
- [ ] Documentation
- [ ] Internal / refactor

## Checklist

- [ ] I ran `pnpm run verify:full` and it passed.
- [ ] I performed a self-review of my own changes.
- [ ] The change is covered by a focused test.
- [ ] **Emitted-OOXML changes only:** the change is grounded in evidence and
      carries a fixture, per
      [docs/evidence-and-fixtures.md](../docs/evidence-and-fixtures.md).
- [ ] **Breaking changes only:** [CHANGELOG.md](../CHANGELOG.md) records the
      change and its migration guidance.
- [ ] I read [AGENTS.md](../AGENTS.md) and
      [docs/agent-development.md](../docs/agent-development.md), which are the
      authoritative contributor docs.

## Notes for the reviewer

<!-- Optional: screenshots, sample code, anything that is easier shown. -->
