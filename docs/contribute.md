---
doc-schema-version: 1
title: "Contributing"
summary: "Where to report a problem, and where the documentation for working on the repository lives."
read_when:
  - Reporting a bug or proposing a change
  - Looking for the development, testing, architecture or release documentation
doc_type: "overview"
---

# Contributing

Report wrong output, a repair prompt or a missing accessor through the
[new-issue chooser](https://github.com/shbernal/ts-pptx/issues/new/choose).
[Errors and warnings](errors-and-warnings.md#which-failures-are-worth-reporting) says which failures are worth a report.
Security issues go through
[SECURITY.md](https://github.com/shbernal/ts-pptx/blob/master/SECURITY.md), never a public issue.

[CONTRIBUTING.md](https://github.com/shbernal/ts-pptx/blob/master/CONTRIBUTING.md) covers setting
up the repository, the git hooks, the scope it is maintained for, reporting a bug or promoting a
downstream need, and installing an unreleased commit.

The documentation for working on the repository is read on GitHub:

- [Development guide](https://github.com/shbernal/ts-pptx/blob/master/docs/contributing/development.md):
  setup, source layout, generated outputs, and contribution rules.
- [Testing guide](https://github.com/shbernal/ts-pptx/blob/master/docs/contributing/testing.md):
  the gate matrix, git hooks and CI jobs, the test suites, coverage, schema validation, and the
  oracles behind a PowerPoint claim.
- [Architecture](https://github.com/shbernal/ts-pptx/blob/master/docs/contributing/architecture.md):
  how the library is structured and where each responsibility lives.
- [Scope and design policy](https://github.com/shbernal/ts-pptx/blob/master/docs/contributing/scope-and-policy.md):
  what the library leaves out on purpose, the areas outside active development, and the rule for
  escape hatches.
- [OOXML agent context](https://github.com/shbernal/ts-pptx/blob/master/docs/contributing/ooxml.md):
  where to look up OOXML and PowerPoint behavior before changing emitted XML, and what counts as
  evidence for a change.
- [Release workflow](https://github.com/shbernal/ts-pptx/blob/master/docs/contributing/releasing.md):
  how a release reaches npm under both package names, and how to finish a half-published one.
- [Design pages](https://github.com/shbernal/ts-pptx/tree/master/docs/contributing/design): how
  animations, embedded fonts, embedded objects, fills and measured text fit are written and read.
