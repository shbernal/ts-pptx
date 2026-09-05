# Contributing

Thanks for your interest in ts-pptx. The README describes what the package does for
the people who install it. This file is the other half: how the repository is built,
tested and contributed to.

## Start here

- **[AGENTS.md](AGENTS.md)**: repository expectations, scope, the API evolution
  policy, and the OOXML/PowerPoint working rules. Read this first, whether you are a
  human or an agent.
- **[docs/agent-development.md](docs/agent-development.md)**: how changes are
  developed and verified in this repo, end to end.
- **[docs/testing.md](docs/testing.md)**: regression, schema, package, demo and
  coverage commands, plus the fast edit-then-test inner loop (a `tsdown` watcher and a
  Vitest watcher in two terminals) and single-test invocation.
- **[docs/project-target.md](docs/project-target.md)**: what this project aims to
  support, and what it leaves to a consumer.

## Setting up

The repository uses `pnpm` and Node.js 24 or newer.

```bash
pnpm install
pnpm run verify        # while iterating
pnpm run verify:full   # before pushing, and for package or release changes
```

Keep source changes in `src/` and tests in `test/`; treat `dist/` as generated output.
The project site is a fourth tree: markdown content in `docs/`, the theme and components
that render it in `www/`. `pnpm run docs:dev` serves the lot, see
[docs/development.md](docs/development.md#site-changes).

Changes to emitted OOXML need a fixture in `test/schema-cases.js`, which `verify`
already runs. The schema suite validates through `ooxml-validate`, which fetches and
caches its oracle binary on first use, so there is nothing to install. Every change to
emitted OOXML must be grounded in fixtures, schema validation, or PowerPoint
compatibility evidence, per AGENTS.md.

Changes to the package boundary (exports, entry points, shipped artifacts) should also
run:

```bash
pnpm run check:package
```

A `lefthook` pre-commit hook runs oxlint and oxfmt on staged files, and pre-push runs
lint, format-check and typecheck. Do not bypass hooks.

## Demos

- `demos/showcases` builds the two flagship decks from one command.
- `demos/node` exercises Node.js ESM generation and stream output.
- The [demos page](https://shbernal.github.io/ts-pptx/demos) builds the quarterly
  review deck in a browser.

## Scope

The project is Node-first: it generates and is tested without a browser or any office
application. Two areas sit outside *active* maintenance, not because they lack merit,
but because no in-house use case drives them, so the maintainer will generally not pick
up bugs or feature requests there:

- **Live-DOM and browser-layout features**, meaning anything whose answer comes from a
  *rendered* page: real `offsetWidth` after layout, the resolved cascade, fonts as the
  browser actually chose them. Converting an HTML `<table>` is not in this category
  (see [HTML tables to slides](docs/html-tables.md)); only real measurement needs a
  browser.
- **Third-party office-suite interop quirks** that appear only after a file is
  round-tripped through another application, for example copy/paste inside WPS Office
  and then opening in PowerPoint, when the generated package is itself valid OOXML.
  The supported bar is that output opens cleanly in Microsoft PowerPoint.

**Contributions in both areas are welcome.** Issues and pull requests are encouraged
even though the maintainer is not actively developing them.
[`docs/project-target.md`](docs/project-target.md) carries the full scope statement and
suggested testing approaches.

## Reporting bugs and proposing changes

GitHub issues are the only tracker; there is no local ledger. The
[new-issue chooser](https://github.com/shbernal/ts-pptx/issues/new/choose) offers two
forms:

- **Bug or fidelity limit**: wrong output, a repair prompt, a regression, or a
  construct that does not survive a round trip. Bring a minimal repro, a small script
  that produces the offending `.pptx`.
- **API gap**: a missing accessor, or a property the write side authors that the read
  side cannot see.

Neither fits? File a blank issue. A good issue in the wrong shape beats a bad issue in
the right one. See [errors](docs/errors.md#which-failures-are-worth-reporting) for
which failures are worth a report.

Describe a downstream consumer's need **anonymously**. Issues are public, see
[AGENTS.md](AGENTS.md). For security issues, **do not** open a public issue, see
[SECURITY.md](SECURITY.md).

Breaking changes are acceptable when they make the API clearer or safer. Record them,
with migration guidance, in [CHANGELOG.md](CHANGELOG.md).

### The skill that files the issue for you

Most code that uses this library is written by an agent, and an agent that hits a
library defect will usually route around it silently. The defect is never reported and
never fixed. `ts-pptx-upstream` is a skill that turns that moment into a filed issue
with a minimal reproduction, which is what becomes a permanent regression test here. It
ships inside the package, so it is already on disk:

```bash
# Name the skill and the runtimes, and take the defaults: this is the form that
# completes unattended, which is how an agent will be running it.
npx skills add ./node_modules/pptx-ts -s '*' -a claude-code -a codex -a universal -y

npx skills add shbernal/ts-pptx   # same flags, straight from the repo instead of node_modules
```

Drop the flags for an interactive prompt if you are at a terminal yourself. Reaching
for `--all` to avoid the prompt installs into every runtime the CLI knows about, around
seventy of them, and leaves an `agent/` directory at your repository root for runtimes
nobody there uses. Name the ones you have.

Two things about living with the installed copy. It is a copy, so **a version bump does
not update it**: re-run the command above, or `npx skills update ts-pptx-upstream`, in
the same commit as the bump. And if you track it, track `skills-lock.json` and ignore
the copy: `skills experimental_install` restores the file from that lock but creates
none of the runtime links, so it is a record, not a restore command.

The skill covers triage (is this ours, your deck's, or out of scope?), reducing a
failure to a script that builds its own deck, and, because presentations carry client
names and unreleased numbers, never uploading one to a public tracker. Once the
reproduction stands on its own it files without interrupting you, and tells you the
issue number afterwards. It also covers the far end of the cycle, which is the half
that usually rots: when a release lands, finding every workaround it retires and
deleting them.

## Installing an unreleased commit

Any commit is installable directly from GitHub, without waiting for a release. This is
how you try a fix before it ships:

```bash
pnpm add github:shbernal/ts-pptx#<commit-sha>
```

`master` (`github:shbernal/ts-pptx`) works too, but pin the sha: a branch spec
re-resolves to whatever is at the head of it when the lockfile is next written.

`dist/` is not committed, so this builds the package on install: your package manager
clones the repo, installs this package's `devDependencies`, and runs its `prepare`
script. That makes the install slow (a couple of minutes) and heavier than a registry
install, and it needs a working Node toolchain. It is meant for trying a fix, not for
production dependencies.

Note that `pres.version` reports the version in `package.json` at that commit, so
several different commits report the same number. The sha in your `package.json` is
what identifies the build.
