# Contributing

Thanks for your interest in ts-pptx. The README describes what the package does for the people who
install it. This file covers how the repository is built, tested and contributed to.

## Start here

- **[AGENTS.md](AGENTS.md)**: repository expectations, scope, the API evolution policy, and the
  OOXML and PowerPoint working rules. Read it first, whether you are a human or an agent.
- **[docs/contributing/development.md](docs/contributing/development.md)**: setup, repository
  layout, source conventions and the everyday commands.
- **[docs/contributing/testing.md](docs/contributing/testing.md)**: the test suites and gates, and
  the [gate matrix](docs/contributing/testing.md#gate-matrix) of what runs where.
- **[docs/contributing/agent-development.md](docs/contributing/agent-development.md)**: how an
  agent-assisted change is developed and verified.
- **[docs/contributing/scope-and-policy.md](docs/contributing/scope-and-policy.md)**: what this
  project aims to support, and what it leaves to a consumer.

## Setting up

The repository uses `pnpm` and Node.js 24 or newer.

```bash
pnpm install
pnpm run verify        # while iterating
pnpm run verify:full   # before pushing, and for package or release changes
```

The [gate matrix](docs/contributing/testing.md#gate-matrix) shows what each command runs.
[Repository layout](docs/contributing/development.md#repository-layout) says where source, tests and
generated output live. The project site keeps its content in `docs/` and its theme and components
in `www/`. `pnpm run docs:dev` serves it, and
[Site changes](docs/contributing/development.md#site-changes) covers both trees.

A change to emitted OOXML needs a fixture in `test/schema-cases.js` and evidence behind it, per
AGENTS.md. The [testing guide](docs/contributing/testing.md) covers schema validation.

A change to the package boundary (exports, entry points, shipped artifacts) also needs
`pnpm run check:package`.

## Git hooks

`pnpm install` installs the hooks through the `prepare` script (`scripts/install-hooks.mjs`). When
`core.hooksPath` points outside this repository, the installer skips and says so, and these hooks
then run only if that hooks path hands control back to them. Do not bypass hooks.

| Hook | Job | What it runs |
| --- | --- | --- |
| pre-commit | `oxlint` | `oxlint --fix` on staged `*.{ts,mjs,js}` files, then re-stages the fixes |
| pre-commit | `oxfmt` | `oxfmt --write` on staged `*.{json,jsonc,yaml,yml,mjs,mts,ts,js}` files except `pnpm-lock.yaml`, then re-stages the result |
| pre-commit | `charcheck` | `charcheck --staged --max-warnings 0` over the staged content, when the commit touches a `*.{md,vue,ts,mts}` file |
| commit-msg | `no-ai-attribution` | Rejects a message that carries an agent-attribution trailer or footer |
| commit-msg | `no-shell-quoting-leak` | Rejects a message holding a leaked here-string delimiter, such as a line that is only `@'` |
| pre-push | `lint` | `pnpm run lint` |
| pre-push | `lint-chars` | `pnpm run lint:chars`, over the whole repository |
| pre-push | `format` | `pnpm run format:check` |
| pre-push | `typecheck` | `pnpm run typecheck` |
| pre-push | `typecheck-scripts` | `pnpm run typecheck:scripts` |
| pre-push | `typecheck-site` | `pnpm run typecheck:site` |

Pre-commit runs its jobs one after another, in the order above. Pre-push runs its jobs in parallel.
The two commit-msg rules come from [`shbernal/lefthook-rules`](https://github.com/shbernal/lefthook-rules)
through the `remotes:` block in `lefthook.yml`, and they skip merge and rebase commits. No hook runs
a test suite.

## Demos

- `demos/showcases` builds the two flagship decks from one command.
- `demos/node` exercises Node.js ESM generation and stream output.
- The [demos page](https://shbernal.github.io/ts-pptx/demos) builds the quarterly
  review deck in a browser.

## Scope

The project is Node-first: it generates and is tested without a browser or any office application.
Two areas sit outside active maintenance, because no in-house use case drives them:

- Live-DOM and browser-layout features, whose answer comes from a rendered page. Converting an HTML
  `<table>` is not one of them; see [HTML tables to slides](docs/html-tables.md).
- Third-party office-suite interop quirks that appear only after another application round-trips a
  file that is itself valid OOXML.

Issues and pull requests in both areas are welcome.
[`docs/contributing/scope-and-policy.md`](docs/contributing/scope-and-policy.md) carries the full
scope statement and suggested testing approaches.

## Reporting bugs and proposing changes

GitHub issues are the only tracker; there is no local ledger. The
[new-issue chooser](https://github.com/shbernal/ts-pptx/issues/new/choose) offers three
forms, from `.github/ISSUE_TEMPLATE/`:

- **Bug or fidelity limit**: wrong output, a repair prompt, a regression, or a
  construct that does not survive a round trip. Bring a minimal repro, a small script
  that produces the offending `.pptx`.
- **API gap**: a missing accessor, or a property the write side authors that the read
  side cannot see.
- **Agent-assisted report**: a defect an agent found while using ts-pptx in another project.
  The library's own error messages link to this form, and it asks which error class and code
  was thrown.

Neither fits? File a blank issue. A good issue in the wrong shape beats a bad issue in
the right one. See [errors and warnings](docs/errors-and-warnings.md#which-failures-are-worth-reporting) for
which failures are worth a report.

Describe a downstream consumer's need anonymously, as
[Promoting a downstream need](#promoting-a-downstream-need) sets out. For security issues, do not
open a public issue; see [SECURITY.md](SECURITY.md).

Breaking changes are acceptable when they make the API clearer or safer. Record them,
with migration guidance, in [CHANGELOG.md](CHANGELOG.md).

### Promoting a downstream need

Most new work here starts with a downstream consumer hitting a generic PPTX gap: an OOXML
serialization fix, an API or typing gap, a layout helper written for the third time, media and SVG
handling, or post-processing that patches generated XML after the fact. Before moving one of those
into this project:

1. Prove the need with a minimal, consumer-agnostic reproduction.
2. Reduce the behavior to a minimal ts-pptx fixture.
3. Add a ts-pptx regression or schema test.
4. Pack or link the project into the downstream consumer to verify.
5. Run the consumer's build, render, lint or eval path against the linked project.
6. Keep only generic code in ts-pptx, and keep project policy downstream.

Report such a gap as a GitHub issue, and describe it anonymously. Issues are public, and the
consumer is not. State the missing PPTX behavior and how any consumer would reproduce it. Never
include the consumer's name, file paths, deck or client names, or content.

A report is evidence about generation bugs and missing features. It is not a vote on the package
target. Repair prompts, invalid OOXML, content types, relationships, chart, table and media
serialization, and current TypeScript or ESM behavior are all candidates. A request that rests on
CommonJS, IIFE globals, a direct CDN script tag or a legacy artifact name is not, until the
documented target changes.
[What stays in the consumer](docs/contributing/scope-and-policy.md#what-stays-in-the-consumer)
lists what this package turns down however good the case is.

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
script. That makes the install slow and heavier than a registry install, and it needs a
working Node toolchain. It is meant for trying a fix, not for production dependencies.

Note that `pres.version` reports the version in `package.json` at that commit, so
several different commits report the same number. The sha in your `package.json` is
what identifies the build.
