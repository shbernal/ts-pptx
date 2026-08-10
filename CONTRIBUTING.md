# Contributing

Thanks for your interest in ts-pptx. This is a short pointer file — the real
contributor guidance lives in the docs below, which are kept authoritative.

## Start here

- **[AGENTS.md](AGENTS.md)** — repository expectations, scope (Node-first; two
  out-of-active-scope domains), the API evolution policy, and the
  OOXML/PowerPoint working rules. Read this first, whether you are a human or an
  agent.
- **[docs/agent-development.md](docs/agent-development.md)** — how changes are
  developed and verified in this repo, end to end.
- **[docs/testing.md](docs/testing.md)** — regression, schema, package, demo,
  and coverage commands, plus the **fast edit → test inner loop** (a `tsdown`
  watcher + a Vitest watcher in two terminals) and single-test invocation.
- **[docs/project-target.md](docs/project-target.md)** — what this project does and
  does not aim to support.

## The short version

- Use `pnpm`. The package requires Node.js `>=24`.
- Keep source changes in `src/` and tests in `test/`; treat `dist/` as generated
  output. The project site is a fourth tree: markdown content in `docs/`, the theme
  and components that render it in `www/`. `pnpm run docs:dev` serves the lot —
  see [docs/development.md](docs/development.md#site-changes).
- The standard gate is one command:

  ```bash
  pnpm run verify        # while iterating
  pnpm run verify:full   # before pushing, and for package/release changes
  ```

  For emitted-OOXML changes, add or update a fixture in `test/schema-cases.js`;
  `verify` already runs the schema suite. See
  [docs/testing.md](docs/testing.md) for the full matrix.
- Any change to emitted OOXML must be grounded in fixtures / schema validation /
  PowerPoint-compatibility evidence, per AGENTS.md.
- A `lefthook` pre-commit hook runs oxlint + oxfmt on staged files, and
  pre-push runs lint, format-check, and typecheck. Do not bypass hooks.

## Reporting bugs and proposing changes

- GitHub issues are the only tracker; there is no local ledger. The
  [new-issue chooser](https://github.com/shbernal/ts-pptx/issues/new/choose)
  offers two forms:
  - **Bug or fidelity limit** — wrong output, a repair prompt, a regression, or a
    construct that does not survive a round trip. Bring a minimal repro: a small
    script that produces the offending `.pptx`.
  - **API gap** — a missing accessor, or a property the write side authors that
    the read side cannot see.

  Neither fits? File a blank issue. A good issue in the wrong shape beats a bad
  issue in the right one.
- Describe a downstream consumer's need **anonymously**. Issues are public — see
  [AGENTS.md](AGENTS.md).
- For security issues, **do not** open a public issue — see
  [SECURITY.md](SECURITY.md).
- Breaking changes are acceptable when they make the API clearer or safer; record
  them (with migration guidance) in [CHANGELOG.md](CHANGELOG.md).
