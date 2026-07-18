# Contributing

Thanks for your interest in PptxGenJS. This is a short pointer file — the real
contributor guidance lives in the docs below, which are kept authoritative.

## Start here

- **[AGENTS.md](AGENTS.md)** — repository expectations, scope (Node-first; two
  out-of-active-scope domains), the API evolution policy, OOXML/PowerPoint
  working rules, and the backlog workflow. Read this first, whether you are a
  human or an agent.
- **[docs/agent-development.md](docs/agent-development.md)** — how changes are
  developed and verified in this repo, end to end.
- **[docs/testing.md](docs/testing.md)** — regression, schema, package, demo,
  and coverage commands, plus the **fast edit → test inner loop** (a `tsdown`
  watcher + a Vitest watcher in two terminals) and single-test invocation.
- **[docs/project-target.md](docs/project-target.md)** — what this fork does and
  does not aim to support.

## The short version

- Use `pnpm`. The package requires Node.js `>=24`.
- Keep source changes in `src/` and tests in `test/`; treat `dist/` as generated
  output.
- Before pushing, the standard gate is:

  ```bash
  pnpm run build
  pnpm run typecheck
  pnpm run test:unit
  ```

  For emitted-OOXML changes, also run `pnpm run test:schema` and add or update a
  fixture. See [docs/testing.md](docs/testing.md) for the full matrix.
- Any change to emitted OOXML must be grounded in fixtures / schema validation /
  PowerPoint-compatibility evidence, per AGENTS.md.
- A `lefthook` pre-commit hook runs ESLint + Prettier on staged files, and
  pre-push runs lint, format-check, and typecheck. Do not bypass hooks.

## Reporting bugs and proposing changes

- Open a GitHub issue for bugs, with a minimal repro (ideally a small script that
  produces the offending `.pptx`).
- For security issues, **do not** open a public issue — see
  [SECURITY.md](SECURITY.md).
- Breaking changes are acceptable when they make the API clearer or safer; record
  them (with migration guidance) in [CHANGELOG.md](CHANGELOG.md). Not-yet-built
  proposals go in the backlog ledger (`docs/backlog.yml`) — see AGENTS.md.
