# pptx-bank — uncommitted corpus of real `.pptx` files

A scratch **bank of PowerPoint decks** for ad-hoc agent and developer use:
probing OOXML structures, reproducing round-trip / read-model behaviour,
debugging, and finding candidates worth promoting to a real test fixture.

**Nothing here is committed** except this README. The directory is gitignored
(`/pptx-bank/*` with a `!README.md` negation in `.gitignore`), so you can drop
in arbitrary decks — including large, copyrighted, or client files — without any
risk of them entering Git history.

## How this differs from the other deck locations

| Location | Committed? | Purpose |
|---|---|---|
| `test/read/fixtures/` | yes — hash-pinned, provenance-tracked, MIT-licensed | curated, minimal inputs the round-trip/read harness depends on |
| `.tmp/` | no | generated **output** scratch (e.g. `pnpm run test:read:emit`) |
| `pptx-bank/` (here) | no | free-form **input** corpus for exploration and debugging |

Keep the curated fixture set small and license-clean. Use the bank for
everything else.

## Conventions

- Drop `.pptx` files in here freely. Subfolders are fine.
- This is **input** only — do not write generated output here; use `.tmp/`.
- If a deck here turns out to be a good, minimal, license-clean regression case,
  **promote** it: copy it into `test/read/fixtures/`, add it to that directory's
  provenance table + SHA-256 list, and wire it into the harness. Do not point the
  committed harness at files in this bank — they are not tracked, so CI and other
  checkouts won't have them.
- Because contents are uncommitted, treat anything here as ephemeral: it may not
  exist in a fresh clone or another machine.
