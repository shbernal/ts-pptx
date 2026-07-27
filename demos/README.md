# Demos

Showcases for `@shbernal/ts-pptx`. **None of this is a test.** No verification aggregate
runs anything here, CI never builds a demo, and a broken demo fails no gate — the published
package is covered by `pnpm run check:package`. See
[docs/testing.md](../docs/testing.md#demos-are-not-tests) for why it works that way.

## Build the showcase decks

```bash
pnpm demos:build                    # both decks
pnpm demos:build quarterly-review   # just one
```

Decks land in `demos/showcases/output/` (git-ignored). The build takes well under a second;
it rebuilds `dist/` first only if it is stale.

## What is here

| Directory     | What it is                                                                     |
| ------------- | ------------------------------------------------------------------------------ |
| `showcases/`  | The two flagship decks. Start here. [README](showcases/README.md)               |
| `common/`     | Shared images and media the Field Notes deck draws on.                         |
| `node/`       | Streaming a generated deck from an HTTP server. [README](node/README.md)       |
| `vite-demo/`  | The browser showcase — React + Vite. [README](vite-demo/README.md)             |

The two showcase decks are deliberately unalike. One is a corporate report built from
charts, tables, and grouped shapes on a themed grid; the other is a photo essay built from
full-bleed images, gradient scrims, and picture effects. Between them they exercise most of
what the library can do, without either being a feature checklist — which is what the demos
used to be, and what made them useless as showcases and unconvincing as tests.
