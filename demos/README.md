# Demos

Showcases for `pptx-ts`, for someone who has cloned this repository and wants a
deck out of it. **None of this is a test.** No verification aggregate runs anything here,
CI never builds a demo, and a broken demo fails no gate — the published package is covered
by `pnpm run check:package`. See
[docs/testing.md](../docs/testing.md#demos-are-not-tests) for why it works that way.

If you only want to *see* a deck, you do not need any of this: the
[demos page](https://shbernal.github.io/ts-pptx/demos) builds one in your browser and shows
you the slides.

## Build the showcase decks

```bash
pnpm demos:build                    # both decks
pnpm demos:build quarterly-review   # just one
```

Decks land in `demos/showcases/output/` (git-ignored). The build takes well under a second;
it rebuilds `dist/` first only if it is stale.

## What is here

| Directory    | What it is                                                              |
| ------------ | ----------------------------------------------------------------------- |
| `showcases/` | The two flagship decks. Start here. [README](showcases/README.md)        |
| `common/`    | Shared images and media. **Also read by the test suite** — see below.    |
| `node/`      | Streaming a generated deck from an HTTP server. [README](node/README.md) |

The two showcase decks are deliberately unalike. One is a corporate report built from
charts, tables, and grouped shapes on a themed grid; the other is a photo essay built from
full-bleed images, gradient scrims, and picture effects. Between them they exercise most of
what the library can do, without either being a feature checklist — which is what the demos
used to be, and what made them useless as showcases and unconvincing as tests.

## `common/` is not demo-only — check before deleting from it

Nothing under `demos/` is a gate, but `demos/common/images/` is not covered by that: three
of the files in it are read by suites that *are*, and by an authoring script that produces a
read fixture.

| Asset              | Read by                                                                        |
| ------------------ | ------------------------------------------------------------------------------ |
| `cc_logo.jpg`      | four `test/regression/` suites, and `test/read/fixtures/authoring/author-table-cell-image-fill.ps1` |
| `logo_square.png`  | the browser lane, both sides (`test/browser/helpers.mjs`, `test/browser/harness/harness.mjs`) |
| `lock-green.svg`   | the same two, plus `test/regression/text/text-definition.test.js`               |

The directory was pruned hard once — the unused half of it, thirty-odd images and nine
media files, is gone — on the rule "no showcase deck references it". These three survive
that rule for a reason it does not express, so apply it by grepping the whole repository
rather than just `demos/showcases/`.

## Where the browser demo went

There used to be a fourth directory here: a React + Vite + Bootstrap app that built the
quarterly review in a tab and downloaded it. It is now a page of the site
(`www/demos/`, mounted at `/demos`), which previews the slides instead of only handing you
a file, and carries no second UI framework to do it. The browser lane still drives it —
that page, not this directory, is the Playwright `demo` fixture.
