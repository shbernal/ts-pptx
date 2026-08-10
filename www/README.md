# www/ — the site's application code

`docs/` holds the site's **content**: markdown pages under a frontmatter schema, checked by
`docs:check` and navigated from `docs.json`. This directory holds the **code that renders
it** — the VitePress theme, its stylesheet, and the Vue components a page can mount.

The split exists because the demos page is a real application: it builds a `.pptx` in the
tab, reads it back through [`pptx-html`](https://www.npmjs.com/package/pptx-html), and
paints the result. Putting that inside `docs/` would make `docs/` stop being the docs — a
tree whose every page is validated by the docs kit is the wrong home for a component with
a state machine in it.

```
www/
  theme/
    index.ts    the VitePress theme: extends the default, registers <DeckPreview />
    style.css   the site's own palette and the demos page's styles
  demos/
    deck-preview.ts    the pipeline and its types — plain TS, typechecked, unit-tested
    DeckPreview.vue    the markup and the wiring around it
```

VitePress only looks for a theme at `<root>/.vitepress/theme`, so `docs/.vitepress/theme/index.ts`
is a one-line re-export of `www/theme`. That shim is the entire cost of the boundary.

## Working on it

```bash
pnpm run docs:dev       # the whole site, hot-reloaded, at http://localhost:5173/ts-pptx/
pnpm run docs:build     # what CI publishes; runs docs:check on both sides of it
pnpm run typecheck:site # tsc over www/**/*.ts and docs/.vitepress/**
```

The `.vue` file's template is **not** typechecked — `tsc` does not read SFCs and this repo
does not carry `vue-tsc`. That is why every non-trivial line lives in `deck-preview.ts`
instead: the component is markup plus a handful of assignments, and the part that could be
wrong is in a file the typechecker reads.

## What the demos page depends on

Two copies of the library, on purpose:

- The **workspace** build (`dist/`, via `ts-pptx-demos-showcases`) writes the deck.
- The **published** `@shbernal/ts-pptx` that `pptx-html` depends on reads it back.

They meet as a `Uint8Array` and share no objects, so the duplication costs bytes and
nothing else. Pinning `pptx-html` to the workspace copy instead would mean the first
breaking change here breaks the docs deploy of the release that introduces it.
