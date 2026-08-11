---
doc-schema-version: 1
title: "Demos"
summary: "The quarterly-review showcase deck, built in your browser and previewed in the page."
read_when:
  - Seeing what a deck built with this library looks like
  - Deciding whether to clone the repo and run the showcases yourself
doc_type: "guide"
# The slides are the page. Dropping the right-hand table of contents gives them the width
# back, and there are only two headings on it to lose.
aside: false
---

# Demos

The deck below is **built in this tab**. No server renders it, nothing is uploaded, and no
picture of a slide is stored anywhere: the page runs the same showcase module that
`pnpm demos:build quarterly-review` runs, gets a `.pptx` back as bytes, and hands those
bytes to [`pptx-html`](https://www.npmjs.com/package/pptx-html), which reads the package
and paints its slides as SVG.

That round trip is the point. A preview drawn from a screenshot would prove nothing about
the package; this one can only appear if the bytes are a deck a reader can open.

<DeckPreview />

## What you are looking at

- **The deck** is `demos/showcases/quarterly-review/` in this repository: eleven slides,
  five slide masters, three charts with real embedded workbooks, a styled table, grouped
  KPI cards and speaker notes. Kestrel Analytics is fictional.
- **The renderer** is a separate library. `pptx-html` reads a package into a slide model and
  renders that model; it does not approximate. Where it cannot model something it says so,
  and those declarations are listed under the preview.
- **The download button** builds the same deck again and saves it, through the browser
  runtime's own file-writing path. Open the result in PowerPoint: that, not the picture
  above, is the output this library is judged on.

## Running the showcases yourself

Two decks ship with the repository. The second one, *Field Notes*, is a photo essay with
embedded video and a 3D model in it; it loads its media from disk by path, which is why it
is not previewed here.

```bash
git clone https://github.com/shbernal/ts-pptx
cd ts-pptx && pnpm install
pnpm demos:build                    # both decks
pnpm demos:build quarterly-review   # just this one
```

Decks land in `demos/showcases/output/`. See
[the demos README](https://github.com/shbernal/ts-pptx/blob/master/demos/README.md) for
what else is in there.
