---
doc-schema-version: 1
title: "Math Equations (LaTeX / MathML → OMML)"
summary: "Author native PowerPoint equations from LaTeX or MathML via the @shbernal/ts-pptx/math subpath."
read_when:
  - Authoring PowerPoint equations from LaTeX or MathML
  - Changing the latexToOmml / mathmlToOmml converters (src/math.ts)
  - Deciding how the math: option on addText should be fed
doc_type: "guide"
---

# Math Equations (LaTeX / MathML → OMML)

The `math:` option on a text item emits a **native, editable PowerPoint equation**
(OMML inside PowerPoint's `<a14:m>` markup-compatibility envelope). That option takes
raw OMML. The `@shbernal/ts-pptx/math` subpath lets you author the equation in LaTeX
or MathML instead and get the OMML to hand it.

```
LaTeX  --temml-->  MathML  --mathml2omml-->  OMML  -->  { math: … } on addText
```

## Install the converters

The converters are **optional peer dependencies** — the core package does not pull
them in, so consumers who never author math carry no extra weight. Install them to use
this subpath:

```sh
npm install temml mathml2omml
```

- `temml` — LaTeX → MathML ([MIT](https://github.com/ronkok/Temml)).
- `mathml2omml` — MathML → OMML ([LGPL-3.0-or-later](https://github.com/fiduswriter/mathml2omml)).

`mathml2omml` is LGPL. It is never bundled into this package's output (it stays a
separate, replaceable dependency in your `node_modules`), and because it is opt-in,
consumers with policies against LGPL can simply not install it.

> **Node-only.** This subpath loads the converters synchronously via Node's
> `createRequire`, so it runs under Node, not in a browser bundle. It is an authoring
> helper; the OMML it produces is plain data you can persist and feed to `math:` from
> anywhere.

## Usage

```js
import TsPptx from '@shbernal/ts-pptx'
import { latexToOmml } from '@shbernal/ts-pptx/math'

const pptx = new TsPptx()
const slide = pptx.addSlide()

slide.addText([{ math: latexToOmml('x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}') }], {
	x: 1,
	y: 2,
	w: 8,
	h: 1,
})

await pptx.writeFile({ fileName: 'quadratic.pptx' })
```

The item's `math` value fully controls the paragraph; any `text` on the same item is
ignored (see [`TextProps.math`](./reference/api/index/interfaces/TextProps.md)).

## API

### `latexToOmml(latex, opts?)`

Convert a LaTeX math expression to OMML.

- `latex: string` — e.g. `\sum_{i=1}^{n} i = \frac{n(n+1)}{2}`.
- `opts.display?: boolean` (default `true`) — display (block) math: render in
  `displayMode` and wrap the result in a centered `<m:oMathPara>` display paragraph.
  With `display: false`, temml renders in inline mode and a bare `<m:oMath>` is returned.
- **Returns** OMML: `<m:oMathPara>…</m:oMathPara>` (display) or `<m:oMath>…</m:oMath>`
  (inline). Both are accepted by the `math:` option; pass the `display: false` form
  together with `inline: true` on the text item to flow the equation mid-paragraph
  (see [Inline math](#inline-math)).
- **Throws** on invalid LaTeX, surfacing temml's parse position, e.g.
  `Invalid LaTeX (position 6): …`.

### `mathmlToOmml(mathml)`

Convert a MathML string (`<math>…</math>`) to OMML.

- **Returns** a bare `<m:oMath>…</m:oMath>` with no namespace declarations (the `math:`
  envelope supplies the `m` prefix at emit time).

## Output form

Both functions emit OMML in the `m:` (`http://schemas.openxmlformats.org/officeDocument/2006/math`)
namespace with **no namespace declarations of their own** — the `<a14:m>` envelope that
`math:` authors declares `m`. `mathmlToOmml` always returns a bare `<m:oMath>`;
`latexToOmml` returns that same `<m:oMath>` wrapped in a centered `<m:oMathPara>` unless
`display: false`. All three shapes (inner OMML, `<m:oMath>`, `<m:oMathPara>`) are valid
inputs to `math:`.

## Inline math

By default a `math:` item is emitted as its own centered display-math paragraph. To
flow an equation *in a sentence*, between plain text runs, set `inline: true` on the
text item and give it the bare `<m:oMath>` form (`latexToOmml(tex, { display: false })`
or `mathmlToOmml(mathml)`):

```js
import { latexToOmml } from '@shbernal/ts-pptx/math'

slide.addText(
  [
    { text: 'where ' },
    { math: latexToOmml('x^2+1=y', { display: false }), inline: true },
    { text: ' holds' },
  ],
  { x: 1, y: 1, w: 8, h: 1 }
)
```

The equation is emitted as an `<a14:m><m:oMath>` run (no `<m:oMathPara>`) within the
same paragraph as the surrounding runs. The `Requires="a14"` envelope stays at the
shape level, exactly as for display math.

## Scope and limits

- **No LaTeX preprocessing / macro packages** — the input goes straight to temml. Custom
  macros, `\usepackage`, and environments temml does not support are out of scope.
- **No raster fallback** — output relies on the `Requires="a14"` envelope, understood by
  PowerPoint 2010+. There is no `mc:Fallback` image for non-a14 consumers.
- **Fidelity is temml + mathml2omml's** — a few constructs map loosely. Accent
  commands (`\hat`, `\bar`, `\vec`, …) render via `<m:limUpp>` rather than `<m:acc>`
  because temml omits `accent="true"` on the `<mover>` it emits (tracked in backlog
  `fork-temml-accent-fidelity`, on hold pending a temml fix). The result is valid OMML
  and opens cleanly in PowerPoint.

## Error policy

Invalid LaTeX **throws** (with temml's parse position) rather than emitting a degenerate
equation — consistent with the library's no-silent-coercion rule. Wrap calls in
`try/catch` if you convert untrusted input.
