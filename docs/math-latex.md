---
doc-schema-version: 1
title: "Math equations"
summary: "Turn LaTeX or MathML into native, editable PowerPoint equations with the pptx-ts/math subpath, as a centered block or inline in a sentence."
read_when:
  - Putting an equation on a slide from LaTeX or MathML
  - Flowing an equation inside a sentence of text
  - Checking what a LaTeX construct becomes in PowerPoint's equation editor
  - Handling invalid LaTeX from user input
doc_type: "guide"
---

# Math equations

`latexToOmml` and `mathmlToOmml` from `pptx-ts/math` convert an equation to OMML, the markup
PowerPoint's equation editor stores. Pass the result to the `math` option of a text item and the
slide gets an equation you can edit in PowerPoint.

```ts
import TsPptx from 'pptx-ts'
import { latexToOmml } from 'pptx-ts/math'

const pptx = new TsPptx()
const slide = pptx.addSlide()

slide.addText([{ math: latexToOmml('x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}') }], { x: 1, y: 2, w: 8, h: 1 })

await pptx.writeFile({ fileName: 'quadratic.pptx' })
```

A `math` item ignores any `text` on the same item. Without `inline: true` it is also its own
centered paragraph.

## Install the converters

The two converters are optional peer dependencies, so a project that never writes math does not
carry them. Install both to use `pptx-ts/math`:

```sh
npm install temml mathml2omml
```

| Package | Converts | License |
| --- | --- | --- |
| [`temml`](https://github.com/ronkok/Temml) | LaTeX to MathML | MIT |
| [`mathml2omml`](https://github.com/fiduswriter/mathml2omml) | MathML to OMML | LGPL-3.0-or-later |

`mathml2omml` is never bundled into this package's output. It stays a separate, replaceable
package in your `node_modules`. A project whose policy rules out LGPL code can leave it
uninstalled and pass OMML to `math` directly.

`pptx-ts/math` runs only under Node, because it loads the converters synchronously through
Node's `createRequire`. The OMML it returns is a plain string: store it, and pass it to `math` in
any runtime.

## Functions at a glance

| Function | Option | Returns | Usable inline |
| --- | --- | --- | --- |
| `latexToOmml(latex)` | `display` omitted or `true` | A centered display block, `<m:oMathPara>` | Prefer `display: false`. With `inline: true` the block's centering wrapper is dropped. |
| `latexToOmml(latex, { display: false })` | `display: false` | A bare equation, `<m:oMath>`, rendered in inline mode | Yes |
| `mathmlToOmml(mathml)` | none | A bare equation, `<m:oMath>` | Yes |

Neither function puts namespace declarations in its result: the `math` option declares them.

The text item takes the result through two options:

| Option | Type | Default | Effect |
| --- | --- | --- | --- |
| `math` | `string` | none | OMML: a whole `<m:oMathPara>`, a whole `<m:oMath>`, or the contents of one. |
| `inline` | `boolean` | `false` | Flow the equation between the neighbouring text runs instead of giving it its own centered paragraph. No effect without `math`. |

## Flow an equation inside a sentence

Set `inline: true` on the item and give it the bare form:

```ts
slide.addText(
  [
    { text: 'where ' },
    { math: latexToOmml('x^2+1=y', { display: false }), inline: true },
    { text: ' holds' },
  ],
  { x: 1, y: 1, w: 8, h: 1 }
)
```

The equation shares one paragraph with the text on either side of it.

## Convert MathML

```ts
import { mathmlToOmml } from 'pptx-ts/math'

slide.addText(
  [{ math: mathmlToOmml('<math><msup><mi>e</mi><mi>x</mi></msup></math>') }],
  { x: 1, y: 3, w: 8, h: 1 }
)
```

`mathmlToOmml` converts your MathML as written. An `<mover>` becomes an accent only when it
states `accent="true"`, and without it the `<mover>` becomes an upper limit.

## Handle invalid LaTeX

`latexToOmml` throws on LaTeX that temml cannot parse, rather than returning a broken equation.
The message carries temml's parse position when temml reports one, as in
`Invalid LaTeX (position 6): ...`. Catch it when the input comes from users:

```ts
import { InvalidOptionError, latexToOmml } from 'pptx-ts/math'

function toOmml(latex: string): string | null {
  try {
    return latexToOmml(latex)
  } catch (err) {
    if (err instanceof InvalidOptionError && err.code === 'math/invalid-latex') return null
    throw err
  }
}
```

## What each LaTeX construct becomes

`latexToOmml` maps these constructs to these PowerPoint equation objects:

| LaTeX | Equation object |
| --- | --- |
| `\frac` | Fraction (`m:f`) |
| `\sqrt`, `\sqrt[n]` | Radical (`m:rad`) |
| `\sum`, `\int` | N-ary operator (`m:nary`) |
| `x^2`, `x_i`, `x_i^2` | Superscript, subscript, both (`m:sSup`, `m:sSub`, `m:sSubSup`) |
| `\lim_{x \to 0}` | Subscript (`m:sSub`), not a lower limit |
| `pmatrix`, `cases` | Matrix (`m:m`) |
| `\hat`, `\^`, `\tilde`, `\~`, `\acute`, `\'`, `\grave`, `` \` ``, `\ddot`, `\"`, `\dot`, `\.`, `\bar`, `\=`, `\breve`, `\u`, `\check`, `\v`, `\mathring`, `\r`, `\H`, `\vec`, `\dddot` | Accent (`m:acc`), carrying the combining mark Word writes |
| `\widehat`, `\overrightarrow`, `\overgroup`, `\overbrace`, `\underbrace` | Group character (`m:groupChr`) |
| `\overline`, `\underline` | Border box (`m:borderBox`) |
| `\stackrel`, `\xrightarrow` | Upper limit (`m:limUpp`) |
| `\utilde` and other under-accents | Lower limit (`m:limLow`). OMML has no accent object that sits below its base. |
| `\ddddot` | Upper limit (`m:limUpp`). Its four-dot mark is two characters, and an accent holds one. |
| `\left( x \right)` | Plain bracket characters, not a delimiter object |

The accent rows describe `latexToOmml`. `mathmlToOmml` follows the `accent` attribute in your
MathML instead.

## Invalid input

| Condition | Warns or throws | Code |
| --- | --- | --- |
| LaTeX temml cannot parse, such as `\frac{` or an unknown command | throws | `InvalidOptionError` `math/invalid-latex` |
| `temml` or `mathml2omml` is not installed | throws on the first call | `UnsupportedFeatureError` `math/missing-optional-peer` |
| An empty LaTeX string | returns an empty equation, no error | none |
| A string with no `<math>` element, or MathML that converts to no equation, passed to `mathmlToOmml` | throws | `InvalidOptionError` `math/invalid-mathml` |

## Limits

- The converters run only under Node. The OMML string they return works anywhere.
- LaTeX goes to temml as given: no custom macros, no `\usepackage`, and no environment temml does
  not support.
- An equation needs a reader that understands PowerPoint 2010's `a14` extension. There is no
  image fallback, and a reader without the extension skips the whole text box.
- Beyond the table above, how an equation looks is up to temml and mathml2omml.

## See also

- [`TextProps`](reference/api/index/interfaces/TextProps.md), for `math` and `inline`
- [Errors and warnings](errors-and-warnings.md)
