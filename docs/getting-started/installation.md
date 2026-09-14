---
doc-schema-version: 1
title: "Installation"
summary: "What ts-pptx needs, how to install it, the optional math dependencies, and how to import it from ESM, CommonJS or a page with no build step."
read_when:
  - Adding ts-pptx to a project
  - Importing it from CommonJS or from a page with no build step
  - Choosing between the pptx-ts and @shbernal/ts-pptx package names
doc_type: "guide"
---

# Installation

Add ts-pptx to a project, then import it from an ES module, from CommonJS, or from a page with no
build step.

## Requirements

- Node.js 24 or later. The package declares `"engines": { "node": ">=24" }`.
- An ES module. The package is `"type": "module"` and ships one build, which CommonJS code can
  still load (see [below](#commonjs)).
- Nothing else for TypeScript: the declarations ship inside the package, so there is no
  `@types` package to add.

## Install

::: code-group

```bash [pnpm]
pnpm add pptx-ts
```

```bash [npm]
npm install pptx-ts
```

:::

`@shbernal/ts-pptx` is the same package under the name it was first published as, released from
the same commit at the same version, so projects that already depend on it keep working. Install
one name, not both: to your program, two copies in one dependency tree are two separate libraries.
These docs use `pptx-ts`.

## Optional dependencies for math

`pptx-ts/math` turns LaTeX and MathML into native equations through two optional peer
dependencies. Install them only if you use it:

```bash
npm install temml mathml2omml
```

[Math equations](../math-latex.md) covers the rest, including why `mathml2omml` is never bundled.

## Import it

From an ES module, in Node or through a bundler:

```ts
import TsPptx from "pptx-ts"

const pptx = new TsPptx()
```

### CommonJS

```js
const { default: TsPptx } = require("pptx-ts")

const pptx = new TsPptx()
```

`require()` returns the module's namespace, so the class is on `.default`.
[Where it runs](runtime.md#require-from-commonjs) explains why that works
on every supported Node version.

### A page with no build step

A `<script type="module">` can import `https://esm.sh/pptx-ts/browser` directly. The full snippet
is in [a browser with a script tag](runtime.md#a-browser-with-a-script-tag).

## Next

[Your first deck](first-deck.md) builds a small deck from data, start to finish.
