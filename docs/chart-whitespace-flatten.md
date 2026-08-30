---
doc-schema-version: 1
title: "Chart Whitespace Flatten"
summary: "Why the chart emitters stopped threading indentation through every builder, and how a whitespace-only byte change was discharged without weakening the rule that normally refuses one."
read_when:
  - Wondering why `ppt/charts/chartN.xml` is emitted flat when other parts are not
  - Reaching for `byte-identity.mjs prove-whitespace`, or considering a second use of it
  - Touching `src/gen/chart/` and expecting to find the `openPrefix`/`childPrefix`/`closePrefix` hooks
doc_type: "decision"
---

# Chart Whitespace Flatten

`src/gen/chart/` used to carry 341 of the repo's 343 `openPrefix`/`childPrefix`/`closePrefix`
arguments. They are gone. `ppt/charts/chartN.xml` is now emitted flat, and the byte change
that made it flat was proved to be whitespace by a program rather than accepted by reading a
diff.

## What the whitespace actually was

Not pretty-printing. A generated `chart1.xml` was one line, with no newlines anywhere in the part, and
space runs scattered through it:

```
<c:ser>  <c:idx val="0"/><c:order val="0"/>  <c:tx>    <c:strRef>      <c:f>Sheet1!$B$1</c:f>
```

Of 351 prefix arguments, 337 were same-line runs of 1, 2, 3, 5, 7, 9 … spaces, not aligned to
depth or to each other: `chart-axes.ts` emitted sibling elements at one space a dozen lines
after emitting their neighbours at two. Fourteen, around `<c:title>`, carried the only real
newlines in the directory. The runs were residue from the template literals the emitters were
migrated off, preserved through that migration precisely because
[byte identity](./development.md) was the thing being proved at the time.

That is the fact that decided this. Flattening did not trade a readable part for a shorter
diff, because there was no readable part. It removed leftovers.

## What it cost, and why it was worth paying

Two call sites building the same element at different depths cannot share a helper unless the
helper takes indentation as a parameter, so every shared builder in the directory grew one:
`axisTextParagraph(defRPr, lang, pPrClosePrefix, openPrefix)`, `labelFontChildren(opts, indent)`,
`chartShapeProps(fill, border, effectIndent, fmt)`. `strRefBlock` had a whole exported type,
`StrRefLayout = 'indented' | 'compact' | 'expanded'`, whose only job was to name which of three
whitespace spellings a caller wanted. `plot-scatter.ts` passed a 24-space string counted by hand.

All of it is deleted. The helpers now take the arguments they are actually about.

## How the STOP was discharged

`AGENTS.md` makes a whitespace-only byte diff a **STOP**, not a known divergence, and the
reason is not that whitespace matters. It is that *judging* a diff to be whitespace is the same
act that would wave through a content change, performed at the moment when attention is lowest.
Waiving that by hand for 57 parts and 2894 positions would have been exactly the failure it
describes.

So the judgement was replaced with a program: `node scripts/byte-identity.mjs prove-whitespace`,
built on `scripts/xml-equivalence.mjs`. It compares the frozen baseline against freshly generated
output and passes only if every difference is a whitespace-only text node in a position where
whitespace cannot be content. It is deliberately stricter than an XML canonicalisation:

- **Raw text, never decoded.** `&amp;` and `&#38;` are different. A DOM comparison cannot see an
  escaping regression, which is the class of bug `gen/oxml/el.ts` exists to centralise away.
- **Attribute order is significant.** Inert per the XML spec, and still a failure here: this
  proves a claim about whitespace, and unifying the two `<a:defRPr>` orderings in
  `chart-parts.ts` is a different change that must not ride along on this one's evidence.
- **Self-closing form is significant.** `<x/>` and `<x></x>` are the same element and different
  bytes; `el()` vs `voidEl()` decides it by arity so that it cannot drift on a value.
- **Intra-tag whitespace is significant.** The space in `<c:xMode val="edge" />` is inside the
  tag, not between elements, so it is outside this change's claim and is frozen.
- **Whitespace relaxes only where it cannot be content.** An element with no element children is
  frozen whatever it holds, so a `<c:v> </c:v>` carrying a single significant space is safe from
  a blanket strip; so is anything with a non-whitespace text child (mixed content), and so is
  anything named in the module's text-bearing list.

The prover was made to fail on purpose before it was trusted, per the same doctrine the
LibreOffice render oracle is held to. `test/scripts/xml-equivalence.test.js` carries the red
cases (changed text, changed attribute, reordered siblings, reordered attributes, changed
escaping, changed self-closing form, whitespace eaten from inside a text leaf), and a separate
run planted thirteen such changes into a real 7.8 KB `chart1.xml` and confirmed each one was
caught.

It earned that during this change rather than after it: the codemod's first pass silently turned
`<c:layoutTarget val="inner" />` into `<c:layoutTarget val="inner"/>` across the manual-layout
block, and `prove-whitespace` reported it. A human reading a 57-part whitespace diff would not
have. Those seven sites keep their `closePrefix` and say why.

## Scope

- **In:** `chart-axes.ts`, `chart-parts.ts`, `chart-xml.ts`, `plot-bubble.ts`, `plot-cat-axis.ts`,
  `plot-pie.ts`, `plot-scatter.ts`, `plot-stock.ts`, `plot-surface.ts`: everything reaching
  `ppt/charts/chartN.xml`.
- **Out:** `embed-xlsx.ts`. Its three prefixes are in the embedded workbook's `[Content_Types].xml`
  and rels, a different part set, and the xlsx string table is the one place in this corpus where
  whitespace inside an element is routinely content.
- **Out:** attribute-order unification, and the doubled space before `b=` that an empty `sz`
  interpolation leaves in `chart-parts.ts`. Both are inert, neither is inter-element whitespace,
  and each wants its own change and its own evidence.

Verified afterwards by `test:schema`, the LibreOffice render oracle (`test:lo`) and the desktop
PowerPoint smoke (`test:com`), the oracles the byte gate was only ever a cheap proxy for. The
other 1115 parts stayed byte-identical, which is what makes "contained to the chart emitters" a
measurement rather than an intention.

## Using `prove-whitespace` again

Don't, without adding a section here first. It is not a looser `check`, and reaching for it
because `check` went red is the misuse it is shaped to resist: `check` stays the gate for every
other write-side refactor. A gate that admits exceptions stops being a gate and becomes a
judgement call. The carve-out is that this exception is itself checked by a program, and that
property only survives if each use is written down.
