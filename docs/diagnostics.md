---
doc-schema-version: 1
title: "Diagnostics"
summary: "How ts-pptx reports non-fatal problems, and how to route, silence, or escalate them."
read_when:
  - Silencing or redirecting library warnings
  - Reacting programmatically to a specific warning condition
  - Adding a new warning to the library
doc_type: "guide"
---

# Diagnostics

ts-pptx never fails silently on input it cannot use. Where a problem is fatal it throws; where the
library can still produce a valid package by ignoring, clamping, or falling back, it emits a
**diagnostic** and carries on. This page is about the second kind.

## The shape of a diagnostic

```ts
interface Diagnostic {
	code: DiagnosticCode // e.g. 'chart/non-finite-value'
	message: string // human-readable explanation
	detail?: Record<string, unknown> // structured context, when a site has any
}
```

`detail` is optional and sparse. Today it is filled in by every rejection of an OOXML `ST_`
enum value: a line dash type, a table cell's 3-D preset, a table's horizontal overflow. Those
carry `received`, what you passed, and `valid`, the whole legal list. Read it defensively.
Which codes carry which keys is not part of the contract, and most codes carry nothing at
all.

**The `code` is API. The `message` is not.**

A code is a stable identifier of a *condition*, in `area/condition` form. Branch on it, log
it, count it, treat it as part of the package's contract. Adding a code is back-compatible.
Removing or renaming one is a breaking change, with a `CHANGELOG.md` entry to match.

The message is prose meant for a human reading a build log. It is free to improve (reworded,
expanded, given a better example) in any release, including patch releases. Do not parse it, and
do not assert on it in tests.

`DiagnosticCode` is a closed union, so your editor will complete the available codes and TypeScript
will reject a typo.

## Routing them

By default each diagnostic is one prefixed line on `console.warn`:

```
ts-pptx: text `columns` must be a number 1-16 (ignoring value)
```

Install a handler to take that over:

```ts
import { setDiagnosticHandler } from 'pptx-ts'

setDiagnosticHandler((d) => logger.warn({ code: d.code }, d.message))
```

Pass `null` to restore the console default:

```ts
setDiagnosticHandler(null)
```

To silence the library entirely, install a handler that does nothing:

```ts
setDiagnosticHandler(() => {})
```

### It is process-global, not per-presentation

The handler is module state, not presentation state. That is a deliberate trade. Diagnostics
come out of a tree of free functions across `gen/**` with no presentation in scope, and
threading a handler through every one of those signatures would be a far larger and worse
change than the problem warrants.

The practical consequence: a process building several decks concurrently cannot attribute a
diagnostic to one of them. If that matters, either correlate on `code`, or set and clear the
handler around each build so the two never overlap.

## Escalating a condition to an error

There is no separate "strict mode" switch. A handler that throws is one:

```ts
setDiagnosticHandler((d) => {
	// A bare number is always inches. In this codebase that is always a mistake.
	if (d.code === 'coord/bare-number-is-inches') throw new Error(d.message)
})
```

The throw propagates out of whatever library call emitted the diagnostic. This composes with
whatever policy you want (escalate one code, escalate everything under `chart/`, escalate in CI
and warn locally) without the library having to model any of it.

## Repeated conditions

Some conditions would otherwise flood a log: the same out-of-range `fontSize` on every cell of a
large table, say. Those are emitted once per distinct code **and** message, so a repeat of the
*same* offending value is reported once while a *different* value reports on its own.

"Once" is process-global, the same scope the handler has. The consequence is not the same one,
though, and it is worth knowing if you build more than one deck per process: the second deck is
silent about anything the first already reported. That reads as "no problems found" rather than
as "already mentioned". Clear the record between builds:

```ts
import { resetDiagnosticState } from 'pptx-ts'

resetDiagnosticState()
```

It also bounds the memory. Most of these messages interpolate the offending value, so a service
building decks from user input accumulates one entry per distinct bad value with nothing to
release them.

The handler is not affected: the two are reset independently, because a host usually installs
its handler once and wants it to outlive any single build.

## Adding one

A new warning site must name its condition in `DiagnosticCode` (`src/codes.ts`) before it
will compile. That is what keeps the vocabulary curated rather than merely accumulated. Reuse
an existing code when the condition is genuinely the same, even if the wording differs, even
if it is reported from a different entry point.

Write the message without a `ts-pptx:` prefix; the default handler stamps that.

### Warn or throw?

Ask what the library does *next*, not how bad the input looks:

- If it can carry on and still produce something the caller would recognise as their deck, **warn**.
  Clamping the value, ignoring the option, falling back to a default glyph all qualify.
- If the request is discarded and the deck comes out missing what was asked for, **throw**. A deck
  that opens cleanly and is quietly missing an image is worse than a failed build.

`addImage()` and a picture bullet split on exactly this line. An image with no usable source
has nothing to place, so it throws. A bullet image with no usable source falls back to a `•`,
so it warns. See [Errors](./errors.md) for the thrown half.

### The rule applied: an out-of-range number

Every option stated as a percentage, a size, or a spacing has a range its OOXML attribute
allows, and one answer for a value outside it:

- **Finite and out of range:** clamp to the nearest bound and warn. There is a legal value next
  to the one asked for, so the deck comes out recognisable. `transparency: 120` paints at 100,
  `bullet.size: 500` draws the glyph at 400%, `fit.fontScale: 150` scales at 100.
- **Not a number at all** (`NaN`, a string, `undefined` where a number is required): throw.
  There is no nearest legal value, so the request is discarded, and clamping it would put
  `val="NaN"` in the package.

`Infinity` belongs to the first group, not the second: it has a nearest bound like any other
out-of-range number.

This is worth stating as one rule because the third option is the tempting one, and it is
wrong. *Rejecting* a finite out-of-range value and emitting nothing looks strict. It is not.
It discards the request and reports it as a warning, which is exactly the pairing the rule
above exists to rule out. The caller reads a warning and gets a deck whose bullet is silently
back at its inherited size. `bullet.size`, `fit.fontScale`, `fit.lnSpcReduction`,
`shadow.transparency` and `shadow.angle` all behaved that way until they were routed through
the shared clamp.

### The rule applied: an empty colour string

Every option that takes a colour has one answer for `''`, and it is not "paint nothing".

A paint has three states and three spellings: omit the option, `{ type: 'inherit' }`, or `{ type: 'none' }`.
`''` is not a fourth. It reaches the library from the caller's own missing value, an unset
template field or a `row.accent` on a row that has none. So it is **reported under
`color/empty-string`, then resolved to whatever omitting the option resolves to**.

What omission resolves to is the surface's own rule, and the surfaces disagree, deliberately:

| Option | Omitted, and so also `''` |
| --- | --- |
| a text box's `fill` | `<a:noFill/>`, because an unfilled box is what a text box is |
| a shape's `line.color` | the shape line default |
| a chart's `dataLabelColor` | `DEF_FONT_COLOR` |
| a slide `background` | inherits from the layout |

That is the whole rule. It is worth stating because both alternatives are tempting and both
are wrong. *Painting a default* on an empty string is what the colour element builder used to
do, falling back to `000000`. It puts visible black on a shape whose caller expected to keep
the theme's paint. *Emitting nothing* looks harmless, but it is a third state again. On the
paths where omission means `<a:noFill/>` or a stated default, silence is not what the caller
would have got by leaving the option out, so `''` and omission drift apart.

One slot is exempt, because it has no absent state to resolve to: a gradient stop, a duotone
half and a `buClr` each *require* a colour, and there is nothing to inherit. Those still paint
`DEF_FONT_COLOR`, and report it under the same code.

Reporting a condition by calling `console.log` or `console.error` directly is neither, and
oxlint rejects it under `eslint/no-console`. Such a line cannot be captured, silenced, or
branched on. There are three exemptions: `diagnostics.ts`, which owns the default handler,
and the two `verbose: true` table tracers (`gen/table/autopage.ts`, `gen/table/html-dom.ts`),
whose output reports no condition and is opt-in anyway.
