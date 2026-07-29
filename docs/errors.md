---
doc-schema-version: 1
title: "Errors"
summary: "The error taxonomy ts-pptx throws: five classes, a stable code on each, and what is and is not part of the contract."
read_when:
  - Catching and classifying a failure from ts-pptx
  - Deciding whether to retry, substitute, or reject on an error
  - Adding a new throw site to the library
doc_type: "guide"
---

# Errors

ts-pptx never silently coerces input it cannot use. Where the library can still produce a valid
package by ignoring, clamping, or falling back, it emits a [diagnostic](diagnostics.md) and carries
on. Where it cannot, it throws. This page is about the second kind.

Every failure the library raises is a `TsPptxError` carrying a stable `code`:

```ts
class TsPptxError extends Error {
	readonly code: ErrorCode // e.g. 'coord/non-finite'
	readonly detail?: Record<string, unknown> // structured context, when a site has any
}
```

## The class and the code are API. The message is not.

You may branch on `instanceof` and on `code`, and treat both as part of the package's contract:
adding a code is back-compatible, and removing or renaming one is a breaking change with a
`CHANGELOG.md` entry.

The message is prose meant for a human reading a stack trace. It is free to improve — reworded,
expanded, given a better example — in any release, including patch releases. Do not parse it, and
do not assert on it in tests.

Every error remains an `instanceof Error`, so a `catch` block that only knows about `Error` keeps
working unchanged.

## The five classes

The classes are a deliberately coarse bucket — they answer *"whose problem is this?"* — and the
`code` carries the specificity.

| class | the failure is | who fixes it |
| --- | --- | --- |
| `InvalidOptionError` | you passed something unusable | your code |
| `UnsupportedFeatureError` | a well-formed request this build, runtime, or shape cannot express | your expectations, or the environment |
| `PackageReadError` | the input bytes are not a readable package | the input file |
| `MediaError` | a referenced image, font, or A/V resource would not load or decode | the resource |
| `InternalError` | an invariant of the library itself did not hold | ts-pptx — please file a bug |

The split is what lets a batch job react rather than just log:

```ts
import { MediaError, PackageReadError, InvalidOptionError } from '@shbernal/ts-pptx'

try {
	await buildDeck(spec)
} catch (err) {
	if (err instanceof MediaError) return retryWithPlaceholderAsset(spec) // transient / fixable asset
	if (err instanceof PackageReadError) return rejectUpload(err.code) // the user's file is bad
	if (err instanceof InvalidOptionError) throw err // our bug — fail the job loudly
	throw err
}
```

`InvalidOptionError` is by far the largest group, and that is deliberate: the project's policy is to
throw rather than coerce, because emitting a degenerate result (a zero-size shape, a silently
dropped option) hides the mistake instead of surfacing it.

## Codes are shared with diagnostics

Errors and warnings draw on **one** vocabulary (`src/codes.ts`). A condition keeps the same code
whichever way it reaches you, so `coord/non-finite` means "a coordinate was `NaN`/`Infinity`"
whether it arrived as a thrown `InvalidOptionError` or as a `Diagnostic` on your handler. A consumer
that special-cases a condition only has to learn one string for it.

Each code belongs to exactly one class, and that pairing is type-enforced —
`new MediaError('coord/non-finite', …)` does not compile. If you are narrowing on both, `code` alone
is sufficient; the class is the coarse view of the same fact.

## The originating failure is preserved

Where the library wraps a lower-level failure, the original is kept on the standard `cause` property
rather than flattened into the message:

```ts
catch (err) {
	if (err instanceof PackageReadError) console.error(err.code, err.cause)
}
```

## Importing the classes

The classes and the code types are re-exported from every entry point — `@shbernal/ts-pptx` and each
subpath (`/read`, `/zip`, `/math`, …). They all resolve to one shared module, so `instanceof` works
regardless of which subpath you imported from and which subpath threw.

## Adding one

A new throw site must name its condition in `src/codes.ts`, under the union belonging to the class
that will carry it, before it will compile. That is the enforcement mechanism for keeping the
vocabulary curated rather than accumulated — reuse an existing code when the condition is genuinely
the same, even if the wording differs and even if it is raised from a different entry point.

Do not add a class per throw site. Five classes is the whole taxonomy; specificity belongs in the
code.

Write the message without an `ERROR:` / `ERROR!` prefix of its own — the class name already labels
the failure in every stack trace and console rendering.

Throwing is for a request the library ends up discarding. When it can carry on and still produce
what the caller would recognise as their deck, warn instead — see
[Warn or throw?](diagnostics.md#warn-or-throw) for where the line sits.
