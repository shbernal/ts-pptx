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

## `InternalError` tells you to report it

`InternalError` is the one class that adds to its own message, appending a pointer to the issue
tracker below the invariant that broke:

```
makeXmlSlideRel: no slide at index 3

This is a bug in ts-pptx, not in your deck or your code. Please report it:
https://github.com/shbernal/ts-pptx/issues/new?template=agent-report.yml
```

That is a message, so it is still not API — do not assert on it. It lives in the constructor rather
than at each throw site so a site added later cannot forget it, and it is on this class alone
because the other four are routine outcomes of bad input or a bad call. A "report this" banner on
every malformed package would train you to skip the line, including the one time it always means
something.

### Which failures are worth reporting

`InternalError` always is. For the rest, the test is whether the environment disagrees with us —
the project's supported bar is *"the output opens cleanly in Microsoft PowerPoint"*, and it reads in
both directions:

| you saw | report it when |
| --- | --- |
| `PackageReadError` | the file opens cleanly in PowerPoint — then we are the ones who cannot read it |
| `MediaError` | the image, font, or A/V asset loads fine in other tools |
| `UnsupportedFeatureError` | PowerPoint can plainly express the thing, so this is a gap rather than a limit |
| `InvalidOptionError` | the deck it refused is one PowerPoint can express |
| no error at all | PowerPoint repairs or misrenders the output, or a round trip loses a construct |

A skill ships inside the package that walks through triage, reducing the failure to a script that
builds its own deck, and filing — including the rule that a deck from a real project never goes to a
public tracker:

```bash
npx skills add ./node_modules/@shbernal/ts-pptx   # offline, matches your installed version
npx skills add shbernal/ts-pptx                   # or straight from the repo
```

You do not need it to report something. <https://github.com/shbernal/ts-pptx/issues> is open.

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
