/**
 * ts-pptx: hex colour text, in the one spelling both halves use.
 *
 * A caller may write a colour with or without a leading `#`, and the library accepts both
 * everywhere — which means every site that validates or parses one strips it first. Those
 * strips had drifted into three spellings across six sites, and one of them was not
 * anchored: `.replace('#', '')` takes the `#` out of `'FF00#0'` and hands `'FF000'` on, which
 * then fails the six-digit test and takes a different branch than the anchored forms would.
 * No caller is known to reach that, which is exactly why it wants one spelling rather than a
 * test at each site.
 *
 * This lives at the root rather than under `gen/` or `read/` because both halves need it:
 * the emitters strip on the way in, the readers strip on the way back out.
 */

/** Strip a single leading `#` from a colour, leaving any other character alone. */
export function stripHash(value: string): string {
	return value.startsWith('#') ? value.slice(1) : value
}
