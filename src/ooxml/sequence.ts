/**
 * Schema child-sequence order for the OOXML complexTypes both halves of the library insert into.
 *
 * **What these are for.** `getOrAddChild(parent, qname, before)` / `insertInOrder(parent, node,
 * before)` place a newly created child before the first sibling named in `before`. So `before`
 * must list exactly the children that legally *follow* the one being inserted. Get it wrong and
 * the part is schema-invalid — which PowerPoint reports as a corrupt file, not as a bad option,
 * so the failure surfaces far from its cause with no compile-time signal.
 *
 * **Why they are derived, not written out.** These successor lists used to be hand-maintained,
 * and the same four had drifted into two copies apiece (`read/api/shapes/oxml.ts` and
 * `read/oxml/theme.ts`), including one pair whose names were transposed (`SPPR_AFTER_XFRM` vs
 * `SPPR_XFRM_AFTER`) while the contents stayed identical. Declaring each complexType's sequence
 * **once** and slicing the successors out of it removes that class of bug outright: there is a
 * single ordering to be right about, and adding a child to a sequence updates every successor
 * list that follows it in one edit. This is the technique `TCPR_AFTER` already used for
 * `CT_TableCellProperties`; this module generalizes it.
 *
 * **Why this module is neither `gen/` nor `read/`.** Document order is a fact about the schema,
 * not about writing or reading it. `read/` inserts into a live DOM and `gen/` emits strings, but
 * both are bound by the same ECMA-376 sequence.
 */

/**
 * One step of a schema sequence: a single element, or a choice group whose members are mutually
 * exclusive. A group is spelled as an array so {@link successorsOf} can treat every member as
 * occupying the same slot — the successors of `a:solidFill` are what follows the *whole*
 * `EG_FillProperties` group, not its sibling choices.
 */
type SequenceStep = string | readonly string[]

/**
 * `EG_FillProperties` — the mutually exclusive fill choices, in schema order.
 * @see ECMA-376 Part 1 §20.1.8 (DrawingML fill group)
 */
export const FILL_CHOICES = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'] as const

/** `EG_EffectProperties` — the two mutually exclusive effect containers. */
const EFFECT_CHOICES = ['a:effectLst', 'a:effectDag'] as const

/** `EG_Geometry` — the two mutually exclusive geometry choices. */
const GEOMETRY_CHOICES = ['a:custGeom', 'a:prstGeom'] as const

/**
 * Everything that legally follows `member` in `sequence`, flattened.
 *
 * `member` may name a single element or any member of a choice group; either way the result
 * starts after the step it occupies, so a choice member never lists its own siblings. Throws on
 * an unknown name rather than returning `[]` — an empty successor list silently *appends*, which
 * is exactly the corruption these lists exist to prevent, so a typo must fail loudly and at
 * module-load time rather than produce a subtly out-of-order part at run time.
 */
function successorsOf(sequence: readonly SequenceStep[], member: string): string[] {
	const index = sequence.findIndex((step) => (typeof step === 'string' ? step === member : step.includes(member)))
	if (index < 0) throw new Error(`successorsOf: ${member} is not part of the given sequence`)
	return sequence.slice(index + 1).flat()
}

/**
 * `CT_ShapeProperties` (`p:spPr`), in declaration order.
 * @see ECMA-376 Part 1 §20.1.2.2.35
 */
const SPPR_SEQUENCE: readonly SequenceStep[] = [
	'a:xfrm',
	GEOMETRY_CHOICES,
	FILL_CHOICES,
	'a:ln',
	EFFECT_CHOICES,
	'a:scene3d',
	'a:sp3d',
	'a:extLst',
]

/**
 * `CT_GroupShapeProperties` (`p:grpSpPr`), in declaration order. A group carries no `a:ln` and
 * no `a:sp3d` — which is exactly why it needs its own successor lists rather than borrowing the
 * shape ones.
 * @see ECMA-376 Part 1 §20.1.2.2.19
 */
const GRPSPPR_SEQUENCE: readonly SequenceStep[] = ['a:xfrm', FILL_CHOICES, EFFECT_CHOICES, 'a:scene3d', 'a:extLst']

/**
 * `CT_LineProperties` (`a:ln`), in declaration order.
 * @see ECMA-376 Part 1 §20.1.2.2.24
 */
const LN_SEQUENCE: readonly SequenceStep[] = [
	FILL_CHOICES,
	['a:prstDash', 'a:custDash'],
	['a:round', 'a:bevel', 'a:miter'],
	'a:headEnd',
	'a:tailEnd',
	'a:extLst',
]

/**
 * `CT_TextCharacterProperties` (`a:rPr`/`a:defRPr`/`a:endParaRPr`), in declaration order.
 * @see ECMA-376 Part 1 §21.1.2.3.9
 */
const RPR_SEQUENCE: readonly SequenceStep[] = [
	'a:ln',
	FILL_CHOICES,
	EFFECT_CHOICES,
	'a:highlight',
	['a:uLnTx', 'a:uLn'],
	['a:uFillTx', 'a:uFill'],
	'a:latin',
	'a:ea',
	'a:cs',
	'a:sym',
	'a:hlinkClick',
	'a:hlinkMouseOver',
	'a:rtl',
	'a:extLst',
]

/**
 * `CT_TableCellProperties` (`a:tcPr`), in declaration order.
 * @see ECMA-376 Part 1 §21.1.3.17
 */
export const TCPR_SEQUENCE: readonly string[] = [
	'a:lnL',
	'a:lnR',
	'a:lnT',
	'a:lnB',
	'a:lnTlToBr',
	'a:lnBlToTr',
	'a:cell3D',
	...FILL_CHOICES,
	'a:headers',
	'a:extLst',
]

/**
 * For each `a:tcPr` child, the children that must follow it — i.e. what a new one is inserted
 * *before*. Every child gets an entry, so a caller can look one up by qname.
 */
export const TCPR_AFTER: Record<string, string[]> = Object.fromEntries(
	TCPR_SEQUENCE.map((qname, index) => [qname, TCPR_SEQUENCE.slice(index + 1)])
)

// --- p:spPr / p:grpSpPr ------------------------------------------------------

/** Successors of `a:xfrm` inside `p:spPr`. */
export const SPPR_AFTER_XFRM = successorsOf(SPPR_SEQUENCE, 'a:xfrm')
/** Successors of a fill choice inside `p:spPr`. */
export const SPPR_FILL_AFTER = successorsOf(SPPR_SEQUENCE, 'a:solidFill')
/** Successors of `a:ln` inside `p:spPr`. */
export const SPPR_LN_AFTER = successorsOf(SPPR_SEQUENCE, 'a:ln')
/** Successors of an effect container inside `p:spPr`. */
export const SPPR_EFFECT_AFTER = successorsOf(SPPR_SEQUENCE, 'a:effectLst')
/** Successors of `a:scene3d` inside `p:spPr`. */
export const SPPR_SCENE3D_AFTER = successorsOf(SPPR_SEQUENCE, 'a:scene3d')
/** Successors of `a:sp3d` inside `p:spPr`. */
export const SPPR_SP3D_AFTER = successorsOf(SPPR_SEQUENCE, 'a:sp3d')

/** Successors of `a:xfrm` inside `p:grpSpPr`. */
export const GRPSPPR_AFTER_XFRM = successorsOf(GRPSPPR_SEQUENCE, 'a:xfrm')
/** Successors of a fill choice inside `p:grpSpPr`. */
export const GRPSPPR_FILL_AFTER = successorsOf(GRPSPPR_SEQUENCE, 'a:solidFill')

/**
 * Successors of `p:spPr` within `p:sp` (and before `p:style` within `p:pic` / `p:cxnSp`).
 * `p:blipFill` and the `p:nv*Pr` wrapper precede `p:spPr` and so are deliberately absent.
 */
export const SHAPE_AFTER_SPPR = ['p:style', 'p:txBody']

// --- a:ln / a:rPr ------------------------------------------------------------

/** Successors of a fill choice inside `a:ln`. */
export const LN_FILL_AFTER = successorsOf(LN_SEQUENCE, 'a:solidFill')
/** Successors of a fill choice inside `a:rPr` (and `a:defRPr`/`a:endParaRPr`). */
export const RPR_FILL_AFTER = successorsOf(RPR_SEQUENCE, 'a:solidFill')
