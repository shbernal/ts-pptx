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

import { EMBEDDED_FONT_SLOTS } from '../embedded-fonts.js'

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
 * `CT_Shape` (`p:sp`), in declaration order.
 *
 * `p:pic` (`CT_Picture`) and `p:cxnSp` (`CT_Connector`) are not this type, but their successors
 * of `p:spPr` are a *subset* of this one's in the same relative order — both run `p:spPr`,
 * `p:style`, `p:extLst` and neither has a `p:txBody`. So {@link SHAPE_AFTER_SPPR} serves all
 * three: naming an element that cannot occur in a given parent is inert, because the successor
 * list is only ever matched against children that are actually there.
 * @see ECMA-376 Part 1 §19.3.1.43
 */
const SP_SEQUENCE: readonly SequenceStep[] = ['p:nvSpPr', 'p:spPr', 'p:style', 'p:txBody', 'p:extLst']

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
 * Successors of `p:spPr` within `p:sp`, and equally within `p:pic` / `p:cxnSp` — see
 * {@link SP_SEQUENCE} for why one list covers all three. `p:blipFill` and the `p:nv*Pr` wrapper
 * precede `p:spPr`, so slicing the sequence drops them without anyone having to remember to.
 */
export const SHAPE_AFTER_SPPR = successorsOf(SP_SEQUENCE, 'p:spPr')

// --- a:ln / a:rPr ------------------------------------------------------------

/** Successors of a fill choice inside `a:ln`. */
export const LN_FILL_AFTER = successorsOf(LN_SEQUENCE, 'a:solidFill')
/** Successors of a fill choice inside `a:rPr` (and `a:defRPr`/`a:endParaRPr`). */
export const RPR_FILL_AFTER = successorsOf(RPR_SEQUENCE, 'a:solidFill')
/** Successors of `a:latin` inside `a:rPr` (and `a:defRPr`/`a:endParaRPr`). */
export const RPR_LATIN_AFTER = successorsOf(RPR_SEQUENCE, 'a:latin')

// --- p:presentation ----------------------------------------------------------

/**
 * `CT_Presentation` (`p:presentation`), in declaration order. Every child is a single element —
 * there are no choice groups — so the declaration is a plain string list.
 * @see ECMA-376 Part 1 §19.2.1.26
 */
const PRESENTATION_SEQUENCE: readonly SequenceStep[] = [
	'p:sldMasterIdLst',
	'p:notesMasterIdLst',
	'p:handoutMasterIdLst',
	'p:sldIdLst',
	'p:sldSz',
	'p:notesSz',
	'p:smartTags',
	'p:embeddedFontLst',
	'p:custShowLst',
	'p:photoAlbum',
	'p:custDataLst',
	'p:kinsoku',
	'p:defaultTextStyle',
	'p:modifyVerifier',
	'p:extLst',
]

/** Successors of `p:sldMasterIdLst` inside `p:presentation`. */
export const PRESENTATION_AFTER_SLD_MASTER_ID_LST = successorsOf(PRESENTATION_SEQUENCE, 'p:sldMasterIdLst')
/** Successors of `p:notesMasterIdLst` inside `p:presentation`. */
export const PRESENTATION_AFTER_NOTES_MASTER_ID_LST = successorsOf(PRESENTATION_SEQUENCE, 'p:notesMasterIdLst')
/** Successors of `p:sldIdLst` inside `p:presentation`. */
export const PRESENTATION_AFTER_SLD_ID_LST = successorsOf(PRESENTATION_SEQUENCE, 'p:sldIdLst')
/** Successors of `p:embeddedFontLst` inside `p:presentation`. */
export const PRESENTATION_AFTER_EMBEDDED_FONT_LST = successorsOf(PRESENTATION_SEQUENCE, 'p:embeddedFontLst')

/**
 * `CT_EmbeddedFontListEntry` (`p:embeddedFont`), in declaration order: the typeface identity
 * followed by one optional element per face slot. Every child is a single element, so the
 * declaration is a plain string list.
 * @see ECMA-376 Part 1 §19.2.1.9
 */
const EMBEDDED_FONT_ENTRY_SEQUENCE: readonly string[] = ['p:font', ...EMBEDDED_FONT_SLOTS.map((slot) => `p:${slot}`)]

/**
 * For each `p:embeddedFont` child, the children that must follow it — i.e. what a new one is
 * inserted *before*. Every child gets an entry, so a caller can look one up by qname.
 */
export const EMBEDDED_FONT_ENTRY_AFTER: Record<string, string[]> = Object.fromEntries(
	EMBEDDED_FONT_ENTRY_SEQUENCE.map((qname, index) => [qname, EMBEDDED_FONT_ENTRY_SEQUENCE.slice(index + 1)])
)

// --- p:pic / a:blipFill ------------------------------------------------------

/**
 * `CT_Picture` (`p:pic`), in declaration order.
 * @see ECMA-376 Part 1 §19.3.1.37
 */
const PIC_SEQUENCE: readonly SequenceStep[] = ['p:nvPicPr', 'p:blipFill', 'p:spPr', 'p:style', 'p:extLst']

/**
 * `CT_BlipFillProperties` (`a:blipFill`), in declaration order. `a:tile` and `a:stretch` are the
 * `EG_FillModeProperties` choice and share one slot, so neither lists the other as a successor.
 * @see ECMA-376 Part 1 §20.1.8.14
 */
const BLIPFILL_SEQUENCE: readonly SequenceStep[] = ['a:blip', 'a:srcRect', ['a:tile', 'a:stretch']]

/** Successors of `p:blipFill` inside `p:pic`. */
export const PIC_BLIPFILL_AFTER = successorsOf(PIC_SEQUENCE, 'p:blipFill')
/** Successors of `a:blip` inside `a:blipFill`. */
export const BLIPFILL_BLIP_AFTER = successorsOf(BLIPFILL_SEQUENCE, 'a:blip')
