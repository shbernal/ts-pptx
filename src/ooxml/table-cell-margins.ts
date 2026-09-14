/**
 * A table cell's default text insets in EMU: the schema defaults of `a:tcPr/@marL`, `@marR`
 * (91440, 0.1in) and `@marT`, `@marB` (45720, 0.05in) on `CT_TableCellProperties`.
 *
 * Unlike the body insets in `body-insets.ts` these are real XSD defaults, so they are transcribed
 * from the schema rather than derived. A cell that states one side leaves the other three at these
 * values, not at zero.
 *
 * A module of its own, importing nothing, because `constants-internal.ts` derives the write path's
 * inch defaults from it and is reached by every entry. Taken from `body-insets.ts`, the record
 * brought that module's other exports into the read entry with it.
 */
export const TABLE_CELL_MARGIN_DEFAULTS_EMU = { left: 91440, right: 91440, top: 45720, bottom: 45720 } as const
