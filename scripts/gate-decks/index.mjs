/**
 * The gate-deck registry.
 *
 * `scripts/byte-identity.mjs` builds these alongside the showcase decks. They exist only to
 * be diffed — see `./README.md` for why they are not showcases, and why a gate deck may
 * only ever grow.
 *
 * Same one-list discipline as `demos/showcases/lib/showcases.mjs`: two hand-maintained
 * lists drift, and the way they drift is silent — the gate keeps passing on a subset while
 * the emitter being refactored goes undiffed.
 */
import { gateDeck as chartMatrix } from './chart-matrix.mjs'
import { gateDeck as shapeMatrix } from './shape-matrix.mjs'

/** Every gate deck, in the order `byte-identity.mjs` builds them. */
export const GATE_DECKS = [chartMatrix, shapeMatrix]
