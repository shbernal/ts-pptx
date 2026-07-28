/**
 * The showcase registry.
 *
 * Both the deck build (`build.mjs`) and the byte-identity gate (`scripts/byte-identity.mjs`)
 * enumerate the decks from here. Two hand-maintained lists would drift, and the way they
 * drift is silent: the gate would keep passing on a subset while a new deck's parts — quite
 * possibly the ones exercising the emitter being refactored — go undiffed.
 */
import { showcase as fieldNotes } from "../field-notes/index.mjs";
import { showcase as quarterlyReview } from "../quarterly-review/index.mjs";

/** Every showcase deck, in the order `pnpm demos:build` builds them. */
export const SHOWCASES = [quarterlyReview, fieldNotes];
