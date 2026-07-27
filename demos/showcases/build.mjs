#!/usr/bin/env node
/**
 * Build the showcase decks.
 *
 *   pnpm demos:build                    both decks
 *   pnpm demos:build field-notes        one deck, by slug
 *
 * These decks are showcases, not tests. Nothing here asserts anything about the output and
 * no verification aggregate runs it — if a deck breaks, the fix belongs in the deck. The
 * published-package contract is covered by `pnpm run check:package`.
 */
import { mkdir } from "node:fs/promises";
import path from "node:path";
import { OUTPUT_DIR, SHOWCASES_DIR } from "./lib/assets.mjs";
import { showcase as fieldNotes } from "./field-notes/index.mjs";
import { showcase as quarterlyReview } from "./quarterly-review/index.mjs";

const SHOWCASES = [quarterlyReview, fieldNotes];

const requested = process.argv.slice(2);
const unknown = requested.filter((slug) => !SHOWCASES.some((s) => s.slug === slug));
if (unknown.length > 0) {
	console.error(`unknown showcase: ${unknown.join(", ")}`);
	console.error(`available: ${SHOWCASES.map((s) => s.slug).join(", ")}`);
	process.exit(1);
}

const selected = requested.length > 0 ? SHOWCASES.filter((s) => requested.includes(s.slug)) : SHOWCASES;

await mkdir(OUTPUT_DIR, { recursive: true });

for (const showcase of selected) {
	const started = Date.now();
	const outFile = path.join(OUTPUT_DIR, showcase.fileName);
	await showcase.build(outFile);
	console.log(`  ${showcase.title}`);
	console.log(`    ${path.relative(SHOWCASES_DIR, outFile)}  (${Date.now() - started}ms)`);
}

console.log(`\nDone. Decks are in ${OUTPUT_DIR}`);
