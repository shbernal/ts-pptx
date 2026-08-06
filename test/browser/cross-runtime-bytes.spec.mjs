import fs from 'node:fs'
import path from 'node:path'
import { expect, test } from '@playwright/test'
import { diffParts, explodePackage, listParts, loadShowcase } from '../../scripts/pptx-parts.mjs'
import { ROOT } from '../../scripts/script-utils.mjs'
import { buildDeckInBrowser } from './helpers.mjs'

/**
 * Cross-runtime byte identity — the assertion this lane exists for.
 *
 * `demos/vite-demo` builds the same showcase module as `pnpm demos:build
 * quarterly-review`, and `src/zip.ts` pins `FIXED_MTIME`, so the deck a browser assembles
 * and the deck Node assembles should agree part for part. That is a far stronger claim
 * than the structural smoke test next door: it says the *whole* emission core — every
 * serializer, the zip writer, part ordering, relationship numbering — is runtime-invariant,
 * in one comparison. A runtime-dependent code path anywhere in `src/gen/` shows up here as
 * a named part.
 *
 * The comparison is the byte-identity gate's, not a second one: same explode, same
 * normalizers, same diff (`scripts/pptx-parts.mjs`). Three values legitimately differ
 * between any two runs — the core.xml timestamps and the two `Math.random` GUIDs
 * (`p14:section` ids, `c16:uniqueId`) — and are normalized there. Everything else is a
 * real finding.
 *
 * `quarterly-review` and not `field-notes` because the demo only offers the former: the
 * latter loads photographs and a video from disk by path, which a browser cannot do.
 * That also means this comparison never crosses `loadMedia`, whose two implementations
 * return different-but-equivalent strings (Node: raw base64; browser: a `FileReader`
 * data URI). Covering that function is Tier 2, step 7.
 */

const SLUG = 'quarterly-review'
const OUT_ROOT = path.join(ROOT, '.tmp', 'browser-parity')

test('a browser-built deck is byte-identical to the Node-built one', async ({ page }) => {
	const { bytes: browserBytes } = await buildDeckInBrowser(page)

	const showcase = await loadShowcase(SLUG)
	const nodeFile = path.join(OUT_ROOT, showcase.fileName)
	fs.mkdirSync(OUT_ROOT, { recursive: true })
	await showcase.build(nodeFile)
	const nodeBytes = new Uint8Array(fs.readFileSync(nodeFile))

	// Left on disk deliberately: on a failure the two exploded trees are what makes the
	// named part diffable, exactly as `.tmp/byte-identity/` is for the Node gate.
	const nodeDir = await explodePackage(nodeBytes, path.join(OUT_ROOT, 'node'))
	const browserDir = await explodePackage(browserBytes, path.join(OUT_ROOT, 'browser'))

	// Guard against a comparison that passes because one side produced nothing.
	expect(listParts(nodeDir).length).toBeGreaterThan(50)

	const diffs = diffParts(nodeDir, browserDir)
	expect(diffs, `node vs browser package differs:\n  ${diffs.join('\n  ')}`).toEqual([])
})
