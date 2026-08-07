import fs from 'node:fs'
import path from 'node:path'
import { diffParts, explodePackage, listParts } from '../../scripts/pptx-parts.mjs'
import { ROOT } from '../../scripts/script-utils.mjs'
import { expect, test } from './fixtures.mjs'
import { buildDeckInHarness, buildDeckInNode, NODE_ASSETS, openHarness, packageBytes } from './helpers.mjs'

/**
 * `loadFontData` (src/runtime/browser.ts) in a real browser.
 *
 * One deck reaches both of its call sites: `registerFontMetrics`, where the bytes go to
 * opentype.js and come back as a baked `fontScale`, and `embedFont`, where they go into
 * the package as an `/ppt/fonts/` part. The two are complementary — the first proves the
 * bytes *parsed*, the second proves they arrived whole — and together they mean a
 * cross-runtime comparison of this deck says everything worth saying about the function.
 *
 * Silkscreen (OFL, committed under test/read/fixtures/fonts) is the same face the Node
 * measured-fit suite uses, for the same reason: it is a real font with a real cmap that
 * ships with the repo, so the metrics path is deterministic on every machine instead of
 * depending on what fonts happen to be installed.
 */

const OUT_ROOT = path.join(ROOT, '.tmp', 'browser-adapter')

test.beforeEach(async ({ page }) => {
	await openHarness(page)
})

test('loadFontData: a font fetched over HTTP measures and embeds exactly as one read off disk', async ({ page }) => {
	const browser = await buildDeckInHarness(page, 'fonts')
	expect(browser.ok, `the harness failed to build the "fonts" deck: ${browser.message}`).toBe(true)
	const nodeBase64 = await buildDeckInNode('fonts')

	const nodeDir = await explodePackage(packageBytes(nodeBase64), path.join(OUT_ROOT, 'fonts', 'node'))
	const browserDir = await explodePackage(packageBytes(browser.base64), path.join(OUT_ROOT, 'fonts', 'browser'))

	// A byte difference here would be one of two failures, and the named part says which:
	// `ppt/fonts/…` means the fetched bytes are not the file, `ppt/slides/slide1.xml`
	// means they parsed to different metrics and baked a different `fontScale`.
	const diffs = diffParts(nodeDir, browserDir)
	expect(diffs, `node vs browser font deck differs:\n  ${diffs.join('\n  ')}`).toEqual([])

	// Same guard as the media spec: agreeing with Node is not the same as agreeing with
	// the font file, and only the second rules out two identical wrong answers.
	const fontParts = listParts(browserDir).filter((part) => part.startsWith('ppt/fonts/'))
	expect(fontParts).toHaveLength(1)
	expect(fs.readFileSync(path.join(browserDir, fontParts[0])).equals(fs.readFileSync(NODE_ASSETS.font))).toBe(true)

	// And the metrics were actually consulted: `fit:'shrink'` bakes a real scale only when
	// a registered face could measure the text. Without it the emitter falls back to a bare
	// `<a:normAutofit/>`, which would still have compared equal across runtimes — an
	// agreement about nothing.
	const slideXml = fs.readFileSync(path.join(browserDir, 'ppt/slides/slide1.xml'), 'utf8')
	expect(slideXml).toMatch(/<a:normAutofit fontScale="\d+"/)
})

test('loadFontData: a 404 font URL rejects with the adapter code', async ({ page }) => {
	const outcome = await buildDeckInHarness(page, 'missingFont')
	expect(outcome.ok).toBe(false)
	// `registerFontMetrics` fails where it is called, not at export, so there is no
	// pipeline wrapper here — the adapter's own error is the one that surfaces.
	expect(outcome.code).toBe('font/fetch-failed')
	expect(outcome.message).toContain('no-such-font.ttf')
})
