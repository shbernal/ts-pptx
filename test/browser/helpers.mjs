import { readFile } from 'node:fs/promises'
import { expect } from '@playwright/test'

/**
 * Drive `demos/vite-demo` through one deck build and hand back the downloaded bytes.
 *
 * This is the whole point of the browser lane: the demo imports the *same* showcase
 * module `pnpm demos:build quarterly-review` runs, so what comes back here is the deck
 * as a browser assembled it — through `src/runtime/browser.ts`'s `writeFile`, the
 * object-URL `<a download>` path that no Node test can reach.
 */
export async function buildDeckInBrowser(page) {
	await page.goto('./')

	const button = page.getByRole('button', { name: /^Build / })
	await expect(button).toBeEnabled()

	const downloadPromise = page.waitForEvent('download')
	await button.click()

	// Wait on the page's own outcome before the download, so a build that threw is
	// reported as its error message rather than as a 30s "no download event" timeout.
	// `App.tsx` renders `role="status"` on success and `role="alert"` on failure.
	const outcome = page.locator('[role="status"], [role="alert"]')
	await expect(outcome).toBeVisible({ timeout: 30_000 })
	const alert = page.getByRole('alert')
	if (await alert.count()) throw new Error('the demo failed to build the deck: ' + (await alert.innerText()))

	const download = await downloadPromise
	const file = await download.path()
	return { bytes: new Uint8Array(await readFile(file)), fileName: download.suggestedFilename() }
}
