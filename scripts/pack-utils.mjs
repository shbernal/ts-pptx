/**
 * Tarball helpers for the two package-boundary gates.
 *
 * Split out of `script-utils.mjs` rather than left beside `ROOT`/`run()` because the
 * audience is different. `script-utils.mjs` is imported by sixteen scripts and seven
 * browser specs, nearly all of them only for `ROOT`; `pnpm pack` logic is wanted by
 * exactly two callers (`package-lint.mjs`, `package-smoke.mjs`). Keeping them together
 * meant every browser spec that needed a repo root also loaded tarball parsing it can
 * never call.
 */

import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { ROOT, run } from './script-utils.mjs'

/**
 * Pack the package into a fresh temporary directory, hand the tarball to `fn`, and remove the
 * directory afterwards unless asked to keep it.
 *
 * Both package gates spelled this sequence out, each with a keep variable of its own, and only
 * the smoke test let the temporary root be moved. One pair now serves both:
 * `TSPPTX_PACKAGE_TMPDIR` for where the directory is made (the OS temporary directory by
 * default), and `TSPPTX_KEEP_PACKAGE_TMP=1` to leave it behind for inspection.
 * @template T
 * @param {string} label - which gate is packing, for the directory name and the keep message
 * @param {(packInfo: Awaited<ReturnType<typeof packPackage>>, tmp: string) => Promise<T>} fn -
 *   receives the pack result and the temporary directory, which it may use for its own files
 * @returns {Promise<T>}
 */
export async function withPackedTarball(label, fn) {
	const tmpRoot = process.env.TSPPTX_PACKAGE_TMPDIR || os.tmpdir()
	await fs.mkdir(tmpRoot, { recursive: true })
	const tmp = await fs.mkdtemp(path.join(tmpRoot, `.ts-pptx-${label}-`))
	try {
		return await fn(await packPackage(path.join(tmp, 'pack')), tmp)
	} finally {
		if (process.env.TSPPTX_KEEP_PACKAGE_TMP === '1') console.log(`Keeping ${label} temp directory: ${tmp}`)
		else await fs.rm(tmp, { recursive: true, force: true })
	}
}

/**
 * Pull the JSON payload out of a `pnpm pack --json` run.
 *
 * pnpm prints progress lines before the JSON, and which stream carries what has moved
 * between versions — hence the scan rather than a plain `JSON.parse`. The last
 * top-level `{`/`[` at a line start wins, because a progress line can legitimately
 * contain a brace; only the payload starts a line with one.
 * @param {string} output combined stdout/stderr from the pack command
 * @returns {unknown} the parsed payload
 */
export function parsePackOutput(output) {
	const text = output.trim()
	const objectStart = text.lastIndexOf('\n{')
	const arrayStart = text.lastIndexOf('\n[')
	const start = Math.max(objectStart, arrayStart)
	if (start >= 0) return JSON.parse(text.slice(start + 1))

	const firstObject = text.indexOf('{')
	const firstArray = text.indexOf('[')
	const firstJson = [firstObject, firstArray].filter((idx) => idx >= 0).sort((a, b) => a - b)[0]
	if (firstJson === undefined) throw new Error('pack command did not print JSON output')
	return JSON.parse(text.slice(firstJson))
}

/**
 * Pack the workspace into `packDir` and return the resulting tarball.
 * @param {string} packDir directory to write the tarball into
 * @returns {Promise<{filename: string, tarball: string}>}
 */
export async function packPackage(packDir) {
	await fs.mkdir(packDir, { recursive: true })
	// `pnpm pack` fires `prepack`, which runs a full `pnpm run build`. Every caller
	// already has a current `dist/` (the freshness guard ran first), so that would
	// be a duplicate 3.3s build per pack — and CI packs more than once.
	//
	// The flag spelling is not the obvious one. pnpm 11 rejects `--ignore-scripts`
	// outright ("Unknown option") and honours it only in the `--config.<name>` form.
	// pnpm 12.3.4 accepts both spellings and both skip `prepack`, while a bare
	// `pnpm pack` still rebuilds, so the `--config.` form stays as the one that
	// works on either major. Verified against 11.3.0 and 12.3.4. Re-check on a
	// major pnpm bump.
	await run(process.execPath, [path.join(ROOT, 'scripts', 'ensure-dist.mjs')])
	const result = await run('pnpm', ['pack', '--config.ignore-scripts=true', '--json', '--pack-destination', packDir], {
		capture: true,
	})
	const output = result.stdout || result.stderr
	let entry
	if (output.trim()) {
		const pack = parsePackOutput(output)
		entry = Array.isArray(pack) ? pack[0] : pack
	}
	if (!entry?.filename) {
		const packedFiles = (await fs.readdir(packDir)).filter((file) => file.endsWith('.tgz')).sort()
		if (packedFiles.length !== 1) throw new Error('pnpm pack did not return exactly one tarball in ' + packDir)
		entry = { filename: packedFiles[0] }
	}

	const tarball = path.isAbsolute(entry.filename) ? entry.filename : path.join(packDir, path.basename(entry.filename))
	await assertFile(tarball)
	return { ...entry, filename: path.basename(entry.filename), tarball }
}

/** @param {string} file */
export async function assertFile(file) {
	await fs.access(file)
}

/** @param {string} file */
export async function assertNoFile(file) {
	try {
		await fs.access(file)
	} catch {
		return
	}
	throw new Error('unexpected file exists: ' + file)
}
