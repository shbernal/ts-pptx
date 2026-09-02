#!/usr/bin/env node
// `lefthook install`, but only where lefthook would own the hooks it writes.
//
// `prepare` runs this on every `pnpm install` — and, since pnpm 11 verifies dependencies before
// running any script, that is effectively *every* `pnpm run` and `pnpm exec` too. So whatever
// `prepare` does, it does constantly, and a non-zero exit from it takes the whole package manager
// down with it: `pnpm run lint` never reaches oxlint, it dies in the install that pnpm ran first.
//
// Calling `lefthook install` directly is exactly that hazard. Lefthook refuses to install while
// `core.hooksPath` points somewhere it does not own, and exits 1 — correctly, because writing its
// wrappers into a hooks directory shared by every repo on the machine would clobber whatever else
// lives there. But "correctly refuses" and "fails the build" are different things, and only the
// first one is wanted here. A developer with a machine-wide hooks path could not run a single
// script in this repo.
//
// Lefthook's own two suggestions are both worse than skipping. `--reset-hooks-path` unsets
// `core.hooksPath` globally, silently disarming every other repo that relies on it;
// `--force` writes this repo's wrappers into that shared directory, breaking those repos loudly.
// Neither is a thing a `prepare` script gets to decide on a developer's behalf.
//
// So: install when git would run `.git/hooks`, skip and say so otherwise. Skipping is not silent,
// because a gate nobody knows is off is the failure mode these hooks exist to avoid — hooks under
// a custom path run only if that path delegates back, and this script cannot know whether it does.
//
//   node scripts/install-hooks.mjs
import { spawnSync } from 'node:child_process'
import { createRequire } from 'node:module'
import { homedir } from 'node:os'
import { isAbsolute, join, relative, resolve } from 'node:path'

import { ROOT } from './script-utils.mjs'

/**
 * A `git` query, or undefined when git declines to answer — no repo, or the key is unset.
 * @param {string[]} args
 * @returns {string | undefined}
 */
function git(...args) {
	const result = spawnSync('git', args, { cwd: ROOT, encoding: 'utf8' })
	if (result.status !== 0) return undefined
	const value = result.stdout.trim()
	return value === '' ? undefined : value
}

/**
 * `core.hooksPath` as an absolute path, or undefined when nothing sets it.
 *
 * `git config --get` reports the raw string from whichever file won, so the two conveniences git
 * applies to path-typed values — `~` expansion, and resolving a relative path against the top of
 * the working tree, where hooks are run from — have to be reapplied here.
 * @param {string} topLevel the working tree root, for resolving a relative setting
 * @returns {string | undefined}
 */
function configuredHooksPath(topLevel) {
	const raw = git('config', '--get', 'core.hooksPath')
	if (raw === undefined) return undefined
	const expanded = raw === '~' || raw.startsWith('~/') ? join(homedir(), raw.slice(1)) : raw
	return isAbsolute(expanded) ? resolve(expanded) : resolve(topLevel, expanded)
}

/**
 * Windows reaches the same directory through more than one spelling; git treats them as one.
 * @param {string} a
 * @param {string} b
 */
function samePath(a, b) {
	return process.platform === 'win32' ? a.toLowerCase() === b.toLowerCase() : a === b
}

/**
 * True when this package is being built as somebody's dependency rather than developed here.
 *
 * `npm i github:shbernal/ts-pptx#<sha>` runs `prepare` — this script included — inside a
 * throwaway clone under the package manager's cache, and the devDependencies npm installs in
 * order to run `prepare` put lefthook on disk there. So the "no lefthook, nothing to do" exit
 * in `install()` does *not* cover that case, and what happens instead is at best pointless
 * (writing hooks into a directory npm is about to delete) and at worst fatal: `install()`
 * propagates lefthook's exit status, and a non-zero `prepare` fails the consumer's install.
 *
 * `INIT_CWD` is the directory the package manager was invoked from, and npm, pnpm and yarn
 * all set it. Developing here, that is the repo or something under it (`demos/*`); in a
 * consumer's install it is their project, nowhere near the temp clone this is running from —
 * measured, not assumed: ROOT was `…/npm-cache/_cacache/tmp/git-cloneXXXXXX`. Unset means no
 * package manager ran this, i.e. somebody invoked it by hand, which is a real install.
 */
function installedAsDependency() {
	const initCwd = process.env.INIT_CWD
	if (initCwd === undefined) return false
	const from = process.platform === 'win32' ? ROOT.toLowerCase() : ROOT
	const to = process.platform === 'win32' ? resolve(initCwd).toLowerCase() : resolve(initCwd)
	const rel = relative(from, to)
	return rel !== '' && (rel.startsWith('..') || isAbsolute(rel))
}

/** @param {string[]} args extra arguments for `lefthook install` */
function install(...args) {
	// The package's own entry, not `node_modules/.bin/lefthook` — the shim is a `.CMD` on Windows,
	// which Node will not spawn without a shell, and a shell here would be one more dialect to get
	// wrong. This path is the same file the shim would have run.
	let entry
	try {
		entry = createRequire(import.meta.url).resolve('lefthook/bin/index.js')
	} catch {
		// No lefthook on disk. That is what a `--prod`/`--ignore-scripts` install looks like, and what
		// a git-URL install of this package looks like from the outside: legitimate states in which
		// there are no hooks to install and no reason to fail.
		console.log('hooks: lefthook is not installed — nothing to do.')
		return
	}

	const result = spawnSync(process.execPath, [entry, 'install', ...args], {
		cwd: ROOT,
		stdio: 'inherit',
	})
	process.exitCode = result.status ?? 1
}

if (process.env.LEFTHOOK === '0') {
	console.log('hooks: LEFTHOOK=0 — skipping install.')
} else if (installedAsDependency()) {
	console.log('hooks: installed as a dependency — skipping install.')
} else {
	const gitDir = git('rev-parse', '--absolute-git-dir')
	const topLevel = git('rev-parse', '--show-toplevel')
	if (gitDir === undefined || topLevel === undefined) {
		console.log('hooks: not a git checkout — skipping install.')
	} else {
		const configured = configuredHooksPath(topLevel)
		if (configured === undefined) {
			install()
		} else if (samePath(configured, resolve(gitDir, 'hooks'))) {
			// Lefthook objects to `core.hooksPath` being set at all, not to where it points, so it
			// refuses even this — a path naming the very directory it was going to write to. `--force`
			// is the documented way past that check, and the warning it carries ("installs into the
			// current hooks path") describes the intended destination here, not a shared one.
			install('--force')
		} else {
			console.log(`hooks: core.hooksPath is set to ${configured}, which lefthook does not own.`)
			console.log('       Skipped install. Hooks fire only if that directory delegates back to')
			console.log("       this repo's .git/hooks, or runs `lefthook run <hook>` itself.")
			console.log('       Verify before relying on them: `git commit` should show lefthook output.')
		}
	}
}
