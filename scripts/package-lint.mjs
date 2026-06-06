#!/usr/bin/env node
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { packPackage, run } from './script-utils.mjs'

const tmp = await fs.mkdtemp(path.join(os.tmpdir(), 'pptxgenjs-package-lint-'))
const keepTmp = process.env.PPTXGENJS_KEEP_PACKAGE_LINT === '1'

try {
	const packInfo = await packPackage(path.join(tmp, 'pack'))
	await run('publint', ['run', packInfo.tarball, '--pack', 'false'])
	await run('attw', [packInfo.tarball, '--profile', 'esm-only'])
	console.log('Package lint passed: ' + packInfo.filename)
} finally {
	if (keepTmp) console.log('Keeping package lint temp directory: ' + tmp)
	else await fs.rm(tmp, { recursive: true, force: true })
}
