#!/usr/bin/env node
import { withPackedTarball } from './pack-utils.mjs'
import { run } from './script-utils.mjs'

await withPackedTarball('package-lint', async (packInfo) => {
	await run('publint', ['run', packInfo.tarball, '--pack', 'false'])
	await run('attw', [packInfo.tarball, '--profile', 'esm-only'])
	console.log('Package lint passed: ' + packInfo.filename)
})
