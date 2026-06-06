#!/usr/bin/env node
import { spawn } from 'node:child_process'
import os from 'node:os'
import path from 'node:path'

const packageManagerCache = path.join(os.tmpdir(), 'pptxgenjs-package-manager-cache')
const child = spawn('npm', ['pack', '--dry-run', '--json'], {
	env: {
		...process.env,
		npm_config_cache: path.join(packageManagerCache, 'npm'),
		NPM_CONFIG_CACHE: path.join(packageManagerCache, 'npm'),
	},
	stdio: 'inherit',
})

child.on('error', (error) => {
	console.error(error)
	process.exitCode = 1
})

child.on('close', (code) => {
	process.exitCode = code ?? 1
})
