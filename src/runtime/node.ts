import { Buffer } from 'node:buffer'
import fs from 'node:fs/promises'
import { IMG_SVG_PLACEHOLDER } from '../core-enums.js'
import type { ISlideRelMedia } from '../core-interfaces.js'
import type { RuntimeAdapter } from './types.js'

export function createNodeRuntime(): RuntimeAdapter {
	return {
		writeFileOutputType: 'nodebuffer',
		loadMedia,
		createSvgPngPreview,
		writeFile,
	}
}

async function loadMedia(rel: ISlideRelMedia & { path: string }): Promise<string> {
	if (rel.path.startsWith('http')) {
		const response = await fetch(rel.path)
		if (!response.ok) throw new Error(`ERROR! Unable to load image (fetch): ${rel.path}`)
		return Buffer.from(await response.arrayBuffer()).toString('base64')
	}

	try {
		return Buffer.from(await fs.readFile(rel.path)).toString('base64')
	} catch (ex) {
		throw new Error(`ERROR: Unable to read media: "${rel.path}"\n${String(ex)}`, { cause: ex })
	}
}

async function createSvgPngPreview(rel: ISlideRelMedia): Promise<string> {
	rel.data = IMG_SVG_PLACEHOLDER
	return 'done'
}

async function writeFile(fileName: string, data: string | ArrayBuffer | Blob | Uint8Array): Promise<string> {
	await fs.writeFile(fileName, data as string | Uint8Array)
	return fileName
}
