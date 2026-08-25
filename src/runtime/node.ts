import { Buffer } from 'node:buffer'
import fs from 'node:fs/promises'
import { MediaError } from '../errors.js'
import type { SlideRelMedia } from '../types/internal.js'
import type { RuntimeAdapter } from './types.js'
import { fetchFontBytes, fetchMediaBase64, isRemote, placeholderSvgPreview } from './fetch-media.js'

export function createNodeRuntime(): RuntimeAdapter {
	return {
		writeFileOutputType: 'nodebuffer',
		loadMedia,
		createSvgPngPreview: placeholderSvgPreview,
		writeFile,
		loadFontData,
	}
}

async function loadFontData(source: string): Promise<Uint8Array> {
	if (isRemote(source)) return await fetchFontBytes(source)
	try {
		return new Uint8Array(await fs.readFile(source))
	} catch (ex) {
		throw new MediaError('font/read-failed', `Unable to read font file: "${source}"\n${String(ex)}`, {
			cause: ex,
		})
	}
}

async function loadMedia(rel: SlideRelMedia & { path: string }): Promise<string> {
	if (isRemote(rel.path)) return await fetchMediaBase64(rel.path)

	try {
		return Buffer.from(await fs.readFile(rel.path)).toString('base64')
	} catch (ex) {
		throw new MediaError('media/read-failed', `Unable to read media: "${rel.path}"\n${String(ex)}`, {
			cause: ex,
		})
	}
}

async function writeFile(fileName: string, data: string | ArrayBuffer | Blob | Uint8Array): Promise<string> {
	await fs.writeFile(fileName, data as string | Uint8Array)
	return fileName
}
