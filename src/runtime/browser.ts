import { IMG_SVG_PLACEHOLDER } from '../core-enums.js'
import type { ISlideRelMedia } from '../core-interfaces.js'
import type { RuntimeAdapter } from './types.js'

export function createBrowserRuntime(): RuntimeAdapter {
	return {
		writeFileOutputType: null,
		loadMedia,
		createSvgPngPreview,
		writeFile,
		loadFontData,
	}
}

async function loadFontData(source: string): Promise<Uint8Array> {
	const response = await fetch(source)
	if (!response.ok) throw new Error(`ERROR! Unable to load font (fetch): ${source}`)
	return new Uint8Array(await response.arrayBuffer())
}

async function loadMedia(rel: ISlideRelMedia & { path: string }): Promise<string> {
	const response = await fetch(rel.path)
	if (!response.ok) throw new Error(`ERROR! Unable to load image (fetch): ${rel.path}`)
	const blob = await response.blob()

	return await new Promise<string>((resolve, reject) => {
		const reader = new FileReader()
		reader.onloadend = () => resolve(reader.result as string)
		reader.onerror = () => reject(new Error(`ERROR! Unable to load image (FileReader): ${rel.path}`))
		reader.readAsDataURL(blob)
	})
}

async function createSvgPngPreview(rel: ISlideRelMedia): Promise<string> {
	return await new Promise((resolve, reject) => {
		const image = new Image()
		const fail = (reason?: unknown) => {
			rel.data = IMG_SVG_PLACEHOLDER
			reject(new Error(`ERROR! Unable to load image (image.onerror): ${rel.path}${reason ? ` - ${String(reason)}` : ''}`))
		}

		image.onload = () => {
			if (image.width + image.height === 0) {
				fail('h/w=0')
				return
			}
			const canvas = document.createElement('CANVAS') as HTMLCanvasElement
			const ctx = canvas.getContext('2d')
			if (!ctx) {
				fail('canvas 2d context unavailable')
				return
			}
			canvas.width = image.width
			canvas.height = image.height
			ctx.drawImage(image, 0, 0)
			try {
				rel.data = canvas.toDataURL(rel.type)
				resolve('done')
			} catch (ex) {
				fail(ex)
			}
		}
		image.onerror = () => fail()
		image.src = typeof rel.data === 'string' ? rel.data : IMG_SVG_PLACEHOLDER
	})
}

async function writeFile(fileName: string, data: string | ArrayBuffer | Blob | Uint8Array): Promise<string> {
	const eleLink = document.createElement('a')
	eleLink.setAttribute('style', 'display:none;')
	eleLink.dataset.interception = 'off'
	document.body.appendChild(eleLink)

	const url = window.URL.createObjectURL(new Blob([data as Blob], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }))
	eleLink.href = url
	eleLink.download = fileName
	eleLink.click()

	setTimeout(() => {
		window.URL.revokeObjectURL(url)
		document.body.removeChild(eleLink)
	}, 100)

	return fileName
}
