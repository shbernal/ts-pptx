/**
 * PptxGenJS: Media Methods
 */

import { IMG_BROKEN } from './core-enums.js'
import type { PresSlideInternal, SlideLayoutInternal, ISlideRelMedia } from './core-interfaces.js'

type SlideMediaRelWithPath = ISlideRelMedia & { path: string }

function hasEncodingPath(rel: ISlideRelMedia): rel is SlideMediaRelWithPath {
	return typeof rel.path === 'string' && rel.path.length > 0 && !rel.path.includes('preencoded')
}

/**
 * Encode Image/Audio/Video into base64
 * @param {PresSlideInternal | SlideLayoutInternal} layout - slide layout
 * @return {Promise} promise
 */
export function encodeSlideMediaRels(layout: PresSlideInternal | SlideLayoutInternal): Array<Promise<string>> {
	// STEP 1: Detect real Node runtime once
	const isNode = typeof process !== 'undefined' && !!process.versions?.node && process.release?.name === 'node'
	// These will be filled only when we’re in Node
	let fs: typeof import('node:fs') | undefined
	let https: typeof import('node:https') | undefined

	// STEP 2: Lazy-load Node built-ins if needed
	const loadNodeDeps = isNode
		? async () => {
			; ({ default: fs } = await import('node:fs')); ({ default: https } = await import('node:https'))
		}
		: async () => { }
	// Immediately start it when we know we’re in Node
	if (isNode) loadNodeDeps()

	// STEP 3: Prepare promises list
	const imageProms: Array<Promise<string>> = []

	// A: Capture all audio/image/video candidates for encoding (filtering online/pre-encoded)
	const candidateRels = layout._relsMedia.filter(
		(rel): rel is SlideMediaRelWithPath => rel.type !== 'online' && !rel.data && hasEncodingPath(rel)
	)

	// B: PERF: Mark dupes (same `path`) to avoid loading the same media over-and-over!
	const unqPaths: string[] = []
	candidateRels.forEach(rel => {
		if (!unqPaths.includes(rel.path)) {
			rel.isDuplicate = false
			unqPaths.push(rel.path)
		} else {
			rel.isDuplicate = true
		}
	})

	// STEP 4: Read/Encode each unique media item
	candidateRels
		.filter(rel => !rel.isDuplicate)
		.forEach(rel => {
			imageProms.push(
				(async () => {
					if (!https) await loadNodeDeps()

					// ────────────  NODE LOCAL FILE  ────────────
					if (isNode && fs && rel.path.indexOf('http') !== 0) {
						try {
							const bitmap = fs.readFileSync(rel.path)
							rel.data = Buffer.from(bitmap).toString('base64')
							candidateRels
								.filter(dupe => dupe.isDuplicate && dupe.path === rel.path)
								.forEach(dupe => (dupe.data = rel.data))
							return 'done'
						} catch (ex) {
							rel.data = IMG_BROKEN
							candidateRels
								.filter(dupe => dupe.isDuplicate && dupe.path === rel.path)
								.forEach(dupe => (dupe.data = rel.data))
							throw new Error(`ERROR: Unable to read media: "${rel.path}"\n${String(ex)}`, { cause: ex })
						}
					}

					// ────────────  NODE HTTP(S)  ────────────
					if (isNode && https && rel.path.startsWith('http')) {
						const httpsModule = https
						return await new Promise<string>((resolve, reject) => {
							httpsModule.get(rel.path, res => {
								let raw = ''
								res.setEncoding('binary') // IMPORTANT: Only binary encoding works
								res.on('data', chunk => (raw += chunk))
								res.on('end', () => {
									rel.data = Buffer.from(raw, 'binary').toString('base64')
									candidateRels
										.filter(dupe => dupe.isDuplicate && dupe.path === rel.path)
										.forEach(dupe => (dupe.data = rel.data))
									resolve('done')
								})
								res.on('error', () => {
									rel.data = IMG_BROKEN
									candidateRels
										.filter(dupe => dupe.isDuplicate && dupe.path === rel.path)
										.forEach(dupe => (dupe.data = rel.data))
									reject(new Error(`ERROR! Unable to load image (https.get): ${rel.path}`))
								})
							})
						})
					}

					// ────────────  BROWSER  ────────────
					return await new Promise<string>((resolve, reject) => {
						// A: build request
						const xhr = new XMLHttpRequest()
						xhr.onload = () => {
							const reader = new FileReader()
							reader.onloadend = () => {
								rel.data = reader.result as string
								candidateRels
									.filter(dupe => dupe.isDuplicate && dupe.path === rel.path)
									.forEach(dupe => (dupe.data = rel.data))
								if (!rel.isSvgPng) {
									resolve('done')
								} else {
									createSvgPngPreview(rel)
										.then(() => resolve('done'))
										.catch(reject)
								}
							}
							reader.readAsDataURL(xhr.response)
						}
						xhr.onerror = () => {
							rel.data = IMG_BROKEN
							candidateRels
								.filter(dupe => dupe.isDuplicate && dupe.path === rel.path)
								.forEach(dupe => (dupe.data = rel.data))
							reject(new Error(`ERROR! Unable to load image (xhr.onerror): ${rel.path}`))
						}
						// B: execute request
						xhr.open('GET', rel.path)
						xhr.responseType = 'blob'
						xhr.send()
					})
				})(),
			)
		})

	// STEP 5: SVG-PNG previews
	// ......: "SVG:" base64 data still requires a png to be generated
	// ......: (`isSvgPng` flag this as the preview image, not the SVG itself)
	layout._relsMedia
		.filter(rel => rel.isSvgPng && rel.data)
		.forEach(rel => {
			(async () => {
				if (isNode && !fs) await loadNodeDeps()
				if (isNode && fs) {
					// console.log('Sorry, SVG is not supported in Node (more info: https://github.com/gitbrent/PptxGenJS/issues/401)')
					rel.data = IMG_BROKEN
					imageProms.push(Promise.resolve('done'))
				} else {
					imageProms.push(createSvgPngPreview(rel))
				}
			})()
		})

	return imageProms
}

/**
 * Create SVG preview image
 * @param {ISlideRelMedia} rel - slide rel
 * @return {Promise} promise
 */
async function createSvgPngPreview(rel: ISlideRelMedia): Promise<string> {
	return await new Promise((resolve, reject) => {
		// A: Create
		const image = new Image()
		const fail = (reason?: unknown) => {
			rel.data = IMG_BROKEN
			reject(new Error(`ERROR! Unable to load image (image.onerror): ${rel.path}${reason ? ` - ${String(reason)}` : ''}`))
		}

		// B: Set onload event
		image.onload = () => {
			// First: Check for any errors: This is the best method (try/catch wont work, etc.)
			if (image.width + image.height === 0) {
				fail('h/w=0')
				return
			}
			const canvas: HTMLCanvasElement = document.createElement('CANVAS') as HTMLCanvasElement
			const ctx = canvas.getContext('2d')
			if (!ctx) {
				fail('canvas 2d context unavailable')
				return
			}
			canvas.width = image.width
			canvas.height = image.height
			ctx.drawImage(image, 0, 0)
			// Users running on local machine will get the following error:
			// "SecurityError: Failed to execute 'toDataURL' on 'HTMLCanvasElement': Tainted canvases may not be exported."
			// when the canvas.toDataURL call executes below.
			try {
				rel.data = canvas.toDataURL(rel.type)
				resolve('done')
			} catch (ex) {
				fail(ex)
			}
		}
		image.onerror = () => fail()

		// C: Load image
		image.src = typeof rel.data === 'string' ? rel.data : IMG_BROKEN
	})
}
