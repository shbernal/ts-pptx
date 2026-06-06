import type { WRITE_OUTPUT_TYPE } from '../core-enums.js'
import type { ISlideRelMedia } from '../core-interfaces.js'

export type RuntimeAdapter = {
	readonly writeFileOutputType: WRITE_OUTPUT_TYPE | null
	loadMedia: (rel: ISlideRelMedia & { path: string }) => Promise<string>
	createSvgPngPreview: (rel: ISlideRelMedia) => Promise<string>
	writeFile: (fileName: string, data: string | ArrayBuffer | Blob | Uint8Array) => Promise<string>
}
