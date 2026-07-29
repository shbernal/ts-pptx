import type { WRITE_OUTPUT_TYPE } from '../enums.js'
import type { SlideRelMedia } from '../types/internal.js'

export type RuntimeAdapter = {
	readonly writeFileOutputType: WRITE_OUTPUT_TYPE | null
	loadMedia: (rel: SlideRelMedia & { path: string }) => Promise<string>
	createSvgPngPreview: (rel: SlideRelMedia) => Promise<string>
	writeFile: (fileName: string, data: string | ArrayBuffer | Blob | Uint8Array) => Promise<string>
	/** Load a font file (path on Node, URL anywhere) into raw bytes for `registerFontMetrics`. */
	loadFontData: (source: string) => Promise<Uint8Array>
}
