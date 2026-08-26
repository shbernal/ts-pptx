import type { WRITE_OUTPUT_TYPE } from '../enums.js'
import type { SlideRelMedia } from '../types/internal.js'

export type RuntimeAdapter = {
	readonly writeFileOutputType: WRITE_OUTPUT_TYPE | null
	/**
	 * Load a media file into a base64 payload for `rel.data`.
	 *
	 * **Two encodings are accepted, deliberately.** Either raw base64 (`iVBORw0…`) or a full
	 * data URI (`data:image/png;base64,iVBORw0…`). The browser adapter returns the second
	 * because `FileReader.readAsDataURL` is what decodes a blob without blocking; Node and the
	 * neutral adapter return the first because neither has a `FileReader`. So an adapter may
	 * return whichever its host makes cheap: `gen/media.ts` puts the result through
	 * {@link toMediaDataUri} on the way into `rel.data`, and everything past that point sees one
	 * form. Do not add a third: that helper recognizes exactly these two.
	 */
	loadMedia: (rel: SlideRelMedia & { path: string }) => Promise<string>
	createSvgPngPreview: (rel: SlideRelMedia) => Promise<string>
	writeFile: (fileName: string, data: string | ArrayBuffer | Blob | Uint8Array) => Promise<string>
	/**
	 * Load a font file into raw bytes for `registerFontMetrics`.
	 *
	 * An `http:`/`https:` URL is fetched by every adapter. Anything else is a filesystem path,
	 * which only the Node adapter can read; elsewhere it is a load failure naming the path.
	 */
	loadFontData: (source: string) => Promise<Uint8Array>
}
