/**
 * PptxGenJS: Background Definition
 *
 * Registers an `addBackground()` image as a slide media rel (color backgrounds carry no rel);
 * the `<p:bg>` XML is emitted later at slide / layout serialize time.
 */
import type { BackgroundProps } from '../../core-interfaces.js'
import type { SlideLayoutInternal } from '../../types/internal.js'
import { imageContentType } from '../../media/content-type.js'

/**
 * Reduce a slide/layout name to something safe to embed in a media part name.
 *
 * A background rel's `Target` is used twice over: it is written into the `.rels` part *and*,
 * with `..` swapped for `ppt`, used verbatim as the ZIP entry name (`pptxgen.ts`). A layout's
 * name is caller-supplied (`defineSlideMaster({ title })`), so without this the caller can put
 * arbitrary characters into an OPC part name. XML-escaping does not help — the escaping is
 * undone before the target is resolved, and the ZIP entry is never escaped at all. Demonstrated
 * breakage before this was added:
 *   - `%` produced an invalid percent-escape, so the target would not decode at all
 *   - `?` and `#` began a query/fragment, truncating the resolved path (`what?now-image-1.png`
 *     resolves to `ppt/media/what`, which is not in the package)
 *   - `/` silently pushed the media into a subdirectory
 *
 * So the safe set is the URI "unreserved" characters. Runs of anything else collapse to a
 * single `-` (subsuming the whitespace handling this replaces), leading/trailing punctuation is
 * trimmed, and a name left empty — including an all-non-ASCII one — falls back to `media`.
 * The result is cosmetic: it names the media part, and nothing resolves a layout by it.
 */
export function sanitizeMediaNamePart(name: string): string {
	const safe = name
		.replace(/[^A-Za-z0-9._-]+/g, '-')
		.replace(/^[-.]+/, '')
		.replace(/[-.]+$/, '')
	return safe || 'media'
}

/**
 * Adds a background image or color to a slide definition.
 * @param {BackgroundProps} props - color string or an object with image definition
 * @param {PresSlideInternal} target - slide object that the background is set to
 */
export function addBackgroundDefinition(props: BackgroundProps | undefined, target: SlideLayoutInternal): void {
	// Handle media
	if (props && (props.path || props.data)) {
		// Allow the use of only the data key (`path` isnt reqd)
		props.path = props.path || 'preencoded.png'
		let strImgExtn = (props.path.split('.').pop() || 'png').split('?')[0] ?? 'png' // Handle "blah.jpg?width=540" etc.
		if (strImgExtn === 'jpg') strImgExtn = 'jpeg' // base64-encoded jpg's come out as "data:image/jpeg;base64,/9j/[...]", so correct exttnesion to avoid content warnings at PPT startup

		target._relsMedia = target._relsMedia || []
		const intRels = target._relsMedia.length + 1
		// NOTE: `Target` cannot have spaces (eg:"Slide 1-image-1.jpg") or a "presentation is corrupt"
		// warning comes up — `sanitizeMediaNamePart` covers that case along with the rest.
		target._relsMedia.push({
			path: props.path,
			type: imageContentType(strImgExtn),
			extn: strImgExtn,
			data: props.data || undefined,
			rId: intRels,
			Target: `../media/${sanitizeMediaNamePart(target._name || '')}-image-${target._relsMedia.length + 1}.${strImgExtn}`,
		})
		target._bkgdImgRid = intRels
	}
}
