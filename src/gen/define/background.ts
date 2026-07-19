/**
 * PptxGenJS: Background Definition
 *
 * Registers an `addBackground()` image as a slide media rel (color backgrounds carry no rel);
 * the `<p:bg>` XML is emitted later at slide / layout serialize time.
 */
import type { BackgroundProps, SlideLayoutInternal } from '../../core-interfaces.js'
import { imageContentType } from '../../gen-utils.js'

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
		// NOTE: `Target` cannot have spaces (eg:"Slide 1-image-1.jpg") or a "presentation is corrupt" warning comes up
		target._relsMedia.push({
			path: props.path,
			type: imageContentType(strImgExtn),
			extn: strImgExtn,
			data: props.data || undefined,
			rId: intRels,
			Target: `../media/${(target._name || '').replace(/\s+/gi, '-')}-image-${target._relsMedia.length + 1}.${strImgExtn}`,
		})
		target._bkgdImgRid = intRels
	}
}
