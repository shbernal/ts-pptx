/**
 * PptxGenJS: Slide-Master Definition
 *
 * `createSlideMaster` walks a `SlideMasterProps` definition onto a layout target: shared
 * chart / image / shape / text children via `addChildDefinition`, plus master-specific text
 * placeholders (which need the object index for `_placeholderIdx`).
 */
import type {
	ObjectOptions,
	PresSlideInternal,
	SlideLayoutInternal,
	SlideMasterProps,
	TextPropsOptions,
} from '../../core-interfaces.js'
import { addChildDefinition } from './group.js'
import { addTextDefinition } from './text.js'

/**
 * Transforms a slide definition to a slide object that is then passed to the XML transformation process.
 * @param {SlideMasterProps} props - slide definition
 * @param {PresSlideInternal|SlideLayoutInternal} target - empty slide object that should be updated by the passed definition
 */
export function createSlideMaster(props: SlideMasterProps, target: SlideLayoutInternal): void {
	// STEP 1: Add all Slide Master objects in the order they were given
	if (props.objects && Array.isArray(props.objects) && props.objects.length > 0) {
		props.objects.forEach((object, idx) => {
			const tgt = target as PresSlideInternal
			if (addChildDefinition(tgt, object)) {
				// handled by the shared chart/image/shape/text dispatch
			} else if ('placeholder' in object) {
				const placeholder = object.placeholder
				const { name, type, ...rawPlaceholderOptions } = placeholder.options
				const placeholderOptions = rawPlaceholderOptions as TextPropsOptions & ObjectOptions
				placeholderOptions.placeholder = name
				placeholderOptions._placeholderType = type
				placeholderOptions._placeholderIdx = 100 + idx
				addTextDefinition(tgt, [{ text: placeholder.text }], placeholderOptions, true)
				// NOTE: only text placeholders are supported (image/other placeholder kinds are not emitted)
			}
		})
	}

	// STEP 2: Add Slide Numbers (NOTE: Do this last so numbers are not covered by objects!)
	if (props.slideNumber && typeof props.slideNumber === 'object') target._slideNumberProps = props.slideNumber
}
