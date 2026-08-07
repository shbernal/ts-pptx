/**
 * ts-pptx: OLE / Embedded Object Definition (`addOleObject()` — PowerPoint's Insert ▸ Object).
 *
 * Registers the payload as an embedded package part plus a preview-picture image rel, and pushes a
 * `SlideObject{ _type: oleObject, ole }` for `gen/slide/objects/ole.ts` to emit as a `<p:graphicFrame>`.
 *
 * Ground truth for the two package-level choices here (verified against a deck authored by
 * PowerPoint via `Shapes.AddOLEObject`):
 *   - an embedded Office file (xlsx/docx/pptx — each an OPC package in its own right) is referenced
 *     with the `.../relationships/package` rel type and gets a content-type `Default` for its
 *     extension; only a non-package OLE server blob uses `.../relationships/oleObject` with a `.bin`
 *     part (ECMA-376 Part 1 §15.2.10);
 *   - the payload lives in `ppt/embeddings/`, the preview picture in `ppt/media/` as an ordinary image.
 */
import { SlideObjectType } from '../../enums.js'
import type { OleObjectProps } from '../../types/media.js'
import type { PresSlideInternal, SlideObject } from '../../types/internal.js'
import { encodeXmlAttrValue, getNewRelId, validateObjectName } from '../utils.js'
import { nextObjectNameIdx } from './object-name.js'
import { registerPreviewImage } from './preview-image.js'
import { InvalidOptionError } from '../../errors.js'
import { OFFICE_REL } from '../oxml/schema-uris.js'

const OD = 'application/vnd.openxmlformats-officedocument.'
/** Rel type for an embedded OPC package (an Office file, itself a zip). */
const PACKAGE_REL = OFFICE_REL + 'package'
/** Rel type for a generic OLE-server blob — a compound-file `.bin` (ECMA-376 Part 1 §15.2.10). */
const OLE_OBJECT_REL = OFFICE_REL + 'oleObject'

interface OleFormat {
	/** OPC content type, emitted as a `Default` entry for the part's extension. */
	contentType: string
	/** `.rels` `Type` URI for the payload part. */
	relType: string
	/** Default `p:oleObj@progId`. */
	progId: string
	/** Default `p:oleObj@name` — how PowerPoint labels the object's kind. */
	name: string
}

/**
 * Embeddable payload formats, keyed by file extension. The six Office extensions are OPC packages
 * and keep their own extension in `ppt/embeddings/`; everything else is embedded as a generic OLE
 * blob under {@link BIN_FORMAT}.
 */
const OLE_FORMATS: Record<string, OleFormat> = {
	xlsx: { contentType: OD + 'spreadsheetml.sheet', relType: PACKAGE_REL, progId: 'Excel.Sheet.12', name: 'Worksheet' },
	xlsm: {
		contentType: 'application/vnd.ms-excel.sheet.macroEnabled.12',
		relType: PACKAGE_REL,
		progId: 'Excel.SheetMacroEnabled.12',
		name: 'Worksheet',
	},
	docx: {
		contentType: OD + 'wordprocessingml.document',
		relType: PACKAGE_REL,
		progId: 'Word.Document.12',
		name: 'Document',
	},
	docm: {
		contentType: 'application/vnd.ms-word.document.macroEnabled.12',
		relType: PACKAGE_REL,
		progId: 'Word.DocumentMacroEnabled.12',
		name: 'Document',
	},
	pptx: {
		contentType: OD + 'presentationml.presentation',
		relType: PACKAGE_REL,
		progId: 'PowerPoint.Show.12',
		name: 'Presentation',
	},
	pptm: {
		contentType: 'application/vnd.ms-powerpoint.presentation.macroEnabled.12',
		relType: PACKAGE_REL,
		progId: 'PowerPoint.ShowMacroEnabled.12',
		name: 'Presentation',
	},
}
/** Fallback for any payload that is not one of the {@link OLE_FORMATS} Office packages. */
const BIN_FORMAT: OleFormat = {
	contentType: OD + 'oleObject',
	relType: OLE_OBJECT_REL,
	progId: 'Package',
	name: 'Object',
}

/** Reverse of {@link OLE_FORMATS}: a caller-supplied `progId` implies the payload's extension. */
const EXTN_BY_PROG_ID = new Map(Object.entries(OLE_FORMATS).map(([extn, fmt]) => [fmt.progId.toLowerCase(), extn]))
/** Reverse of {@link OLE_FORMATS}: a `data:` URI's MIME implies the payload's extension. */
const EXTN_BY_CONTENT_TYPE = new Map(
	Object.entries(OLE_FORMATS).map(([extn, fmt]) => [fmt.contentType.toLowerCase(), extn])
)

/**
 * Resolve the embedded part's extension from whatever the caller gave us, in decreasing order of
 * explicitness: `extn`, the `data:` URI's MIME, the `path`'s extension, then the `progId`. Anything
 * unrecognized becomes `bin`, the generic OLE blob — which is what PowerPoint itself writes for a
 * shell-packaged payload, so an unknown extension never leaks into `[Content_Types].xml`.
 */
function resolvePartExtn(opt: OleObjectProps): string {
	const explicit = (opt.extn || '').replace(/^\./, '').toLowerCase()
	if (explicit) return explicit in OLE_FORMATS ? explicit : 'bin'

	const mime = /^data:([^;,]+)[;,]/.exec(opt.data || '')?.[1]?.toLowerCase()
	const byMime = mime ? EXTN_BY_CONTENT_TYPE.get(mime) : undefined
	if (byMime) return byMime

	const file = (opt.path || '').split(/[?#]/)[0] ?? ''
	const byPath = file.includes('.') ? (file.split('.').pop() ?? '').toLowerCase() : ''
	if (byPath in OLE_FORMATS) return byPath

	return EXTN_BY_PROG_ID.get((opt.progId || '').toLowerCase()) ?? 'bin'
}

/**
 * Adds an embedded OLE object to a slide definition.
 * @param {PresSlideInternal} target - slide the object will be added to
 * @param {OleObjectProps} opt - OLE object options
 */
export function addOleObjectDefinition(target: PresSlideInternal, opt: OleObjectProps): void {
	const strData = opt.data || ''
	const strPath = opt.path || ''

	// STEP 1: REALITY-CHECK. The payload is the whole point; there is no meaningful default.
	if (!strPath && !strData) {
		throw new InvalidOptionError('ole/missing-source', 'addOleObject(): either `data` or `path` are required!')
	}

	// STEP 2: Resolve the payload format — part extension, content type, rel type, progId.
	const extn = resolvePartExtn(opt)
	const format = OLE_FORMATS[extn] ?? BIN_FORMAT
	const nameIdx = nextObjectNameIdx(target, SlideObjectType.oleObject)
	const objectName = opt.objectName
		? encodeXmlAttrValue(validateObjectName(opt.objectName, 'oleObject'))
		: `Object ${nameIdx + 1}`

	// STEP 3: Register the payload part rel. Deliberately NOT deduped against an identical payload
	// elsewhere in the deck: PowerPoint gives every OLE object its own embedding part, and sharing
	// one would make editing either object rewrite the other's source.
	const objectRid = getNewRelId(target)
	const mediaSlideKey =
		target._slideNum == null ? 'sm' : target._slideNum >= 1000 ? `sl-${target._slideNum}` : target._slideNum
	target._relsMedia.push({
		path: strPath || `preencoded.${extn}`,
		type: format.contentType,
		extn,
		data: strData,
		rId: objectRid,
		oleRelType: format.relType,
		Target: `../embeddings/oleObject-${mediaSlideKey}-${target._relsMedia.length + 1}.${extn}`,
	})

	// STEP 4: Register the preview picture (gray placeholder when the caller supplied no cover).
	const previewRid = registerPreviewImage(target, opt.cover)

	// LAST: Push the slide object for the `<p:graphicFrame>` emitter.
	const slideData: SlideObject = {
		_type: SlideObjectType.oleObject,
		options: {
			x: opt.x ?? 0,
			y: opt.y ?? 0,
			// The library never opens the payload, so there is no natural size to measure.
			w: opt.w ?? 4,
			h: opt.h ?? 3,
			objectName,
			...(opt.altText ? { altText: opt.altText } : {}),
			...(opt.objectLock ? { objectLock: opt.objectLock } : {}),
		},
		ole: {
			objectRid,
			previewRid,
			progId: opt.progId || format.progId,
			name: format.name,
			showAsIcon: !!opt.showAsIcon,
			...(typeof opt.imgW === 'number' ? { imgW: opt.imgW } : {}),
			...(typeof opt.imgH === 'number' ? { imgH: opt.imgH } : {}),
		},
	}
	target._slideObjects.push(slideData)
}
