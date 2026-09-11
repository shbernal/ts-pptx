/**
 * ts-pptx: the comments contribution to a written package
 *
 * The deck-wide author registry (`ppt/commentAuthors.xml`) plus one comment part per slide that
 * has comments. A deck with no comments writes neither, so the registry is resolved once here and
 * the whole contribution collapses to nothing.
 */

import type { ZipWriter } from '../../zip.js'
import type { ContentTypeOverride } from '../../gen/opc/content-types.js'
import type { PresentationPropsInternal } from '../../types/internal.js'
import { makeXmlCommentAuthors, makeXmlComments, resolveCommentAuthors } from '../../gen/slide/comments.js'
import { commentPath, overrideName } from '../../gen/opc/part-paths.js'
import type { PartContributor } from './shared.js'
import { COMMENT_AUTHORS_CONTENT_TYPE, COMMENTS_CONTENT_TYPE } from '../../ooxml/rel-types.js'

const COMMENT_AUTHORS_PATH = 'ppt/commentAuthors.xml'

export const commentsContributor: PartContributor = {
	order: 20,
	parts: {
		afterMaster(pres: PresentationPropsInternal, zip: ZipWriter): void {
			// Resolve the deck-wide author registry once, then emit the shared commentAuthors part
			// plus a per-slide comment part for each slide that has comments.
			const resolved = resolveCommentAuthors(pres.slides)
			if (resolved.authors.length === 0) return
			zip.add(COMMENT_AUTHORS_PATH, makeXmlCommentAuthors(resolved.authors))
			pres.slides.forEach((slide, idx) => {
				if ((slide._comments || []).length > 0) {
					zip.add(commentPath(idx + 1), makeXmlComments(slide, resolved.meta))
				}
			})
		},
	},
	contentTypes: {
		trailing(pres: PresentationPropsInternal): ContentTypeOverride[] {
			const entries: ContentTypeOverride[] = []
			pres.slides.forEach((slide, idx) => {
				if ((slide._comments || []).length > 0) {
					entries.push({ partName: overrideName(commentPath(idx + 1)), contentType: COMMENTS_CONTENT_TYPE })
				}
			})
			if (entries.length > 0)
				entries.push({ partName: '/' + COMMENT_AUTHORS_PATH, contentType: COMMENT_AUTHORS_CONTENT_TYPE })
			return entries
		},
	},
}
