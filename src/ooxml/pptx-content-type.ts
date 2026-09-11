/**
 * The content type of a `.pptx` file itself: what the `zip` entry and the browser download label
 * their Blob with, and what an OLE object declares for an embedded presentation.
 *
 * It is one of the content types `rel-types.ts` holds, and that module re-exports it. It has a
 * module of its own because the `zip` entry needs this one string and nothing else from that
 * table: imported from there, it handed the entry the whole table's chunk, about as many bytes
 * again as the entry itself, and pushed the entry past its size budget. Like the rest of `ooxml/`,
 * this module imports nothing.
 */
export const PPTX_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
