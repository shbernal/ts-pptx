/**
 * `pptxgenjs/read` — open an existing `.pptx`, inspect its OPC structure, and
 * save it back with untouched parts byte-identical (lossless round-trip).
 *
 * This subsystem is isomorphic: bytes in, bytes out, no `node:fs`. File I/O
 * is the caller's job.
 */
export { OpcPackage, type OpcInput } from './read/opc/package.js'
export { Part } from './read/opc/part.js'
export { ContentTypes } from './read/opc/content-types.js'
export { Relationships, type Relationship } from './read/opc/relationships.js'
export { resolveRelativePartName, relsPartNameFor } from './read/opc/partnames.js'
