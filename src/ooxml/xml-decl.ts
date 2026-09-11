/**
 * The XML prolog every OOXML part this library writes begins with, including the parts the read
 * side regenerates when it serializes a package it has changed.
 *
 * It lived with the generator-only constants, out of the read side's reach, so the read side's
 * serializers spelled it out again four times. A prolog that differs between two parts by one
 * attribute is not something any check reports, which is why it has one home that both sides can
 * import: this module, like the rest of `ooxml/`, imports nothing.
 */
export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
