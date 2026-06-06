# OOXML Agent Context

This repository has project-scoped Codex MCP configuration in `.codex/config.toml`.
Codex loads it only when the project is trusted. In a running Codex session, use
`/mcp` to confirm the active servers.

## Configured MCP Servers

### `ooxml`

Endpoint: `https://api.ooxml.dev/mcp`

Purpose: ECMA-376 / Office Open XML reference lookup. This is a third-party MCP
server, not an ISO, Ecma, or Microsoft service. Treat it as a fast retrieval and
schema navigation layer, then verify behavior with generated fixtures and the
OOXML validator.

Available tool families:

- Prose search over ECMA-376: `ooxml_search`, `ooxml_section`, `ooxml_parts`.
- Deterministic schema lookup: `ooxml_element`, `ooxml_type`, `ooxml_children`,
  `ooxml_attributes`, `ooxml_enum`, `ooxml_namespace`.
- OPC package metadata: `ooxml_package_part`.

Use it when you need to answer questions such as:

- Which children are legal under a PresentationML or DrawingML element?
- Which attributes and enum values are valid?
- Which package content type or relationship type belongs to a `.pptx` part?
- Which ECMA-376 section describes a serialization rule?

### `microsoft_learn`

Endpoint: `https://learn.microsoft.com/api/mcp`

Purpose: official Microsoft documentation retrieval. Use this for Microsoft Open
Specifications and implementation behavior, especially when PowerPoint accepts,
repairs, extends, or rejects OOXML in ways that are not obvious from the schema.

Available tools:

- `microsoft_docs_search`: search Microsoft Learn and official Microsoft docs.
- `microsoft_docs_fetch`: fetch a full Microsoft documentation page as markdown.
- `microsoft_code_sample_search`: search official Microsoft code samples.

Use it when you need to answer questions such as:

- What does `[MS-PPTX]` say about a PowerPoint extension element?
- What does `[MS-OI29500]` say about Office implementation behavior?
- How does the Open XML SDK validate or model a package part?
- Is a namespace, extension URI, or relationship type Microsoft-specific?

## Retrieval Workflow

1. Start with local evidence. Search `src/`, `test/`, `tools/ooxml-validator/`,
   `README.md`, and `docs/testing.md` before changing behavior.
2. Use `ooxml` for normative structure: schema order, child elements,
   attributes, simple type enums, namespaces, and OPC package metadata.
3. Use `microsoft_learn` for PowerPoint-specific behavior, Microsoft extension
   namespaces, Open XML SDK behavior, and Office compatibility notes.
4. If the two sources disagree, document the difference in the code comment or
   test name only when it affects the implementation. Prefer PowerPoint and
   Open XML SDK behavior for this library's generated `.pptx` compatibility.
5. Validate with a minimal generated fixture. For serialization changes, add or
   update a focused case in `test/schema.test.js` and run
   `pnpm run test:schema`.

## What Not To Do

- Do not commit full OOXML standards PDFs, large copied spec excerpts, or bulk
  extracted standard text.
- Do not rely on schema validity alone for user-visible compatibility. Some
  PowerPoint "needs repair" failures are implementation constraints caught by
  Open XML SDK validation or by opening the deck in PowerPoint.
- Do not introduce ad hoc XML ordering rules without either a local fixture, a
  PowerPoint-authored comparison, or a referenced standard/source note.
- Do not treat Microsoft extension namespaces as ECMA-defined without checking
  the Microsoft documentation.

## Local Validation Tools

- Install the validator once with `./tools/ooxml-validator/install.sh`.
- Run OOXML fixtures with `pnpm run test:schema`.
- Run the normal regression suite with `pnpm run test:unit`.
- Run the full default test command with `pnpm test` when both regression and
  schema validation are relevant.

Useful local files:

- `tools/ooxml-validator/README.md`
- `test/schema.test.js`
- `test/schema-validation.test.mjs`
- `test/validator.js`
- `src/gen-xml.ts`
- `src/gen-charts.ts`
- `src/gen-objects.ts`
