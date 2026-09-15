---
doc-schema-version: 1
title: "OOXML agent context"
summary: "Where to look up OOXML and PowerPoint behavior, in what order, how emitters build XML, what counts as evidence for a change, and when to wait for a genuine PowerPoint fixture instead of guessing."
read_when:
  - Changing emitted OOXML
  - Researching PowerPoint or Open XML compatibility
  - Updating OOXML validation guidance
  - Deciding whether there is enough evidence to start implementing
  - Implementing a read accessor or write-side behaviour with no fixture yet
  - Authoring a PowerPoint-produced fixture or oracle
doc_type: "reference"
---

# OOXML agent context

Read this before changing emitted OOXML. It sets where to look a question up, how new emitter code
builds XML, what counts as evidence, and when to stop and wait for a PowerPoint-authored fixture.

## Lookup order

1. **Local evidence.** Search `src/`, `test/` and the [testing guide](testing.md) before changing
   behavior.
2. **The `ooxml` MCP**, for schema structure. After a validation failure, `ooxml_explain` on the
   diagnostic is usually the fastest next step.
3. **The `microsoft_learn` MCP**, for everything the schema graph does not answer. Always try it
   before web search.
4. **Web search**, last: community findings, third-party library behavior, and anything newer than
   both servers.

When two sources disagree, prefer PowerPoint and Open XML SDK behavior, because the bar is a
`.pptx` that PowerPoint opens cleanly. Record the difference in a code comment or a test name only
when it changes the implementation.

`.mcp.json` configures both servers for Claude Code, and `.codex/config.toml` for Codex, which
loads it only in a trusted project. `/mcp` lists what loaded in a running session.

## The `ooxml` MCP

[`mcp-server-ooxml`](https://www.npmjs.com/package/mcp-server-ooxml) runs locally over stdio
(`npx -y mcp-server-ooxml`) and needs Node 24+. It serves the ECMA-376 XSDs, Transitional and
Strict, as a SQLite graph shipped inside the package. There is no account and no network call after
the first download, and the same question returns the same answer every time.

It answers:

- which children are legal under an element, and in what order (`ooxml_children`,
  `ooxml_element`, `ooxml_type`);
- which attributes, values, enums and patterns are valid (`ooxml_attributes`, `ooxml_values`,
  `ooxml_enum`);
- which namespace URI goes with a prefix (`ooxml_namespace`);
- what Transitional adds over Strict for one name, and so whether a construct survives in Strict
  (`ooxml_diff_profiles`);
- what was legal at the position an `ooxml-validate` diagnostic names. Hand `ooxml_explain` the
  diagnostic's `id`, `description`, `xpath` and `partUri`. It explains validator output and does
  not validate.

`ooxml_search` is a case-insensitive substring match on symbol names. It is not semantic and not
full-text, so "how do I make text bold" finds nothing.

It has no specification prose, no OPC part, content-type or relationship catalogue, no
Microsoft-proprietary detail, and no Annex D geometry (see
[Preset shape geometry](#preset-shape-geometry)). Those questions go to `microsoft_learn`.
ooxml.dev is the hosted service that indexes the specification prose. The server is pre-1.0, and a
change to its tools is announced in its
[changelog](https://github.com/shbernal/ooxml-ai-tooling/blob/main/CHANGELOG.md).

## The `microsoft_learn` MCP

`https://learn.microsoft.com/api/mcp` serves official Microsoft documentation. Use it for:

- the Open Specifications, such as [MS-OE376], [MS-PPTX] and [MS-OI29500];
- built-in GUIDs and other proprietary enumerations;
- how PowerPoint accepts, repairs, extends or rejects OOXML, across versions;
- the Open XML SDK, and how it validates or models a part;
- OPC parts, content types and relationship types, and whether a namespace, extension URI or
  relationship type is Microsoft-specific.

Search with `microsoft_docs_search`, then fetch a returned page with `microsoft_docs_fetch`.
`microsoft_code_sample_search` finds official code samples.

## Preset shape geometry

The per-shape adjust guide names and geometry formulas live in the ECMA-376 Annex D electronic
addenda (`OfficeOpenXML-DrawingMLGeometries.zip` inside the Part 1 ZIP), not in any XSD, so the
`ooxml` MCP cannot answer "round2SameRect adj1 adj2".

- For guide names, use `docs/preset-shape-adj-guides.tsv`: every preset that takes adjust values,
  mapped to its guide names. It is sourced from Annex D and cross-checked against LibreOffice
  `oox-drawingml-adj-names` and ONLYOFFICE `OOXMLShapes/`.
- For full formulas (`avLst` defaults, `gdLst`, paths), use
  [ONLYOFFICE/core `OOXMLShapes/`](https://github.com/ONLYOFFICE/core/tree/master/MsBinaryFile/Common/Vml/PPTXShape/OOXMLShapes),
  one `.cpp` file per shape with its `avLst` and `gdLst` inline, and
  [LibreOffice/core `presetooxhandleadjustmentrelations.cxx`](https://github.com/LibreOffice/core/blob/master/svx/source/svdraw/presetooxhandleadjustmentrelations.cxx)
  for handle constraints.

### Do not clamp adjust values

`genXmlPresetGeom` writes any finite adjust value as given. It refuses only a non-finite value and
a `rectRadius` with no side to divide by, which would write `val NaN` or `val Infinity`, a token
outside the formula grammar.

Do not build a per-preset range table to clamp against. PowerPoint stores an out-of-range guide
verbatim, opens the deck without a repair prompt, and pins the value at render time through the
preset's own guide formulas: `roundRect` at `adj` 50000, 60000 and 266667 exports identical PNGs,
while two in-range values do not. Clamping would only discard what the caller wrote. The bounds sit
in each preset's handle definitions, often as formulas over width and height, and `a:gd/@fmla` is
`xsd:string`, so no schema check sees a violation. The `prstgeom` leg of `pnpm run test:com` keeps
that evidence runnable, with its sensitivity pair.

## Emitting XML: the `el()` builder

`src/gen/oxml/el.ts` is the write-side element builder, the mirror of `src/read/oxml/dom.ts`. Use
it instead of template strings in new emitter code. It escapes text and attribute values in one
place, so a forgotten `encodeXmlEntities` cannot produce invalid XML.

- `el(name, attrs, children, fmt)` always emits a paired tag. `voidEl(name, attrs, fmt)` always
  self-closes. The function you call decides, never the child's value. `encodeXmlEntities(undefined)`
  is `''`, so a rule based on the value would turn `<dc:title></dc:title>` into `<dc:title/>`.
- `raw(xml)` inserts markup that is already serialized, or a value that deliberately skips escaping.
- The builder drops nullish attributes and children, so an optional part inlines as
  `cond ? raw(...) : null`.
- `fmt` (`openPrefix`, `childPrefix`, `closePrefix`) places whitespace explicitly. Most parts are
  flat and need none. The pretty-printed parts are not always depth-regular, so each element states
  its own indentation.
- Attribute values and text children use different escapers. Attributes go through
  `encodeXmlAttrValue`, which also writes tab, LF and CR as `&#9;`, `&#10;` and `&#13;`, because XML
  1.0 §3.3.3 has a parser turn those literal characters into spaces inside an attribute value. Text
  children go through `encodeXmlEntities`, where the same characters are content and stay literal.
  An emitter that writes an attribute with a template string, such as `cNvPrOpen`, must call
  `encodeXmlAttrValue` itself.

Moving an existing emitter onto the builder must not change a byte. Run
`pnpm run byte-identity:baseline` before the change and `pnpm run byte-identity:check` after each
step. `raw-xml:check` counts the hand-built tags left in `src/`, and the count may fall but never
rise.

## Evidence and fixtures

Do not start implementing without at least one evidence path from this checkout:

- a minimal ts-pptx reproduction;
- generated `.pptx` output;
- an extracted package XML path and the observed problem;
- a `pnpm run test:schema` result, or a planned fixture;
- a PowerPoint repair or open result, when available;
- an Open XML SDK or Microsoft documentation reference, when PowerPoint behavior is not obvious
  from the schema.

An emitted OOXML change carries a focused fixture in `test/schema-cases.js`.
[OOXML schema validation](testing.md#ooxml-schema-validation) covers running it.

### Fixture-gated work waits for the fixture

Some work can only be judged against XML PowerPoint itself wrote: a read accessor checked against
real Office XML, or a write path whose target is whatever PowerPoint authors for preset ids, part
wiring, namespaces or inheritance. Hand-typed XML, synthetic XML and write-then-read round trips all
agree with whatever the code already does, which is exactly the question being asked.

So when that fixture does not exist, stop. Open a GitHub issue naming the exact construct the
fixture has to contain, and leave the feature unimplemented until it lands. Then:

1. Author the fixture with the `powerpoint-fixture-authoring` skill, and verify it with
   `.agents/skills/powerpoint-fixture-authoring/scripts/verify-powerpoint-fixture.ps1`.
2. Record provenance and SHA-256 in
   [test/read/fixtures/README.md](https://github.com/shbernal/ts-pptx/blob/master/test/read/fixtures/README.md).
3. Wire the test to the fixture: through the read harness for a read accessor, or a
   `test/schema-cases.js` comparison for a write-side oracle.
4. Implement last.

A claim about whether PowerPoint paints a construct needs render evidence:
[Check rendering with pixels, not COM properties](testing.md#check-rendering-with-pixels-not-com-properties).

## What not to do

- Do not commit OOXML standards PDFs, large copied specification excerpts, or bulk extracted
  standard text.
- Do not rely on schema validity alone for compatibility. Some PowerPoint repair prompts come from
  implementation constraints that only the Open XML SDK validator or PowerPoint itself catches.
- Do not introduce an XML ordering rule without a local fixture, a PowerPoint-authored comparison,
  or a referenced standard.
- Do not treat a Microsoft extension namespace as ECMA-defined without checking the Microsoft
  documentation.

## Local files

- `test/schema-cases.js` holds the schema fixtures, `test/schema-validation.test.js` runs them, and
  `test/validator.js` adapts `ooxml-validate`, which fetches and caches its oracle binary on first
  use.
- `src/gen/` holds the generators: `define/*` normalizes options, and
  `slide|drawingml|chart|pres|opc|anim|table/*` serialize.
- `scripts/byte-identity.mjs` is the byte-identity gate for emitter refactors.
- `scripts/raw-xml-ratchet.mjs` keeps the count of hand-built XML from rising.
