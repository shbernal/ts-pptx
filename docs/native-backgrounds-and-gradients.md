---
doc-schema-version: 1
title: "Native Backgrounds And Gradients"
summary: "Implementation plan for native PPTX gradient fills on slide backgrounds and shapes."
read_when:
  - Planning native gradient fill support
  - Changing slide background or shape fill OOXML
  - Deciding whether downstream can replace raster gradient backgrounds
doc_type: "decision"
---

# Native Backgrounds And Gradients

## Status

The first schema-validated implementation slice is in place for simple linear
gradient fills on slide backgrounds and shape fills.

The configured `ooxml` and `microsoft_learn` MCP servers were consulted on
2026-06-07. The OOXML shape below is schema-grounded, but PowerPoint-authored
comparison fixtures are still required before broadening the API or treating
native gradients as a drop-in replacement for deterministic raster gradients.

## Goal

Add reliable native PPTX gradient fill support for reusable PptxGenJS surfaces,
starting with simple linear gradients on slide backgrounds and shape fills.

This belongs in PptxGenJS when the behavior is generic, schema-valid, and stable
across PowerPoint-compatible consumers. Downstream should continue using
deterministic raster backgrounds when exact rendered appearance is more reliable
than native PPTX gradients.

## MCP Review Findings

Sources consulted:

- `ooxml` MCP schema lookups for `p:bgPr`, `a:gradFill`, `a:gsLst`,
  `a:gs`, `a:lin`, `a:EG_FillProperties`, `a:EG_EffectProperties`,
  `a:EG_ShadeProperties`, `a:ST_PositiveFixedPercentage`,
  `a:ST_PositiveFixedAngle`, and `a:ST_TileFlipMode`.
- `ooxml` MCP prose sections: ECMA-376 Part 1 sections 19.3.1.2,
  20.1.8.33, 20.1.8.36, 20.1.8.37, 20.1.8.41, 20.1.10.43, and
  20.1.10.44.
- `microsoft_learn` MCP pages for Open XML SDK
  [`GradientFill`](https://learn.microsoft.com/dotnet/api/documentformat.openxml.drawing.gradientfill?view=openxml-3.0.1)
  and
  [`ShapeProperties`](https://learn.microsoft.com/dotnet/api/documentformat.openxml.presentation.shapeproperties?view=openxml-3.0.1),
  plus PowerPoint JavaScript API
  [`SlideBackgroundFill`](https://learn.microsoft.com/javascript/api/powerpoint/powerpoint.slidebackgroundfill?view=powerpoint-js-preview),
  [`SlideBackgroundGradientFillOptions`](https://learn.microsoft.com/javascript/api/powerpoint/powerpoint.slidebackgroundgradientfilloptions?view=powerpoint-js-preview),
  and
  [`SlideBackgroundGradientFillType`](https://learn.microsoft.com/javascript/api/powerpoint/powerpoint.slidebackgroundgradientfilltype?view=powerpoint-js-preview).

Confirmed XML constraints for the first slice:

1. `p:bgPr` uses `CT_BackgroundProperties`: exactly one DrawingML fill choice
   first, then optional DrawingML effects, then optional `p:extLst`. Keep
   emitting `<a:effectLst/>` after the fill to match the current regression
   guard and schema order.
2. Legal `p:bgPr` fill choices are `a:noFill`, `a:solidFill`, `a:gradFill`,
   `a:blipFill`, `a:pattFill`, and `a:grpFill`. Image backgrounds should keep
   using `a:blipFill`; this plan only adds `a:gradFill`.
3. Presentation shape `p:spPr` uses the same DrawingML fill group after geometry
   and before `a:ln`. The existing shape fill insertion point is the right one.
4. `a:gradFill` serializes children in this order: optional `a:gsLst`, optional
   shade properties (`a:lin` or `a:path`), then optional `a:tileRect`. The
   first implementation should emit `a:gsLst` followed by `a:lin` and should
   omit `a:path` and `a:tileRect`.
5. `a:gradFill` supports optional `rotWithShape` and `flip` attributes. `flip`
   values are `none`, `x`, `y`, and `xy`, but it is only relevant once
   `a:tileRect` is supported, so leave it out of the public API initially.
6. `a:gsLst` contains two or more `a:gs` children. Each `a:gs` requires a
   `pos` attribute and exactly one DrawingML color choice child.
7. `a:gs/@pos` is an `ST_PositiveFixedPercentage`; public stop positions should
   be finite `0..100` percentages and serialize as `0..100000`.
8. `a:lin/@ang` is an `ST_PositiveFixedAngle` in 60000ths of a degree with
   legal values from `0` through less than `21600000`. Public angles must be
   normalized into `0 <= angle < 360` before serialization. Do not allow
   `360` to serialize as `21600000`.
9. `a:lin/@scaled` is an optional boolean. Emit it only when the public API
   specifies it, or choose a documented default after comparing PowerPoint
   fixtures.
10. `p:bgPr/@shadeToTitle` and the PowerPoint API
    `ShadeFromTitle` gradient type are separate from simple linear gradients.
    Leave shade-to-title support out of the first slice.

Microsoft Learn confirms that PowerPoint exposes background fill types
`Solid`, `Gradient`, `PictureOrTexture`, and `Pattern`, and gradient subtypes
`Linear`, `Radial`, `Rectangular`, `Path`, and `ShadeFromTitle`. The JavaScript
API only exposes a high-level gradient type option for slide backgrounds, so it
is useful compatibility signal but not a substitute for inspecting
PowerPoint-authored package XML.

## PowerPoint Fixture Review Still Outstanding

Before broadening support beyond the first linear-gradient slice, complete the
behavior review that MCP lookup cannot replace:

1. Compare at least one minimal PowerPoint-authored PPTX fixture for a slide
   background gradient and one for a shape gradient.
2. Confirm PowerPoint's authored defaults for `a:gradFill/@rotWithShape`,
   `a:lin/@scaled`, gradient stop ordering, and whether it writes two stops or
   inserts midpoint stops for common presets.
3. Record only small repo-specific findings and section references. Do not
   vendor full standards text or large extracted reference material.

Do not expand the public guarantee until the PowerPoint fixture review confirms
PowerPoint's authored XML shape and rendering behavior.

## Proposed API Slice

Start with a narrow `ShapeFillProps` extension for linear gradients:

```ts
fill: {
	type: 'gradient',
	gradient: {
		kind: 'linear',
		angle: 90,
		scaled: true,
		rotateWithShape: true,
		stops: [
			{ position: 0, color: '451DC7' },
			{ position: 100, color: '0B003D', transparency: 10 },
		],
	},
}
```

The first slice should avoid path, radial, tile-rectangle, table, chart, and
line-gradient promises until each parent surface is independently validated.

The same model should apply directly to slide backgrounds because
`BackgroundProps` already extends `ShapeFillProps`:

```ts
slide.background = {
	type: 'gradient',
	gradient: {
		kind: 'linear',
		angle: 90,
		stops: [
			{ position: 0, color: '451DC7' },
			{ position: 100, color: '0B003D' },
		],
	},
}
```

Initial validation rules:

1. Require `type: 'gradient'` plus `gradient.kind: 'linear'`.
2. Require at least two stops.
3. Require finite stop positions in `0..100`; reject or warn on invalid input
   according to the local fill API style chosen during implementation.
4. Prefer nondecreasing stop positions for deterministic output. If the helper
   sorts stops, document that behavior in the public API docs and tests.
5. Require finite angles, normalize to `0 <= angle < 360`, and serialize only
   legal `ST_PositiveFixedAngle` values.
6. Support `HexColor`, `ThemeColor`, and existing alpha/transparency handling
   for each stop by routing stop colors through `createColorElement()`.

## Implementation Steps

1. Extend `ShapeFillProps` in `src/core-interfaces.ts` with a linear gradient
   model and exported stop type. Expand `type` from `'none' | 'solid'` to
   include `'gradient'`.
2. Add a `genXmlGradientFill()` helper in `src/gen-utils.ts`, called by
   `genXmlColorSelection()` when `type: 'gradient'`. Keep `genXmlColorSelection`
   returning a complete fill-choice element.
3. Reuse `createColorElement()` for stop colors so hex, theme color, and alpha
   behavior stays consistent with existing solid fills.
4. Convert stop positions from public `0..100` values to OOXML `0..100000`
   values with `Math.round(position * 1000)`.
5. Normalize gradient angles before calling `convertRotationDegrees()`, or
   harden `convertRotationDegrees()` so `360`, negative values, and values above
   one turn cannot produce invalid `ST_PositiveFixedAngle` values.
6. Serialize the first slice as:

```xml
<a:gradFill rotWithShape="1">
	<a:gsLst>
		<a:gs pos="0"><a:srgbClr val="451DC7"/></a:gs>
		<a:gs pos="100000"><a:srgbClr val="0B003D"/></a:gs>
	</a:gsLst>
	<a:lin ang="5400000" scaled="1"/>
</a:gradFill>
```

7. Update slide background emission in `src/gen-xml.ts` so native background
   fills are emitted when `slide.background` contains either a solid color or a
   gradient, while `_bkgdImgRid` image backgrounds keep taking precedence.
   The current `slide.background?.color` condition is too narrow for gradient
   backgrounds.
8. Keep shape fill support on the existing `genXmlColorSelection()` path so
   `p:spPr` still emits geometry, then fill, then `a:ln`.
9. Avoid using deprecated `background.fill` as the gradient API. It currently
   aliases a legacy color string to `background.color`; direct
   `background: { type: 'gradient', gradient: ... }` is clearer and matches
   `BackgroundProps extends ShapeFillProps`.
10. Leave chart fills, table fills, and line fills out of the initial public
   guarantee unless the full review proves their current call sites can support
   gradient-only fill objects safely.

## Fixtures And Tests

Add executable evidence with a small blast radius:

1. Add regression tests that inspect `ppt/slides/slide1.xml` for `a:gradFill`
   under a shape `p:spPr` and under slide background `p:bgPr`.
2. Add schema fixtures in `test/schema.test.js` for a native shape gradient and
   a native slide background gradient.
3. Assert that slide background gradients keep `<a:effectLst/>` after
   `a:gradFill` inside `p:bgPr`.
4. Assert that shape gradients appear in the fill slot before `a:ln` inside
   `p:spPr`.
5. Add invalid-input coverage for fewer than two stops, out-of-range stop
   positions, and angle normalization around `0`, `359`, `360`, and negative
   values.
6. Add or document minimal PowerPoint-authored comparison fixtures for the same
   two cases.
7. Run:

```bash
pnpm run build
pnpm run typecheck
pnpm run test:unit
pnpm run test:schema
```

If `test:schema` cannot run because the validator is missing, install it with
`./tools/ooxml-validator/install.sh` before treating the implementation as
accepted.

## Downstream Fallback Guidance

Do not remove raster gradient guidance from downstream immediately.

Native gradients can replace deterministic raster backgrounds only after:

1. The generated PPTX opens without repair in PowerPoint.
2. The gradient survives import or rendering through the target review tools.
3. The rendered result is close enough for the specific deck workflow.
4. A single native background object remains easier to preserve semantically
   than an image object.

Until then, downstream should continue preferring one deterministic image
object over many adjacent solid-color rectangles for continuous gradient
backgrounds.
