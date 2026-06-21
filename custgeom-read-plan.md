# Plan: `dn-custgeom-read` — read accessor for `a:custGeom` freeform path geometry

> Status: **Step 1 fixture authored** — `test/read/fixtures/custgeom.pptx` exists,
> opens clean in desktop PowerPoint, and is documented (provenance/hash/purpose) in
> `test/read/fixtures/README.md`. Ready for Step 2 (accessor). The recorded authored
> XML — path `w`/`h`, ordered segments with literal `a:pt` coords, cubic control-point
> ordering, and shape `cx`/`cy` — is captured in that README entry; Step 5 asserts
> those literals.
> Backlog item `dn-custgeom-read` (deferred, p3). This doc carries the fixture spec
> so it can be authored, then the accessor implemented against real Office XML.

## Context

The downstream replication audit can detect `<a:custGeom>` freeform shapes, but
the `pptxgenjs/read` model exposes **no** accessor for their path geometry. The
first real consumer (the `202604-ia4cyb-overall:2` lab build — three native glyphs:
a gear + two rocket layers) had to reverse-engineer each path with a one-off
extractor and commit a generated stopgap (downstream
`decks/replication-lab/scripts/slides/slide-002/freeforms.ts`). The goal is a
first-class read accessor so a glyph replica becomes a copy-paste from `style.json`
(the route tables took), built on this reader.

This is **fixture-gated read-model work**: per `CLAUDE.md` ("OOXML And PowerPoint
Work") and `docs/backlog-workflow.md`, a read accessor whose oracle must be genuine
PowerPoint output may **not** be implemented against synthetic/hand-typed path XML.

The write side already defines the vocabulary we mirror:
- `GeometryPoint` DSL — `src/core-interfaces.ts:224` (moveTo/lnTo/cubic/quadratic/arc/close).
- `genXmlCustGeom()` emitter — `src/gen-xml.ts:261` (exact OOXML shape).

Schema confirmed via `ooxml` MCP (`CT_Path2D`):
- children (choice, repeatable): `close`, `moveTo`, `lnTo`, `arcTo`, `quadBezTo`, `cubicBezTo`
- attrs: `w` (default 0), `h` (default 0), `fill` = `ST_PathFillMode` (default `norm`), `stroke` (bool, default `true`), `extrusionOk` (bool, default `true`)
- segment points are `a:pt` with `x`/`y` (`ST_AdjCoordinate`); `arcTo` carries `wR`/`hR`/`stAng`/`swAng` and no `a:pt`.

---

## Step 1 — Fixture required (BLOCKING; author with PowerPoint)

Author **`test/read/fixtures/custgeom.pptx`** with the `powerpoint-fixture-authoring`
skill / desktop PowerPoint COM (Windows). Nothing below starts until it exists.

One slide with **freeform/custom-geometry shapes exercising every segment type**,
each given a stable shape name so the test can look it up (`shapeNamed(...)`):

1. **`freeform-lines`** — closed polygon using only `moveTo` + `lnTo` + `close`
   (e.g. a triangle). Pins the common case.
2. **`freeform-cubic`** — a path with at least one `cubicBezTo` (PowerPoint emits
   cubics when you curve a freeform node). Pins control-point ordering.
3. **`freeform-hole`** — a rectangle with an elliptical hole, authored via
   PowerPoint's Merge Shapes → Subtract (`ExecuteMso("ShapesSubtract")` on an
   overlapping rect + ellipse). **Empirical finding (2026-06-21):** desktop
   PowerPoint emits a hole as a **single** `a:path` holding **two closed contours**
   (`moveTo`…`close`, `moveTo`…`close`) in document order — *not* two `a:path`
   elements. Union/Combine/Subtract (even of disjoint shapes) all consolidate into
   one `a:path`; a genuine two-`a:path` `a:pathLst` is **not authorable** via
   PowerPoint's merge tools. So this shape pins the multi-contour / hole case
   (`paths.length === 1`, multiple `moveTo`+`close` + `cubicBezTo` within one path),
   which is the real shape of PowerPoint freeform output the glyph consumer hits;
   the original "two `a:path` elements" goal is dropped as unauthorable.

After authoring, **record the authored XML** so the test asserts real values:
- each `a:path` `w`/`h`/`fill`/`stroke`,
- the ordered segment list with literal `a:pt` `x`/`y` (and cubic control points),
- the shape `cx`/`cy` (EMU) from `a:xfrm/a:ext` — documents the path-unit →
  box-local scaling the consumer will do.

Then mirror the `preset-geometry.pptx` precedent in `test/read/fixtures/README.md`:
- add a provenance-table row ("Locally provided / locally authored": Application/AppVersion/Slides),
- add its **SHA-256** to the hash list (`shasum -a 256`),
- add a one-line purpose note (segments exercised).

> Do not proceed to Step 2 until `custgeom.pptx` is committed and documented.

---

## Step 2 — Accessor types (`src/read/api/shapes.ts`)

Faithful, multi-path read model (chosen over flattening to the write
`GeometryPoint[]` DSL: the `a:pathLst` schema allows repeatable `a:path`, each with
independent `fill`/`stroke`, which the single-path write DSL cannot represent; and
the read model elsewhere favors faithful exposure over lossy collapsing — cf.
`gradientStops`, `src/read/api/shapes.ts:620`). Command verbs reuse the write DSL
names so a consumer maps to `GeometryPoint[]` trivially.

> **Note (empirical, 2026-06-21):** desktop PowerPoint's own Merge Shapes never
> writes more than one `a:path` per `custGeom` — a hole is one `a:path` with two
> `moveTo`…`close` contours (see the `freeform-hole` fixture). So in practice the
> `paths[]` array length is 1 for PowerPoint-authored freeforms; multi-`a:path`
> input is schema-legal but comes from other producers (e.g. SVG import). The
> array model is still the faithful choice, but the per-path `fill`/`stroke`
> branch will not be exercised by a PowerPoint oracle.

```ts
/** One path segment; verbs mirror the write-side GeometryPoint DSL. Coordinates
 *  are raw path-unit integers in the path's own 0..w / 0..h space. */
export type GeometryCommand =
  | { cmd: 'moveTo'; x: number; y: number }
  | { cmd: 'lnTo'; x: number; y: number }
  | { cmd: 'cubicBezTo'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { cmd: 'quadBezTo'; x1: number; y1: number; x: number; y: number }
  | { cmd: 'arcTo'; wR: number; hR: number; stAng: number; swAng: number }  // angles in degrees
  | { cmd: 'close' }

/** One `<a:path>` in the path list, with its viewport extents + render attrs. */
export interface CustomGeometryPath {
  w: number          // a:path/@w (path-unit width; default 0)
  h: number          // a:path/@h (path-unit height; default 0)
  fill: string       // a:path/@fill ST_PathFillMode; default 'norm'
  stroke: boolean    // a:path/@stroke; default true
  commands: GeometryCommand[]
}

/** Custom freeform geometry (`spPr/a:custGeom`). */
export interface CustomGeometry {
  paths: CustomGeometryPath[]
}
```

Parse conventions:
- **Coordinates**: parse `a:pt/@x|@y`, `arcTo/@wR|@hR` with the existing `intValue(...)`
  helper. Authored freeforms use literal integers; a guide-name reference
  (`ST_AdjCoordinate` string form) is not expected — handle the numeric case and
  treat a non-numeric value as `0` without crashing (documented edge, not built out).
- **Arc angles**: `stAng`/`swAng` are authored in 60000ths of a degree; expose
  **degrees** (`raw / 60000`) to mirror the write DSL's degree input
  (`convertRotationDegrees`, `src/gen-xml.ts:281`). Document the unit in JSDoc.
- **Defaults**: apply schema defaults when an attr is absent (`fill='norm'`,
  `stroke=true`, `w=0`, `h=0`).

## Step 3 — Accessor implementation (`src/read/api/shapes.ts`)

Getter on the **`Shape` base class** (covers both `p:sp` autoshapes and `p:pic`
pictures clipped to a `custGeom`); both already route through `this.properties()`
(used by `adjustValues`/`resolvedFill`, `:599`/`:656`):

```ts
/** Custom freeform geometry (`spPr/a:custGeom/a:pathLst`), or `null` when the
 *  shape uses preset geometry / none. Pair with {@link presetGeometry}. */
get customGeometry(): CustomGeometry | null
```

1. `firstChild(this.properties(), 'a:custGeom')` → `null` if absent.
2. `firstChild(custGeom, 'a:pathLst')`; for each `getElements(pathLst, 'a:path')`
   build a `CustomGeometryPath` (read `w`/`h`/`fill`/`stroke` + defaults).
3. Walk the path's child elements **in document order** (order *is* the geometry —
   don't group by qname). Map by local name: `moveTo`/`lnTo` read one `a:pt`;
   `cubicBezTo` three `a:pt` (c1, c2, end); `quadBezTo` two (c1, end); `arcTo` the
   four attrs; `close` → `{ cmd: 'close' }`.

Reuse DOM helpers from `src/read/oxml/dom.ts` (`firstChild`, `getElements`, `attr`)
+ `intValue`. Add a tiny local ordered-children walk if one doesn't already exist;
no new module.

## Step 4 — Export types (`src/read.ts`)

Add `CustomGeometry`, `CustomGeometryPath`, `GeometryCommand` to the
`export { ... } from './read/api/shapes.js'` block (near `GradientStop`, `:33`).

## Step 5 — Test (`test/read/custgeom.test.js`, new)

Follow `test/read/style-accessors.test.js` (the `adjustValues`/preset-geometry
precedent, `:178`): `open('custgeom')`, `shapeNamed(...)`, `assertEqual` against the
**values recorded in Step 1**:
- `freeform-lines`: `paths.length === 1`; path `w`/`h`/`fill`/`stroke`; exact ordered
  `commands` with literal coords.
- `freeform-cubic`: a `cubicBezTo` with correct control-point ordering.
- `freeform-hole`: `paths.length === 1`; the ordered `commands` carry **two**
  `moveTo`+`close` contours in document order (ellipse hole built from four
  `cubicBezTo`, then the outer rectangle from three `lnTo`) — pins multi-contour
  single-path traversal.
- a non-freeform shape (`preset-rect`): `customGeometry === null`.

Assert literal numbers from real PowerPoint output — never synthesize XML.

## Step 6 — Backlog + changelog bookkeeping

- `docs/backlog.yml` `dn-custgeom-read`: `status: implemented`, `last_reviewed` →
  today, `next_action: none`, populate `evidence.local_files` (accessor + test),
  `evidence.kinds` (+ `powerpoint-result`), `evidence.powerpoint_result` (the
  fixture), update `current_project_notes`. Keep the `stopgap` line until
  downstream migrates off `freeforms.ts` (note the reader exists; harvester
  routing is the downstream follow-up). Use only `vocabulary:`-listed values;
  run `pnpm run backlog:validate`.
- `CHANGELOG.md`: add the new read accessor under the read subsystem.
- Read-side add-only (no emit changed) → no `test/schema.test.js` fixture needed.

## Verification

1. `pnpm run build && pnpm run typecheck`.
2. `pnpm run test:unit` (read suite) — `custgeom.test.js` passes against the fixture.
3. `pnpm run backlog:validate` — ledger clean.
4. Sanity: `customGeometry` is `null` for a preset shape, populated for the freeforms.

## Out of scope (note, don't build)

- `a:custGeom` guide lists (`a:gdLst`/`a:avLst`/`a:ahLst`/`a:cxnLst`/`a:rect`) — only
  `a:pathLst` is read.
- Guide-name (string) coordinate references in `a:pt`.
- downstream harvester routing (emit `freeform.points` into `style.json`) —
  downstream consumer change, tracked in downstream.
