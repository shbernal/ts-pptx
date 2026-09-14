---
doc-schema-version: 1
title: "Animations and transitions"
summary: "Slide transitions and preset build animations on shapes: the options, how effects group into click steps, what importing a slide or shape does to them, and reading transitions back."
read_when:
  - Adding a transition between slides
  - Animating a shape on entrance, emphasis or exit
  - Importing a shape or slide that carries animations
  - Reading or changing the transition on an existing slide
doc_type: "guide"
---

# Animations and transitions

A transition plays as the slide show moves onto a slide. A build animation plays an effect on one
shape. Both belong to the slide:

```ts
const slide = pptx.addSlide()
slide.addText("Revenue up 12%", { x: 1, y: 1, w: 8, h: 1, fontSize: 32, objectName: "headline" })

slide.transition = { type: "push", durationMs: 800, variant: { dir: "u" } }
slide.addAnimation({ preset: "fadeIn", objectName: "headline" })
```

## Transition options

| Option | Type | Default | Effect |
| --- | --- | --- | --- |
| `type` | `TransitionType` | required | Which transition plays. See [transition types](#transition-types). |
| `durationMs` | number | none | The exact duration, which PowerPoint 2010 and later use. |
| `speed` | `'slow'`, `'med'` or `'fast'` | derived from `durationMs`, otherwise `fast` | The coarse speed older readers fall back to. |
| `advanceOnClick` | boolean | `true` | Whether a click moves past the slide. |
| `advanceAfterMs` | number | none | Moves past the slide automatically after this many milliseconds. |
| `variant` | `Record<string, string>` | none | The transition's own settings, such as the direction of a push. |
| `sound` | `TransitionSoundProps` | none | A sound played with the transition. |

When only `durationMs` is given, the speed follows it: up to 500 ms is `fast`, up to 1000 ms is `med`,
and anything longer is `slow`. Set `slide.transition = undefined` to remove a transition.

### Transition types

| Types | `variant` keys and values |
| --- | --- |
| `blinds`, `checker`, `comb`, `randomBar` | `dir`: `horz` (default) or `vert` |
| `cover`, `pull` | `dir`: `l` (default), `u`, `r`, `d`, or a corner `lu`, `ru`, `ld`, `rd` |
| `push`, `wipe` | `dir`: `l` (default), `u`, `r` or `d` |
| `strips` | `dir`: `lu` (default), `ru`, `ld` or `rd` |
| `zoom` | `dir`: `out` (default) or `in` |
| `split` | `orient`: `horz` (default) or `vert`; `dir`: `out` (default) or `in` |
| `wheel` | `spokes`: a whole number, `4` by default |
| `cut`, `fade` | `thruBlk`: `true` to pass through black, `false` by default |
| `circle`, `diamond`, `dissolve`, `newsflash`, `plus`, `random`, `wedge` | none |

### Add a sound

A transition sound is a WAV file embedded in the deck, given by `path` or as base64 `data`:

```ts
slide.transition = { type: "fade", sound: { path: "chime.wav", loop: false } }
```

`name` sets the name PowerPoint shows for the sound, and `loop: true` repeats it until the next sound
starts. `{ stopPrevious: true }` instead silences a sound an earlier slide started. PowerPoint's
built-in sounds are ordinary WAV files, so embed the one you want the same way.

## Animation options

| Option | Type | Default | Effect |
| --- | --- | --- | --- |
| `preset` | `PresetEffect` | required | The effect. See [presets](#presets). |
| `objectName` | string | none | The target shape, by name, including a shape inside a group. |
| `shapeIndex` | number | none | The target shape by its 0-based add order among the slide's top-level shapes. |
| `trigger` | `'onClick'`, `'withPrevious'` or `'afterPrevious'` | `'onClick'` | When the effect starts. |
| `durationMs` | number | the preset's | How long the effect runs. |

Name the target with `objectName` or `shapeIndex`. `shapeIndex` counts only objects that draw something,
so speaker notes added between two shapes do not shift the count.

### Presets

| Preset | Kind | Default duration |
| --- | --- | --- |
| `appear` | entrance | 500 ms |
| `fadeIn` | entrance | 500 ms |
| `flyIn` | entrance | 500 ms |
| `wipe` | entrance | 500 ms |
| `grow` | emphasis | 2000 ms |
| `spin` | emphasis | 2000 ms |
| `fadeOut` | exit | 500 ms |
| `flyOut` | exit | 500 ms |

## Order effects into click steps

Effects play in the order they were added, and `trigger` decides how they group:

- `onClick` starts a new click step. The first effect on a slide always starts one.
- `withPrevious` starts at the same moment as the effect added before it.
- `afterPrevious` starts once the effect added before it has run for its duration.

```ts
slide.addAnimation({ preset: "fadeIn", objectName: "title" })                               // first click
slide.addAnimation({ preset: "fadeIn", objectName: "subtitle", trigger: "withPrevious" })   // with the title
slide.addAnimation({ preset: "grow", objectName: "chart", trigger: "afterPrevious" })       // 500 ms later
slide.addAnimation({ preset: "fadeOut", objectName: "title" })                              // second click
```

## Invalid input

| Condition | Result | Code |
| --- | --- | --- |
| `type` is not one of the transition types | warns; the slide is written with no transition | `transition/unknown-type` |
| a `variant` key the transition does not take | warns; that key is left out | `transition/unknown-variant` |
| an animation with neither `objectName` nor `shapeIndex` | warns; the effect is dropped | `animation/target-missing` |
| `objectName` names nothing on the slide | warns; the effect is dropped | `animation/target-not-found` |
| `shapeIndex` is past the last shape | warns; the effect is dropped | `animation/target-index-out-of-range` |
| reading side: `durationMs` or `advanceAfterMs` is not a finite number from 0 | throws `InvalidOptionError` | `transition/invalid-time` |
| reading side: a `sound` other than the one the slide already has | throws `InvalidOptionError` | `transition/sound-unsupported` |

## Import slides and shapes

| Operation | What happens to animations |
| --- | --- |
| `importSlide`, `importSlides` | The slide's transition and animations come with it. |
| `importShape`, `importShapes` | The shape arrives static. With `{ carryAnimation: true }`, its click steps are added after the destination slide's own and retargeted at the new shape. |
| `slide.flattenAnimations()` | Removes every build animation, so every shape shows at once. Media playback settings stay. |

## Limits

- Eight presets. There are no motion paths, custom effects or per-effect settings beyond `durationMs`.
- Transitions newer than the list above, such as Morph, can be read and are preserved, but cannot be authored.
- `variant` values are written as given; only the keys are checked.
- In a deck opened for reading, animations are kept as they are: you can check for them, remove them, or carry a shape's along on import, but not list or edit individual effects.
- The reading side's transition setter can keep or remove a sound, not add a new one. Add sounds when authoring the deck.

## Reading it back

A slide opened through `pptx-ts/read` exposes:

- `slide.transition`: the type, `speed`, `durationMs`, `advanceOnClick`, `advanceAfterMs`, `variant` and
  `sound`, or `null`. It is also settable, so `slide.transition = { ...slide.transition, speed: "slow" }`
  changes the speed and keeps the rest, and `slide.transition = null` removes the transition.
- `slide.hasAnimations`: whether the slide has build animations.
- `slide.flattenAnimations()`: removes them.

See [PPTX read and round-trip](reference/pptx-read.md).

## See also

- [PPTX read and round-trip](reference/pptx-read.md)
- API reference: [`TransitionProps`](reference/api/index/interfaces/TransitionProps.md),
  [`TransitionSoundProps`](reference/api/index/interfaces/TransitionSoundProps.md),
  [`AnimationProps`](reference/api/index/interfaces/AnimationProps.md)
