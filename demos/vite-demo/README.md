# Browser showcase (React + TypeScript + Vite)

Builds the **Meridian Q3 FY26 Business Review** — the same eleven-slide deck
`pnpm demos:build quarterly-review` produces — entirely in the browser. No server, no round
trip: the `.pptx` is assembled in the tab and handed to the download manager.

The deck code is not duplicated here. This app imports
`ts-pptx-demos-showcases/quarterly-review` and calls its `build()`, which is the whole
demonstration: nothing in the deck knows which runtime it is on. In Node `writeFile()`
writes a file; in a browser it triggers a download.

The Field Notes showcase is deliberately not offered here — it loads photographs and a video
from disk by path, which a browser cannot do unless those assets are served.

## Run it

```bash
pnpm install                            # once, at the repository root
pnpm --dir demos/vite-demo run dev      # dev server
pnpm --dir demos/vite-demo run build    # production build
pnpm --dir demos/vite-demo run preview  # preview a production build
```

## Notes

- This is the maintained browser integration path. Script tags, CDN globals, and IIFE
  bundles are not supported package targets.
- `vite build` reports `node:fs/promises` externalized for browser compatibility. That is
  the bundler noting an import it dropped on the browser path, not a broken build.
- This is a showcase, not a test. Nothing here gates a commit, and CI never builds it.
