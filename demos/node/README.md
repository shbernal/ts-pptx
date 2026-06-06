# Node.js Demo

This demo exercises the maintained Node.js package target: ESM usage on
Node.js `>=24`.

## Run From The Repository

Install dependencies at the repository root:

```bash
pnpm install
```

Run a focused text demo:

```bash
pnpm --dir demos/node run demo-text
```

Run the default demo:

```bash
pnpm --dir demos/node run demo
```

Run all demo objects:

```bash
pnpm --dir demos/node run demo-all
```

Generated decks are written to `demos/node/output/`. The output directory is
ignored by git, and each demo command overwrites its previous `.pptx`.

Run the stream demo:

```bash
pnpm --dir demos/node run demo-stream
```

Then visit `http://localhost:3000/` in a browser to download the streamed
presentation.

## Demo Smoke Path

To test this demo through the repository smoke command:

```bash
pnpm run test:demo:node
```

## Notes

- This demo is ESM-only.
- CommonJS `require("pptxgenjs")` is not supported.
