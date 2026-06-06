# Vite TypeScript Demo

This demo exercises the maintained browser application target: React,
TypeScript, Vite, and ESM package consumption.

## Run From The Repository

Install dependencies at the repository root:

```bash
pnpm install
```

Build the demo:

```bash
pnpm --dir demos/vite-demo run build
```

Run the development server:

```bash
pnpm --dir demos/vite-demo run dev
```

Preview a production build:

```bash
pnpm --dir demos/vite-demo run preview
```

## Package Smoke Path

To test this demo against a packed package instead of the workspace dependency,
use the repository smoke command:

```bash
pnpm run test:demo:vite
```

## Notes

- This is the maintained browser integration path.
- Direct script tags, CDN globals, and IIFE bundles are not supported package
  targets.
