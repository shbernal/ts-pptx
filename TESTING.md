# PptxGenJS Testing Guide

This document outlines how to manually test PptxGenJS across supported platforms and environments prior to release.

> ✅ Run these tests to ensure compatibility with major bundlers, runtimes, and front-end frameworks.

Config Notes

> ⚠️ Disable VPN on the server machine, otherwise, clients using the local IP address cannot connect.

Testing Steps

1. Run `npm run ship` to refresh release artefacts (`dist/pptxgen.{cjs,es,min}.js`, `demos/browser/js/pptxgen.bundle.js`).
2. Run `npm run release-test` for the automated end-to-end suite — drives the browser demo, Web Worker demo, Node CLI demo, Node stream demo, and Vite build, validating each generated `.pptx` against the OOXML schema.
3. Run the manual sections below for items that require human eyes (Microsoft 365 web viewer, iOS rendering, PowerPoint visual inspection).

## 🧪 Test Suites Overview

| Platform        | Tooling              | Status |
| --------------- | -------------------- | ------ |
| Browser         | Standalone HTML demo | ✅      |
| Node.js         | Native CLI           | ✅      |
| Web Worker      | JS Worker demo       | ✅      |
| Vite/TypeScript | Modern front-end SPA | ✅      |
| Webpack         | SharePoint Framework | ✅      |

---

## 🌐 Browser Tests

**Purpose:** Validate browser compatibility using the standalone bundle as script.

### Automated Coverage

Browser desktop and Web Worker paths are exercised by `npm run release-test` (see `test/release/browser.test.js` and `test/release/worker.test.js`). The harness drives every `#btnRunBasicDemo` / `#btnRunSandboxDemo` / `#btnGenFunc_*` / `#btnRunAllDemos` button on `demos/browser/index.html` plus the `#generatePptWorker` flow on `demos/browser/worker_test.html`, then validates each generated `.pptx` against the OOXML schema.

The manual desktop / iOS sections below remain for human-eye verification (visual rendering, gesture handling, and devices the headless harness does not cover).

### Desktop & Mobile Browsers

Run local test server:

```bash
cd demos
node browser_server.mjs
```

1. Open the [Demo Page](http://localhost:8000/browser/index.html).
2. In DevTools, confirm the latest `pptxgen.bundle.js` is loaded (`Sources` tab).
3. Run all UI-driven demos and verify demo presentation render correctly.
4. Open the [Demo Page](http://192.168.254.x:8000/browser/index.html) on iPhone & test.

### Web Worker API

1. Open the [Web Worker Demo Page](localhost:8000/browser/worker_test.html).
2. Note: Use Chrome (Safari *will not work*)
3. Run the test; verify result & library version

### Microsoft 365 Check

1. Upload the full demo output from above to M365/Office/OneDrive.
2. Use web viewer to validate file

---

## 📦 Node.js Tests

**Purpose:** Validate functionality of CommonJS module in pure Node environments.

Automated by `npm run release-test`. The harness spawns `node demo.js`, `node demo.js All`, and `node demo_stream.js` from `demos/node/` (via `npm install --prefix demos/node` on first run) and validates each generated `.pptx` against the OOXML schema. See `test/release/node.test.js` for the full case list.

---

## ⚛️ Vite + TypeScript Tests

**Purpose:** Validate integration in modern front-end SPA toolchains (Vite, TypeScript, React-compatible).

Automated by `npm run release-test`. The harness runs `tsc -b && vite build` against `demos/vite-demo/` (via `npm install --prefix demos/vite-demo` on first run, with a post-install `gulp reactTestCode reactTestDefs` to override the published artefacts with the freshly-built `dist/pptxgen.es.js` and `types/index.d.ts`) and asserts the entry HTML plus at least one hashed JS chunk under `dist/assets/`. See `test/release/vite.test.js`.

### IDE IntelliSense (Manual)

Type-definition autocomplete still warrants a quick manual check in an IDE that the headless harness cannot replicate:

- Open `demos/vite-demo/src/tstest/Test.tsx`.
- Use IntelliSense to autocomplete things like `pptxgen.ChartType.`.

### Mobile Smoke (Manual)

For iOS / Android visual inspection, run the dev server interactively:

```bash
cd demos/vite-demo
npm run dev
```

Then export and open a `.pptx` on each device to verify MIME handling and visual fidelity.

---

## 🚀 Build for gh-pages (Manual)

After confirming the above:

```bash
npm run build
```

1. Copy the entire `dist` folder from `demos/vite-demo/` to a safe location.
2. Use this copy when updating the `gh-pages` branch after the release.

> ⚠️ DO NOT use the "deploy" script displayed onscreen by Vite. Manual copying ensures full control over final content.

---

## 🏁 Test Completion Checklist

| Dist File         | Test       | Tested Via             | Automation                                        | Result |
| ----------------- | ---------- | ---------------------- | ------------------------------------------------- | ------ |
| pptxgen.es.js     | Webpack 4  | SPFx (v1.16.1) project | 👤 manual (SPFx runtime)                           | ✅?🟡    |
| pptxgen.es.js     | Webpack 5  | SPFx (v1.19.1) project | 👤 manual (SPFx runtime)                           | ✅?🟡    |
| pptxgen.es.js     | Rollup 4   | Vite (v6) demo         | 🤖 `npm run release-test` (`vite.test.js`)         | ✅?🟡    |
| pptxgen.es.js     | Webworkers | worker_test demo       | 🤖 `npm run release-test` (`worker.test.js`)       | ✅?🟡    |
| pptxgen.cjs.js    | Node/CJS   | Node demo              | 🤖 `npm run release-test` (`node.test.js`)         | ✅?🟡    |
| pptxgen.bundle.js | Script     | Browser demo (desktop) | 🤖 `npm run release-test` (`browser.test.js`)      | ✅?🟡    |
| pptxgen.bundle.js | Script     | Browser demo (iOS)     | 👤 manual (no headless iOS runner)                 | ✅?🟡    |
