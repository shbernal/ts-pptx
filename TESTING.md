# PptxGenJS Testing Guide

This document outlines how to manually test PptxGenJS across supported platforms and environments prior to release.

> ✅ Run these tests to ensure compatibility with major bundlers, runtimes, and front-end frameworks.

Config Notes

> ⚠️ Disable VPN on the server machine, otherwise, clients using the local IP address cannot connect.

Testing Steps

1. Run `npm test` for regression and schema validation.
2. Run `npm run build:dist` to refresh local package artifacts under `dist/`.
3. Run `npm run pack:check` to verify the npm package contents.
4. Run the manual sections below for browser, worker, Node, Vite, and PowerPoint visual checks.

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

### Build Coverage

Run `npm run ship` before manual browser checks when you need `demos/browser/js/pptxgen.bundle.js` to reflect the current source. Browser and worker demos are manual release checks.

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

Run the Node demo manually after building the package artifacts:

```bash
npm install --prefix demos/node
npm run build:dist
npm run copy:node
npm --prefix demos/node run demo
```

---

## ⚛️ Vite + TypeScript Tests

**Purpose:** Validate integration in modern front-end SPA toolchains (Vite, TypeScript, React-compatible).

Run the Vite build manually after building the package artifacts:

```bash
npm install --prefix demos/vite-demo
npm run build:dist
npm run copy:vite
npm --prefix demos/vite-demo run build
```

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
| pptxgen.es.js     | Rollup 4   | Vite (v6) demo         | 👤 manual (`npm --prefix demos/vite-demo run build`) | ✅?🟡    |
| pptxgen.es.js     | Webworkers | worker_test demo       | 👤 manual                                           | ✅?🟡    |
| pptxgen.cjs.js    | Node/CJS   | Node demo              | 👤 manual (`npm --prefix demos/node run demo`)      | ✅?🟡    |
| pptxgen.bundle.js | Script     | Browser demo (desktop) | 👤 manual                                           | ✅?🟡    |
| pptxgen.bundle.js | Script     | Browser demo (iOS)     | 👤 manual (no headless iOS runner)                 | ✅?🟡    |
