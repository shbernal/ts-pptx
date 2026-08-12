# Node stream demo

Generating a deck per request and streaming it straight to an HTTP response — no temp file,
nothing to clean up.

The showcase decks in [`../showcases`](../showcases/README.md) all end in `writeFile()`.
This one uses `pptx.toBytes()` instead, which is the shape a server actually needs.

## Run it

```bash
pnpm install                            # once, at the repository root
pnpm --dir demos/node run demo-stream
```

Then visit `http://localhost:3000/` to download the generated deck. Ctrl-C to stop.

Express is here only to have a server to attach the response to; it is not a ts-pptx
dependency.

## Notes

- ESM only. `require("@shbernal/ts-pptx")` is not supported.
- This is a showcase, not a test. Nothing here gates a commit.
