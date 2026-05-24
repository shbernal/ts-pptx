# OOXML-Validator CLI

Schema validation for `.pptx` files using Microsoft's
[`OpenXmlValidator`](https://github.com/dotnet/Open-XML-SDK) wrapped by
[`mikeebowen/OOXML-Validator`](https://github.com/mikeebowen/OOXML-Validator).

This catches Microsoft-implementation-specific issues that pure schema
validation (raw OOXML XSDs) misses — most "needs repair" causes are in
this category.

## Install

The binary is **not committed** to this repo (it's ~125 MB extracted).
Install it locally with:

```bash
./tools/ooxml-validator/install.sh
```

The script downloads the platform-appropriate self-contained `.NET`
binary from upstream GitHub Releases into `bin/` and verifies it runs.

## Versions

The pinned upstream version is in `version.json`. To check whether a
newer release is available, run:

```bash
node tools/ooxml-validator/check-updates.js
```

To upgrade, bump `version.json` and re-run `install.sh`.

## Layout

- `version.json`     — pinned upstream version (committed)
- `install.sh`       — installer script (committed)
- `check-updates.js` — upstream-update checker (committed)
- `bin/`             — installed binary (gitignored)

## License

The `OOXMLValidatorCLI` binary is MIT-licensed by its author Mike Bowen,
wrapping Microsoft's MIT-licensed `Open-XML-SDK`. The binary itself is
not redistributed by this repository — it is fetched from upstream
GitHub Releases at install time.
