# Extension packages (verification + atomic install)

Formula extensions are distributed as *extension packages* (`.fextpkg`). Packages contain the
extension's `package.json` manifest and the full extension file tree.

Because extension files are executable code, **verification must be tied to extraction**: never
write package contents to disk unless the package has been successfully verified first.

## Formats

### v1 (legacy)

- Container: gzipped JSON bundle
- Signature: **detached** (`signatureBase64` delivered out-of-band, e.g. via HTTP header)

### v2 (current)

- Container: deterministic tar archive
- Signature: **embedded** (`signature.json`), over `manifest.json` + `checksums.json`
- Integrity: per-file SHA-256 checksums in `checksums.json`

## Recommended install flow

Use the high-level helper in `shared/extension-package`:

```js
const { verifyAndExtractExtensionPackage } = require("../shared/extension-package");

await verifyAndExtractExtensionPackage(packageBytes, destDir, {
  publicKeyPem,                 // required
  signatureBase64,              // required for v1, optional for v2
  expectedId: "publisher.name", // optional safety check
  expectedVersion: "1.2.3",     // optional safety check
});
```

This function:

1. Detects the package format (or uses `formatVersion` if provided).
2. **Verifies before writing any files**:
   - v1: verifies the detached signature against the raw package bytes first.
   - v2: verifies the embedded signature + checksums before extracting.
3. Extracts into a **temporary staging directory** under the same parent directory as `destDir`.
4. Atomically "commits" the install by renaming the staging directory into `destDir` (and safely
   replacing an existing install on updates).
5. Cleans up the staging directory on any failure.

Callers should only update any persistent state (e.g. an "installed extensions" state file) **after**
`verifyAndExtractExtensionPackage` succeeds.

## Why verification must be tied to extraction

If extraction happens before verification:

- A tampered download can write attacker-controlled files to disk.
- Failed installs/updates can leave partially-written extension directories behind.
- Updates can accidentally delete a working install before discovering the new package is invalid.

Tying verification to atomic extraction prevents these footguns and ensures installs are all-or-nothing.

