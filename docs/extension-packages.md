# Extension packages (verification + atomic install)

Formula extensions are distributed as *extension packages* (`.fextpkg`). Packages contain the
extension's `package.json` manifest and the full extension file tree.

Because extension files are executable code, **verification must be tied to installation**: never
persist package contents (disk extraction or IndexedDB installs) unless the package has been successfully verified first.

## Formats

### v1 (legacy)

- Container: gzipped JSON bundle
- Signature: **detached** (`signatureBase64` delivered out-of-band, e.g. via HTTP header)

### v2 (current)

- Container: deterministic tar archive
- Signature: **embedded** (`signature.json`), over `manifest.json` + `checksums.json`
- Integrity: per-file SHA-256 checksums in `checksums.json`

## Browser/WebView install flow (Web + Desktop/Tauri)

The production Web + Desktop/Tauri runtime installs extensions without Node/`fs`:

- `WebExtensionManager` (`packages/extension-marketplace/src/WebExtensionManager.ts`, exported as `@formula/extension-marketplace`)
  downloads `.fextpkg` bytes, verifies them client-side (SHA-256 + Ed25519), and persists verified packages in IndexedDB
  (`formula.webExtensions`).
  - Note: on Desktop/Tauri, Ed25519 verification uses WebCrypto when available, but can fall back to a Rust-backed verifier
    via Tauri IPC on WebViews that don't support Ed25519 in WebCrypto (e.g. WKWebView/WebKitGTK).
- Extensions are loaded into `BrowserExtensionHost` (exported as `@formula/extension-host/browser`) from in-memory
  `blob:`/`data:` module URLs (no extracted extension directory on disk).

For details, see:

- [`docs/10-extensibility.md`](./10-extensibility.md) (Desktop/Tauri runtime + CSP/sandbox model)
- [`docs/extension-package-format.md`](./extension-package-format.md) (package format + browser install model)

## Node filesystem install flow (legacy/test harness)

The repository also includes a Node-only helper for verified, atomic extraction to disk. This is used by
Node-based tooling/tests and legacy installers; it is **not used** by the Desktop/Tauri WebView runtime.

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

## Why verification must be tied to installation

If installation happens before verification:

- A tampered download can write attacker-controlled files to disk.
- Failed installs/updates can leave partially-written extension directories behind.
- Updates can accidentally delete a working install before discovering the new package is invalid.

Tying verification to atomic extraction (or to IndexedDB persistence in browser/WebView runtimes) prevents these
footguns and ensures installs are all-or-nothing.
