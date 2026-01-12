# Extension package formats

Formula extensions are distributed as a single binary “package” blob downloaded from the Marketplace.

## v2 (current)

**Container:** deterministic TAR archive (no compression).

**Required entries:**

- `manifest.json` — canonical JSON (object keys sorted)
- `checksums.json` — canonical JSON with SHA-256 + size for each payload file
- `signature.json` — Ed25519 signature over canonical `{ manifest, checksums }`
- `files/<path>` — payload files (normalized POSIX-style paths)

**Manifest consistency:**

- `files/package.json` is required and must be JSON-identical to `manifest.json`.

**Path safety rules (enforced by the Marketplace and clients):**

- No absolute paths
- No `.` or `..` path segments
- No `:` in path segments (portable paths; avoids Windows drive/alternate stream semantics)
- No Windows reserved device names (`CON`, `NUL`, `COM1`, etc) and no trailing `.` / space in any path segment
- No Windows-invalid characters in any path segment (`<`, `>`, `:`, `"`, `|`, `?`, `*`)
- No case-insensitive duplicate paths (portable across Windows/macOS default filesystems)
- No symlinks (tar entries with non-file typeflags are rejected)

**Signature:**

- Algorithm: **Ed25519** (`signature.json.algorithm === "ed25519"`)
- Signed bytes: canonical JSON encoding of `{ manifest, checksums }`

## Browser/WebView installation model (Web + Desktop/Tauri)

The Web runtime and the Desktop/Tauri WebView runtime do **not** load extension module graphs directly from the network.
Instead they use the `WebExtensionManager` install flow (`packages/extension-marketplace/src/WebExtensionManager.ts`,
exported as `@formula/extension-marketplace`) which keeps the initial code load fully verified:

1. Download the `.fextpkg` bytes from the Marketplace (`/api/extensions/:id/download/:version`).
2. Verify the v2 package **client-side**:
   - parse the tar archive
   - validate `manifest.json`, `checksums.json`, `signature.json`
   - compute SHA-256 checksums (WebCrypto)
   - verify the Ed25519 signature:
     - **Web:** requires WebCrypto Ed25519 support (installs must fail when unavailable).
     - **Desktop/Tauri:** uses WebCrypto when available, but can fall back to a Rust-backed verifier via
       Tauri IPC (`verify_ed25519_signature`) on WebViews that don't support Ed25519 in WebCrypto
       (notably WKWebView/WebKitGTK).
3. Persist the verified package bytes + verification metadata in IndexedDB (keyed by `{id, version}`).
4. Extract the entrypoint (`manifest.browser`, falling back to `module`/`main`) from the archive, create a `blob:` URL
   for that module, and load it into `BrowserExtensionHost` (exported as `@formula/extension-host/browser`).

This same model is used in Desktop/Tauri builds: verified packages persist in IndexedDB
(`formula.webExtensions`) and are loaded from in-memory `blob:`/`data:` module URLs (no Node runtime, no extracted
extension directory on disk).

**Entrypoint requirement:** module loading from `blob:` URLs cannot resolve relative imports, so
`manifest.browser` should be a **single-file ESM bundle** (no `./` imports). Remote `http(s):` imports
are disallowed by the loader to avoid fetching unverified code at runtime.

## v1 (legacy / transition)

**Container:** gzipped JSON bundle containing base64-encoded files.

- Signature is **detached** and verified over the raw package bytes.
- The embedded `manifest` object must match the bundled `package.json` file contents.
- Supported only for backward compatibility while v2 rolls out.

## Tooling

- `pnpm extension:pack <dir> --out <file> [--private-key <pem>]`
- `pnpm extension:verify <file> --pubkey <pem> [--signature <base64>]` (v1 requires `--signature`)
- `pnpm extension:inspect <file>`
