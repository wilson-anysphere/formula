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

## Browser installation model (web runtime)

The web app does **not** load extension module graphs directly from the network. Instead it uses an
install flow that keeps the initial code load fully verified:

1. Download the `.fextpkg` bytes from the Marketplace (`/api/extensions/:id/download/:version`).
2. Verify the v2 package **client-side**:
   - parse the tar archive
   - validate `manifest.json`, `checksums.json`, `signature.json`
   - compute SHA-256 checksums (WebCrypto)
   - verify the Ed25519 signature (WebCrypto Ed25519)
3. Persist the verified package bytes + verification metadata in IndexedDB (keyed by `{id, version}`).
4. Extract the entrypoint (`manifest.browser`, falling back to `module`/`main`) from the archive,
   create a `blob:` URL for that module, and load it into `BrowserExtensionHost`.

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
