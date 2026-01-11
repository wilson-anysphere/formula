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
- No symlinks (tar entries with non-file typeflags are rejected)

**Signature:**

- Algorithm: **Ed25519** (`signature.json.algorithm === "ed25519"`)
- Signed bytes: canonical JSON encoding of `{ manifest, checksums }`

## v1 (legacy / transition)

**Container:** gzipped JSON bundle containing base64-encoded files.

- Signature is **detached** and verified over the raw package bytes.
- The embedded `manifest` object must match the bundled `package.json` file contents.
- Supported only for backward compatibility while v2 rolls out.

## Tooling

- `pnpm extension:pack <dir> --out <file> [--private-key <pem>]`
- `pnpm extension:verify <file> --pubkey <pem> [--signature <base64>]` (v1 requires `--signature`)
- `pnpm extension:inspect <file>`
