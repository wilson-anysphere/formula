# Extension package formats

Formula extensions are distributed as a single binary “package” blob downloaded from the Marketplace.

## v2 (current)

**Container:** deterministic TAR archive (no compression).

**Required entries:**

- `manifest.json` — canonical JSON (object keys sorted)
- `checksums.json` — canonical JSON with SHA-256 + size for each payload file
- `signature.json` — Ed25519 signature over canonical `{ manifest, checksums }`
- `files/<path>` — payload files (normalized POSIX-style paths)

**Path safety rules (enforced by the Marketplace and clients):**

- No absolute paths
- No `.` or `..` path segments
- No symlinks (tar entries with non-file typeflags are rejected)

**Signature:**

- Algorithm: **Ed25519** (`signature.json.algorithm === "ed25519"`)
- Signed bytes: canonical JSON encoding of `{ manifest, checksums }`

## v1 (legacy / transition)

**Container:** gzipped JSON bundle containing base64-encoded files.

- Signature is **detached** and verified over the raw package bytes.
- Supported only for backward compatibility while v2 rolls out.

## Tooling

- `pnpm extension:pack <dir> --out <file> [--private-key <pem>]`
- `pnpm extension:verify <file> --pubkey <pem>`
- `pnpm extension:inspect <file>`

