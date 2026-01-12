# Marketplace service

The marketplace service stores and serves Formula extension packages. It is designed to support
enterprise governance requirements:

- Signature verification (publisher keys, rotation, revocation)
- Post-publish security scanning with persisted results
- Provenance headers on download
- Admin endpoints for publisher and scan governance

## Configuration

The server is created via `createMarketplaceServer({ dataDir, adminToken })` in
`services/marketplace/src/server.js`.

The `MarketplaceStore` (created internally) reads these environment variables:

- `MARKETPLACE_SCAN_ALLOWLIST` – CSV of scan finding IDs to ignore (e.g. `js.eval,js.child_process`)
- `MARKETPLACE_REQUIRE_SCAN_PASSED=1` – if set, downloads and `latestVersion` selection require the
  scan status to be `passed` (pending/unknown are blocked).
- `MARKETPLACE_EXTERNAL_SCANNER_ENABLED=1` – enable an optional external scanner hook.
- `MARKETPLACE_EXTERNAL_SCANNER_CMD` – scanner command (e.g. `clamdscan` or `clamscan`).
- `MARKETPLACE_EXTERNAL_SCANNER_ARGS` – JSON array of arguments (optional).
- `MARKETPLACE_EXTERNAL_SCANNER_TIMEOUT_MS` – timeout (ms) for external scans (optional).

## Local development (run a marketplace server)

There is intentionally no production CLI baked into this repo yet; tests and local development start the server by
calling `createMarketplaceServer` directly.

From the repo root:

```bash
node --input-type=commonjs - <<'NODE'
const { createMarketplaceServer } = require("./services/marketplace/src/server.js");

(async () => {
  const { server } = await createMarketplaceServer({
    dataDir: "./.marketplace-data",
    adminToken: "admin-secret",
  });

  server.listen(8787, "127.0.0.1", () => {
    console.log("Marketplace listening on http://127.0.0.1:8787");
  });
})();
NODE
```

Notes:

- The API origin is `http://127.0.0.1:8787`.
- Marketplace API routes are under `/api/*` (e.g. `/api/search`, `/api/extensions/:id`, …).
- Desktop/Tauri uses a base URL whose routes are under `/api` (e.g. `http://127.0.0.1:8787/api`).
  - For convenience, it also accepts an origin (`http://127.0.0.1:8787`) and normalizes it to `/api`.

## Publisher bootstrap (local dev / tests)

To publish extensions to a fresh marketplace instance you need to register a publisher (token + signing public key).
In production this is handled by operator tooling; in local dev and tests the server exposes:

- `POST /api/publishers/register` (requires the configured `adminToken`)

Example (from repo root):

```bash
# 1) Generate an Ed25519 keypair (write to publisher-private.pem + publisher-public.pem)
node --input-type=commonjs - <<'NODE'
const crypto = require("node:crypto");
const fs = require("node:fs");

const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
fs.writeFileSync("publisher-private.pem", privateKey.export({ type: "pkcs8", format: "pem" }));
fs.writeFileSync("publisher-public.pem", publicKey.export({ type: "spki", format: "pem" }));
console.log("Wrote publisher-private.pem + publisher-public.pem");
NODE

# 2) Register a publisher (admin-only)
node --input-type=commonjs - <<'NODE'
const fs = require("node:fs/promises");

(async () => {
  const publicKeyPem = await fs.readFile("publisher-public.pem", "utf8");
  const res = await fetch("http://127.0.0.1:8787/api/publishers/register", {
    method: "POST",
    headers: {
      Authorization: "Bearer admin-secret",
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      publisher: "mycompany",
      token: "publisher-token",
      publicKeyPem,
      verified: true,
    }),
  });
  console.log(res.status, await res.text());
})();
NODE

# 3) Publish an extension directory (origin without /api; `/api` is also accepted and normalized)
node tools/extension-publisher/src/cli.js publish extensions/my-extension \\
  --marketplace http://127.0.0.1:8787 \\
  --token publisher-token \\
  --private-key ./publisher-private.pem
```

## Package scanning

After a successful publish, the marketplace:

1. Inserts a `package_scans` row with `status=pending`
2. Performs an in-process scan and updates the row to `passed` or `failed`

The scan is intentionally conservative and includes:

- Defense-in-depth manifest validation + entrypoint validation
- SBOM-like file inventory (`files_json`, per-file sha256 + size) for both v1 and v2 packages
- Native executable detection (ELF / PE / Mach-O signatures), shebang scripts, and NUL bytes in text-like files
- Basic JS heuristics (e.g. `child_process`, `eval`, `new Function`, hex-escape obfuscation)
- Optional external scanner hook (feature-flagged)

Downloads are blocked when scan status is `failed`. Operators can tighten policy via
`MARKETPLACE_REQUIRE_SCAN_PASSED=1`.

## Download provenance headers

`GET /api/extensions/:id/download/:version` includes:

- `X-Package-Sha256`
- `X-Package-Signature`
- `X-Package-Scan-Status` (`pending|passed|failed|unknown`)
- `X-Package-Files-Sha256` (sha256 of `extension_versions.files_json`)
- `X-Package-Format-Version` (1 or 2)
- `X-Package-Published-At`
- `X-Publisher`
- `X-Publisher-Key-Id` (when known)

## Extension metadata

`GET /api/extensions/:id` returns extension metadata including a `versions[]` array. Each version entry
includes:

- `version`, `sha256`, `uploadedAt`, `yanked`
- `scanStatus` (from `package_scans`)
- `signingKeyId` (publisher signing key id used for that version, when known)
- `formatVersion` (1 or 2)

Extensions from revoked publishers are hidden from public `search`/`getExtension`/download routes.

## Publisher governance

Publisher keys are stored in `publisher_keys` (migration `004_publisher_keys.sql`).

Publishing enforces:

- Publisher must not be revoked (`publishers.revoked`)
- Package signatures must verify against any non-revoked key in `publisher_keys`

## Admin endpoints

Admin endpoints require `adminToken` to be configured when starting the server.

### Publisher management

- `POST /api/publishers/register` (bootstrap a publisher; local dev/tests only)
- `POST /api/admin/publishers/:publisher/rotate-token`
- `POST /api/admin/publishers/:publisher/rotate-key`
- `POST /api/admin/publishers/:publisher/revoke`
- `PATCH /api/admin/publishers/:publisher` (set `{ "verified": true|false }`)
- `GET /api/admin/publishers/:publisher` (publisher record + key history)
- `POST /api/publishers/:publisher/keys/:id/revoke` (revoke a specific signing key)

### Scan management

- `GET /api/admin/extensions/:id/versions/:version/scan`
- `POST /api/admin/extensions/:id/versions/:version/scan` (force rescan)
- `GET /api/admin/scans` (list scans; filters: `status`, `publisher`, `extensionId`, pagination via
  `limit`/`offset`)
- `POST /api/admin/scans/rescan-pending` (bulk rescan; body: `{ "limit": 50 }`)

### Extension visibility for admins

- `GET /api/admin/extensions/:id` returns extension metadata even when hidden (blocked/malicious or
  publisher revoked).
