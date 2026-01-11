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

## Package scanning

After a successful publish, the marketplace:

1. Inserts a `package_scans` row with `status=pending`
2. Performs an in-process scan and updates the row to `passed` or `failed`

The scan is intentionally conservative and includes:

- Defense-in-depth manifest validation + entrypoint validation
- SBOM-like file inventory (`files_json`, per-file sha256 + size) for both v1 and v2 packages
- Basic JS heuristics (e.g. `child_process`, `eval`, `new Function`, hex-escape obfuscation)

Downloads are blocked when scan status is `failed`. Operators can tighten policy via
`MARKETPLACE_REQUIRE_SCAN_PASSED=1`.

## Download provenance headers

`GET /api/extensions/:id/download/:version` includes:

- `X-Package-Sha256`
- `X-Package-Signature`
- `X-Package-Scan-Status` (`pending|passed|failed|unknown`)
- `X-Package-Files-Sha256` (sha256 of `extension_versions.files_json`)
- `X-Package-Format-Version` (1 or 2)
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

- `POST /api/admin/publishers/:publisher/rotate-token`
- `POST /api/admin/publishers/:publisher/rotate-key`
- `POST /api/admin/publishers/:publisher/revoke`
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
