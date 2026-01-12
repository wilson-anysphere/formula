# Marketplace HTTP API

The marketplace service is a small HTTP API used by the web + desktop clients and the extension publisher CLI.

## Base URL conventions

All marketplace routes in this document are under the `/api/*` prefix (for example: `/api/search`).

Depending on the caller, the configurable “base URL” differs:

- **Browser/WebView clients** (`MarketplaceClient` / `WebExtensionManager`) conceptually use a base URL that includes the API
  prefix, e.g. `https://marketplace.formula.app/api`.
  - For convenience, origin-only values (e.g. `https://marketplace.formula.app`) are accepted and normalized to `.../api`.
- **Publisher tooling** (`tools/extension-publisher`) takes the marketplace **origin** and appends `/api/...` internally,
  e.g. `https://marketplace.formula.app`. (Passing a URL that ends with `/api` is also accepted and will be normalized.)

## CORS (browser clients)

Read-only endpoints are CORS-enabled so browser/WebView clients can call the marketplace cross-origin:

- `GET /api/search`
- `GET /api/extensions/:id`
- `GET /api/extensions/:id/download/:version`

These responses include:

- `Access-Control-Allow-Origin: *`
- `Access-Control-Expose-Headers: ETag, X-Package-Sha256, X-Package-Signature, X-Package-Scan-Status, X-Package-Files-Sha256, X-Package-Format-Version, X-Package-Published-At, X-Publisher, X-Publisher-Key-Id, ...`

Mutation endpoints (publish/admin) are intentionally **not** CORS-enabled.

## Content types

- Extension package bytes: `application/vnd.formula.extension-package`
- JSON: `application/json; charset=utf-8`

## Authentication

Publisher and admin endpoints use bearer tokens:

```
Authorization: Bearer <token>
```

## Caching (ETag)

`GET` endpoints that return immutable or rarely-changing resources include an `ETag` header.

- Clients may send `If-None-Match: <etag>` to receive `304 Not Modified` with an empty body.
- For package downloads, the `ETag` is the package sha256.
- Responses are served with `Cache-Control: public, max-age=0, must-revalidate` so intermediaries may cache
  but must revalidate via `ETag` before reuse.
- `304` package download responses still include the same `X-*` metadata headers as `200` responses (sha256,
  signature, scan status, files sha256, format version, publisher, key id, etc) so clients can update caches without
  re-downloading bytes.

## Endpoints

### `POST /api/publish-bin` (binary publish)

Upload a raw extension package (no JSON/base64 wrapper).

**Headers**

- `Authorization: Bearer <publisherToken>` (required)
- `Content-Type: application/vnd.formula.extension-package` (required)
- `X-Package-Signature: <base64>` (required for **v1** packages; optional for v2)
- `X-Package-Sha256: <hex>` (optional; if present, must be a 64-character hex sha256 and the server rejects uploads whose bytes do not match)

**Body**

- Raw package bytes.

**Notes**

- Request bodies are capped at 20MB; larger uploads are rejected.
- v1 packages use a detached signature (provided via `X-Package-Signature`).
- v2 packages embed the signature in-package; `X-Package-Signature` is ignored/optional.

**Response**

`200 OK`

```json
{ "id": "publisher.name", "version": "1.2.3" }
```

### `POST /api/publish` (legacy JSON publish)

Backward compatible endpoint for older clients.

**Caching**

Authenticated/mutation endpoints are served with `Cache-Control: no-store`.

**Body**

```json
{
  "packageBase64": "<base64-encoded package bytes>",
  "signatureBase64": "<base64 detached signature>" // required for v1 packages
}
```

### `GET /api/search` (search)

Search the marketplace for extensions.

**Query parameters**

- `q` (string): free-text query (matches id, name, publisher, tags, etc)
- `category` (string): filter by category
- `tag` (string): filter by tag
- `verified` (boolean): filter to verified publishers/extensions
- `featured` (boolean): filter to featured extensions
- `sort` (string): e.g. `relevance|updated|downloads`
- `limit` (number): max results to return (server clamps)
- `offset` (number): legacy pagination offset
- `cursor` (string): opaque cursor for pagination (preferred over `offset`)

**Response**

`200 OK`

```json
{
  "total": 123,
  "results": [/* extension summaries */],
  "nextCursor": "..."
}
```

**Notes**

- Search results are *moderated* and may omit extensions that are:
  - blocked
  - malicious
  - deprecated
  - published by a revoked publisher
- Version visibility is also affected by scan status/yanking (e.g. failed scans are hidden).
- Clients that need authoritative security metadata should call `GET /api/extensions/:id` for a full record.

### `GET /api/extensions/:id` (extension metadata)

Returns extension metadata, including versions and the publisher signing key(s).

**Security flags**

The extension record includes security metadata that clients **must** honor:

- `blocked: boolean` — extension is blocked and must not be installed.
- `malicious: boolean` — extension is marked as malicious and must not be installed.
- `deprecated: boolean` — extension is deprecated; clients may allow install but should warn.
- `publisherRevoked?: boolean` — the publisher has been revoked; clients must not install.

Each entry in `versions` may also include:

- `scanStatus: string` — package scan status (for UI/display and client policy decisions).

**Response headers**

- `ETag`: changes when the extension metadata changes.

**Response body**

- `publisherPublicKeyPem`: the publisher's *primary* public key (backward compatibility for older clients).
  - When all publisher keys are revoked, this is `null` and installs must fail (by design).
- `publisherKeys`: array of known publisher keys:

  ```ts
  publisherKeys: Array<{ id: string; publicKeyPem: string; revoked: boolean }>;
  ```

### `GET /api/extensions/:id/download/:version` (download package)

Downloads raw package bytes.

**Response headers**

- `Content-Type: application/vnd.formula.extension-package`
- `ETag`: the package sha256 (used for conditional requests)
- `X-Package-Sha256`: hex sha256 of the response body (**clients must verify**)
- `X-Package-Signature`: base64 signature (detached for v1; for v2 this matches the in-package signature payload)
- `X-Package-Scan-Status`: `pending|passed|failed|unknown` (may be used by clients to enforce install policy)
- `X-Package-Files-Sha256`: sha256 of the server-side file inventory for the package version (used as an optional cross-check
  against the client-verified file list)
- `X-Package-Format-Version`: `1` or `2`
- `X-Package-Published-At`: publish timestamp for the package version (optional)
- `X-Publisher`: publisher id
- `X-Publisher-Key-Id`: key id (sha256 fingerprint) identifying which publisher key signed this version

**Client integrity requirement**

Clients should compute sha256 over the downloaded bytes and reject the download if it does not match `X-Package-Sha256`.

Clients must also verify the package signature using the publisher public key(s) returned by
`GET /api/extensions/:id` (`publisherKeys` / `publisherPublicKeyPem`):

- v1: `X-Package-Signature` is the detached signature over the raw package bytes.
- v2: the signature is embedded in `signature.json` (the header is redundant/back-compat).
- Web runtimes require WebCrypto Ed25519 support; Desktop/Tauri can fall back to a Rust-backed verifier via
  Tauri IPC (`verify_ed25519_signature`) when the embedded WebView does not support Ed25519 in WebCrypto
  (notably WKWebView/WebKitGTK).

### `POST /api/publishers/:publisher/keys/:id/revoke` (admin)

Revoke a publisher signing key.

**Notes**

- Admin-only (requires the marketplace admin bearer token).
- Revoked keys are excluded from publish-time verification and clients ignore them during install.
