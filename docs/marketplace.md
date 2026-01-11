# Marketplace HTTP API

The marketplace service is a small HTTP API used by the desktop client and the extension publisher CLI.

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

## Endpoints

### `POST /api/publish-bin` (binary publish)

Upload a raw extension package (no JSON/base64 wrapper).

**Headers**

- `Authorization: Bearer <publisherToken>` (required)
- `Content-Type: application/vnd.formula.extension-package` (required)
- `X-Package-Signature: <base64>` (required for **v1** packages; optional for v2)

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

**Body**

```json
{
  "packageBase64": "<base64-encoded package bytes>",
  "signatureBase64": "<base64 detached signature>" // required for v1 packages
}
```

### `GET /api/extensions/:id` (extension metadata)

Returns extension metadata, including versions and the publisher public key.

**Response headers**

- `ETag`: changes when the extension metadata changes.

### `GET /api/extensions/:id/download/:version` (download package)

Downloads raw package bytes.

**Response headers**

- `Content-Type: application/vnd.formula.extension-package`
- `ETag`: the package sha256 (used for conditional requests)
- `X-Package-Sha256`: hex sha256 of the response body (**clients must verify**)
- `X-Package-Signature`: base64 signature (detached for v1; for v2 this matches the in-package signature payload)
- `X-Package-Format-Version`: `1` or `2`
- `X-Publisher`: publisher id

**Client integrity requirement**

Clients should compute sha256 over the downloaded bytes and reject the download if it does not match `X-Package-Sha256`.

