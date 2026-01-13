# Security (Enterprise Readiness)

This repository includes a **reference implementation** of enterprise security primitives used by Formula (desktop + cloud). It focuses on:

- Encryption-at-rest (desktop + cloud)
- Key management + rotation (desktop keychain + cloud KMS abstraction)
- Encryption-in-transit defaults (TLS 1.3 minimum) + optional certificate pinning
- Data residency policy representation + enforcement helpers
- Retention policy enforcement with legal hold override
- Audit logging for policy changes

The code lives in:

- `packages/security/crypto/**` – AES-256-GCM primitives, envelope encryption, KMS + keychain abstractions
- `apps/desktop/src-tauri/src/storage/**` – desktop encryption hooks (testable JS reference + Rust placeholder)
- `services/api/src/**` – Fastify API implementation (org settings/policies, retention sweep, audit logging)

## Encryption at Rest

### Desktop (local documents / SQLite)

The desktop store is encrypted with **AES-256-GCM** using a **master key** stored in the OS keychain (or an in-memory mock in tests).

Implementation:

- `apps/desktop/src-tauri/src/storage/encryptedDocumentStore.js`
- `packages/versioning/src/store/sqliteVersionStore.js` (optional encryption for on-disk SQLite via AES-256-GCM + keychain-backed `KeyRing`)
- Key material stored via a `KeychainProvider` (`packages/security/crypto/keychain/**`)
- Encryption uses a versioned `KeyRing` so key rotation can be performed without losing access to existing data.

Enable/disable encryption performs an on-disk migration:

1. Enable: plaintext → ciphertext; keyring written to keychain
2. Disable: ciphertext → plaintext; keyring optionally removed from keychain

### Cloud (documents + backups)

Cloud storage uses **envelope encryption**:

1. Generate a random **DEK** (AES-256 key) per object
2. Encrypt the object with AES-256-GCM
3. Wrap the DEK using a configured **KMS provider**

Implementation:

- `services/api/src/crypto/envelope.ts`
- `services/api/src/crypto/keyring.ts`
- `services/api/src/crypto/kms/*` (providers)
- `services/api/src/db/documentVersions.ts` (applies envelope encryption to DB blobs)

The included `LocalKmsProvider` is meant for tests and single-node dev. Production deployments should supply AWS/GCP/Azure implementations.

`services/api` also uses the same canonical envelope encryption primitives to encrypt
`document_versions.data` in Postgres (see `services/api/src/db/documentVersions.ts`). Historical
deployments may have legacy rows in an older envelope schema; the API can still decrypt those, and
provides a migration script (`npm -C services/api run versions:migrate-legacy`) to re-wrap DEKs into
the canonical schema without re-encrypting ciphertext. `LOCAL_KMS_MASTER_KEY` is only needed for
decrypting/migrating those legacy rows.

AWS KMS support is optional and loaded lazily. If you configure `kms_provider='aws'`, ensure
`@aws-sdk/client-kms` is installed in the service/runtime image that uses it (e.g. `services/api`).

### Sync server (Yjs persistence)

`services/sync-server` supports **encryption at rest** for persisted Yjs state in both supported persistence backends:

- `file`: append-only `.yjs` update logs on disk
- `leveldb`: encrypted LevelDB *values* (updates, state vectors, metadata)

Enable:

- `SYNC_SERVER_PERSISTENCE_ENCRYPTION=keyring`
- Provide key material via **one** of:
  - `SYNC_SERVER_ENCRYPTION_KEYRING_JSON` (KeyRing JSON string), or
  - `SYNC_SERVER_ENCRYPTION_KEYRING_PATH` (path to a JSON file containing KeyRing JSON)
  - (Optional shorthand) `SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64` (base64 32-byte key; creates a single-version keyring — use KeyRing JSON for rotation)

Generate / rotate / validate KeyRing material:

```bash
pnpm -C services/sync-server -s keyring:generate --out keyring.json
pnpm -C services/sync-server -s keyring:validate --in keyring.json
pnpm -C services/sync-server -s keyring:rotate --in keyring.json --out keyring.json
```

The keyring JSON contains **secret key material**. Store it in your secret manager or lock down file permissions (e.g. `chmod 600 keyring.json`).

For production images / built output, you can run the compiled entrypoint:

```bash
node services/sync-server/dist/keyring-cli.js generate --out keyring.json
```

#### File persistence (`SYNC_SERVER_PERSISTENCE_BACKEND=file`)

When enabled, plaintext `.yjs` files in `SYNC_SERVER_DATA_DIR` are migrated to the encrypted format on startup (write temp + rename, idempotent).

On-disk format is a small, append-friendly container:

- Header: `FMLYJS01` (8 bytes) + 1 byte flags (`bit0 = encrypted`) + 3 reserved bytes
- Records: `[u32be recordLen][recordBytes...]`
  - Encrypted record bytes: `[u32be keyVersion][12B iv][16B tag][ciphertext...]`

Records use AES-256-GCM with an AAD context that includes a stable scope + schema version + `doc = sha256(docName)` so ciphertext cannot be swapped across documents.

Key rotation is handled by replacing the KeyRing JSON with a new `currentVersion` + key while retaining older key versions so existing records remain decryptable until compaction rewrites them.

#### LevelDB persistence (`SYNC_SERVER_PERSISTENCE_BACKEND=leveldb`)

When enabled, LevelDB persistence encrypts all persisted *values* (Yjs updates, state vectors, and metadata) using AES-256-GCM. LevelDB keys remain plaintext; enable `SYNC_SERVER_LEVELDB_DOCNAME_HASHING=1` to avoid writing raw doc ids into the LevelDB keyspace.

Migration strictness:

- `SYNC_SERVER_PERSISTENCE_ENCRYPTION_STRICT=1|0`
  - Default: `1` in `NODE_ENV=production`, `0` in dev/test.
  - Strict (`1`): rejects legacy plaintext values.
  - Non-strict (`0`): allows reading legacy plaintext values for migration/backcompat; new writes are encrypted.

### Sync server (range restriction enforcement)

The collaboration client enforces document roles + range restrictions locally, but a malicious client can bypass
UI checks by crafting raw Yjs updates. `services/sync-server` can enforce spreadsheet **cell/range edit
restrictions server-side** when they are present in the auth context (via JWT claims or token introspection
responses).

Enable enforcement:

- `SYNC_SERVER_ENFORCE_RANGE_RESTRICTIONS=1` (default: `true` in `NODE_ENV=production`, `false` otherwise)

When enabled and `rangeRestrictions` are provided, the server reads:

- `role` (document role)
- optional `rangeRestrictions` (JWT `rangeRestrictions` claim, or `rangeRestrictions` field from token
  introspection responses; compatible with `packages/collab/permissions.normalizeRestriction`)

If an incoming Yjs update attempts to modify a cell where `canEdit` is false, the server:

- closes the WebSocket with policy violation (`1008`)
- logs `permission_violation` (includes `docName`, `userId`, `role`)
- does not apply the update

The validator is best-effort and intentionally **fails closed** if it cannot confidently determine which cell keys
were affected.

Audit hardening: for touched cells, the server rewrites `modifiedBy` to the authenticated `userId` so clients cannot
spoof edit attribution.

Example `rangeRestrictions` value (JWT claim or token introspection response field):

```json
{
  "rangeRestrictions": [
    {
      "range": { "sheetId": "Sheet1", "startRow": 0, "endRow": 0, "startCol": 0, "endCol": 0 },
      "editAllowlist": ["user-id"]
    }
  ]
}
```

## Encryption in Transit

Cloud services should enforce **TLS 1.3 minimum**.

In production, TLS is typically terminated at the load balancer / edge proxy; the API expects to run behind an HTTPS-only ingress.

### Certificate pinning (enterprise option)

Certificate pinning settings are stored per org (see `org_settings.certificate_pinning_enabled` and `org_settings.certificate_pins`) and validated by the API when admins update org settings (`services/api/src/routes/orgs.ts`).

## Data Residency

Residency is represented per organization:

- `us` | `eu` | `apac` | `custom`
- For `custom`, `allowedRegions` must be specified.

Helpers:

- Org residency settings are stored in `org_settings` and validated by the API in `services/api/src/routes/orgs.ts`.

## Retention + Legal Hold

Retention policies apply to:

- Versions/snapshots
- Audit logs (archive + delete from hot storage)
- Deleted documents (time-based purge)

Legal holds override deletion when enabled in policy (`legalHoldOverridesRetention`).

When a soft-deleted document is **hard-deleted** (purged) by the retention sweep, the API can optionally trigger deletion of any persisted CRDT/Yjs state in `services/sync-server` via its internal purge endpoint (`DELETE /internal/docs/:docId`). This closes the gap where database records are removed but sync persistence could otherwise remain on disk indefinitely.

To enable sync-server state purge:

- `services/api`:
  - `SYNC_SERVER_INTERNAL_URL` (base HTTP URL for sync-server, e.g. `http://sync-server:1234`)
  - `SYNC_SERVER_INTERNAL_ADMIN_TOKEN` (sent as `x-internal-admin-token`)
- `services/sync-server`:
  - `SYNC_SERVER_INTERNAL_ADMIN_TOKEN` (must match the API's `SYNC_SERVER_INTERNAL_ADMIN_TOKEN`) to authorize `DELETE /internal/docs/:docId` requests.

Implementation:

- `services/api/src/retention.ts` (Postgres retention sweep: archives `audit_log` → `audit_log_archive`, deletes old `document_versions`, purges soft-deleted `documents`)
- `services/api/migrations/0002_enterprise_security_policies.sql` (adds `audit_log_archive` + `document_legal_holds` + org policy columns)
- `services/api/src/routes/docs.ts` (`/docs/:docId/legal-hold` endpoints + soft-delete)

## Audit Logging

All org-level policy changes are logged:

- `services/api/src/routes/orgs.ts` emits:
  - `org.policy.encryption.updated`
  - `org.policy.dataResidency.updated`
  - `org.policy.retention.updated`

## Running tests

```bash
npm test
```
