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
- `services/api/**` – org policies + enforcement helpers + retention service

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

- `packages/security/crypto/envelope.js`
- `packages/security/crypto/kms/*` (providers)
- `services/api/storage/documentStorageRouter.js` (example routing + encryption wrapper)

The included `LocalKmsProvider` is meant for tests and single-node dev. Production deployments should supply AWS/GCP/Azure implementations.

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

## Encryption in Transit

Cloud services should enforce **TLS 1.3 minimum**.

- `services/api/policies/tls.js` exports `createTlsServerOptions()` (sets `minVersion: "TLSv1.3"`).

### Certificate pinning (enterprise option)

Certificate pinning is supported via a custom `checkServerIdentity` function:

- `services/api/policies/tls.js` exports `createPinnedCheckServerIdentity({ pins })`
- Pins are SHA-256 fingerprints of the server certificate (hex, with or without `:` separators).

## Data Residency

Residency is represented per organization:

- `us` | `eu` | `apac` | `custom`
- For `custom`, `allowedRegions` must be specified.

Helpers:

- `services/api/policies/dataResidency.js`
  - `getAllowedRegions()`
  - `resolvePrimaryStorageRegion()`
  - `resolveAiProcessingRegion()`
  - `assertRegionAllowed()`

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

- `services/api/retention/retentionService.js`
- `services/api/src/retention.ts` (Postgres retention sweep: archives `audit_log` → `audit_log_archive`, deletes old `document_versions`, purges soft-deleted `documents`)
- `services/api/migrations/0002_enterprise_security_policies.sql` (adds `audit_log_archive` + `document_legal_holds` + org policy columns)
- `services/api/src/routes/docs.ts` (`/docs/:docId/legal-hold` endpoints + soft-delete)

## Audit Logging

All org-level policy changes are logged:

- `services/api/org/orgPolicyService.js` emits:
  - `org.policy.encryption.updated`
  - `org.policy.dataResidency.updated`
  - `org.policy.retention.updated`

For tests/examples:

- `services/api/audit/auditLogger.js` provides `InMemoryAuditLogger`.

## Running tests

```bash
npm test
```
