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
