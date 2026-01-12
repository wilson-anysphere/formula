# ADR-0003: Consolidate envelope encryption + KMS in `packages/security`

- **Status:** Accepted
- **Date:** 2026-01-12

## Context

Formula encrypts sensitive blobs at rest (e.g. `document_versions.data`) using envelope encryption:

- Generate a per-object **DEK** (Data Encryption Key) and encrypt the payload with AES-256-GCM.
- **Wrap** the DEK using a **KMS provider** (local dev/test or cloud KMS in production).

Historically the repo had two overlapping implementations:

1. A TypeScript envelope/KMS stack under `services/api/src/crypto/*`.
2. A JavaScript/ESM implementation under `packages/security/crypto/*`.

Maintaining two stacks increased security/operational risk (format drift, unclear rotation semantics, duplicated primitives).

## Decision

### 1) Canonical crypto lives in `packages/security/crypto`

`packages/security/crypto` is the single canonical implementation of:

- AES-256-GCM primitives
- envelope encryption (`packages/security/crypto/envelope.js`)
- KMS provider implementations (`packages/security/crypto/kms/*`)

The API (compiled to CommonJS) loads the ESM modules via `services/api/src/crypto/securityImport.ts`.

### 2) Envelope format + DB storage

The canonical envelope API is:

```ts
encryptEnvelope({ plaintext: Buffer, kmsProvider, encryptionContext? }) => Promise<{
  schemaVersion: number;
  wrappedDek: unknown;
  algorithm: string;
  iv: string;         // base64
  ciphertext: string; // base64
  tag: string;        // base64
}>
```

For `document_versions.data`, `services/api` stores envelope fields in dedicated Postgres columns:

- `data_ciphertext`, `data_iv`, `data_tag` (base64 text for pg-mem compatibility)
- `data_encrypted_dek` (wrapped DEK)
- `data_kms_provider`, `data_kms_key_id` (provider + key identifier/version for debugging)
- `data_aad` (JSON), bound into AES-GCM AAD via deterministic JSON encoding
- `data_envelope_version` (document_versions metadata schema)
  - `1`: legacy schema (HKDF local KMS; wrapped DEK stored as base64 bytes)
  - `2`: canonical schema (`packages/security` envelope + wrapped-key object stored as JSON)

Note: the AAD itself includes an `envelopeVersion` field that is intentionally stable so ciphertext remains valid even if the wrapped-DEK metadata format changes.

### 3) Canonical KMS provider interface

Providers implement:

```ts
type EncryptionContext = Record<string, unknown> | null;

interface EnvelopeKmsProvider {
  provider: string; // e.g. "local", "aws"
  wrapKey(args: { plaintextKey: Buffer; encryptionContext?: EncryptionContext }): Promise<unknown>;
  unwrapKey(args: { wrappedKey: unknown; encryptionContext?: EncryptionContext }): Promise<Buffer>;
}
```

Supported providers:

- **LocalKmsProvider** (`provider: "local"`)
  - Canonical implementation: `packages/security/crypto/kms/localKmsProvider.js`
  - Per-org KEK material is persisted + versioned in Postgres (`org_kms_local_state`) for dev/test.
  - Legacy support: `LOCAL_KMS_MASTER_KEY` is only required to decrypt envelope schema v1 rows.
- **AwsKmsProvider** (`provider: "aws"`) (optional)
  - Canonical implementation: `packages/security/crypto/kms/providers.js`
  - Loads `@aws-sdk/client-kms` lazily and includes monorepo-friendly resolution fallbacks.

Future GCP/Azure providers follow the same interface.

### 4) Key rotation semantics

Key rotation for the local provider is implemented in `services/api/src/crypto/kms.ts`:

1. Identify orgs due for rotation via `org_settings.key_rotation_days` and `org_settings.kms_key_rotated_at`.
2. Rotate the org’s local KEK version in `org_kms_local_state`.
3. **Re-wrap** stored DEKs in `document_versions.data_encrypted_dek` without re-encrypting ciphertext.

### 5) Backfill + migration tooling

See `docs/09-security-enterprise.md` for operational details, including:

- encrypting existing plaintext `document_versions.data` rows
- migrating legacy envelope schema v1 rows to schema v2 (DEK re-wrap only; ciphertext unchanged)

## Consequences

- `services/api` uses a single envelope/KMS stack (`packages/security/crypto`).
- Existing `document_versions` rows remain readable across envelope schema versions.
