# Marketplace publisher signing keys

## Why this exists

Extensions published to the Formula marketplace are signed by the publisher. Clients verify the
signature before installing.

Historically the marketplace stored **one** public key per publisher (`publishers.public_key_pem`).
If a publisher rotated their signing key, clients would only see the new key and would fail to
verify (and therefore install) **older extension versions** that were signed with the previous key.

This document describes the **key history** + **per-version key capture** model that enables safe
publisher key rotation.

---

## Data model

### `publisher_keys`

The marketplace stores all known (non-private) signing keys for a publisher:

| column | meaning |
|---|---|
| `id` | key id (SHA-256 fingerprint of the public key SPKI bytes) |
| `publisher` | owning publisher |
| `public_key_pem` | Ed25519 public key in PEM/SPKI form |
| `created_at` | when the key was first seen |
| `revoked`, `revoked_at` | key revocation state |
| `is_primary` | current “primary” key (used for backward-compatible `publisherPublicKeyPem`) |

### `extension_versions.signing_key_id`

Each extension version persists the key that successfully verified its package at publish time:

- `signing_key_id` → foreign key to `publisher_keys.id`
- `signing_public_key_pem` → denormalized PEM used for verification (convenience/debugging)

This ensures the marketplace can always tell clients which key to use for a given version, even if
the publisher rotates keys later.

---

## Publish-time verification (server)

When a publisher publishes an extension version, the marketplace verifies the package signature
against **any non-revoked key** for that publisher:

- v1 packages: verify the detached signature over the raw bytes with each key until one succeeds
- v2 packages: call `verifyExtensionPackageV2(packageBytes, publicKeyPem)` with each key until one succeeds

The first key that verifies the package is persisted into `extension_versions.signing_key_id`.

---

## Download-time key delivery (server → client)

### `X-Publisher-Key-Id`

The package download response includes:

```
X-Publisher-Key-Id: <publisher_keys.id>
```

This identifies the exact key that signed the requested version.

### `GET /api/extensions/:id`

Extension metadata includes the publisher key set:

```ts
publisherKeys: Array<{ id: string; publicKeyPem: string; revoked: boolean }>;
```

For backward compatibility, the response still includes:

```ts
publisherPublicKeyPem: string | null; // the primary key
```

---

## Client verification strategy

Clients should verify packages using:

1. The version-specific key id from `X-Publisher-Key-Id` + the matching entry in `publisherKeys`, or
2. A fallback attempt across **all non-revoked** keys in `publisherKeys` (for older servers/edge cases).

Clients must **not** attempt verification with revoked keys.

---

## Revocation semantics

Revocation is intended for emergency response (e.g. key compromise):

- Revoked keys are excluded from publish-time verification (publish attempts signed with revoked keys fail).
- Clients ignore revoked keys during installation.
- Extension versions signed with a revoked key will no longer verify/install (by design).

Admin API (currently internal):

```
POST /api/publishers/:publisher/keys/:id/revoke
```

