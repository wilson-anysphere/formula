# Security & Enterprise Features

## Overview

Enterprise adoption requires meeting the highest bars for security, compliance, and administration. This includes data protection, access control, audit logging, and compliance certifications.

---

## Security Architecture

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  SECURITY LAYERS                                                            │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│  ┌─────────────┐    ┌─────────────┐    ┌─────────────┐                     │
│  │ Application │    │   Network   │    │    Data     │                     │
│  │  Security   │    │  Security   │    │  Security   │                     │
│  └─────────────┘    └─────────────┘    └─────────────┘                     │
│        │                  │                  │                              │
│  ┌─────▼─────────────────▼──────────────────▼─────┐                        │
│  │                                                 │                        │
│  │  • Authentication    • TLS 1.3        • Encryption                      │
│  │  • Authorization     • Certificate     • Access Control                  │
│  │  • Session Mgmt        Pinning       • Data Classification              │
│  │  • Input Validation  • Rate Limiting  • DLP                             │
│  │  • CSRF Protection   • WAF            • Backup/Recovery                  │
│  │                                                 │                        │
│  └─────────────────────────────────────────────────┘                        │
│                                                                             │
│  ┌─────────────────────────────────────────────────┐                        │
│  │              Audit & Compliance                  │                        │
│  │  • Activity Logging  • SIEM Integration          │                        │
│  │  • Compliance Reports • Retention Policies       │                        │
│  └─────────────────────────────────────────────────┘                        │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## Authentication

### Authentication Methods

| Method | Use Case | Implementation |
|--------|----------|----------------|
| Email/Password | Consumer users | Bcrypt hashing, secure session |
| SSO (SAML 2.0) | Enterprise | SAML assertions, IdP integration |
| SSO (OIDC) | Enterprise | OAuth 2.0 + OIDC, JWT tokens |
| API Keys | Automation | SHA-256 hashed, scoped permissions |
| MFA | All users (optional) | TOTP, WebAuthn, SMS backup |

### SSO Integration

See also: [docs/saml.md](./saml.md), [docs/oidc.md](./oidc.md)

#### OIDC SSO administration (services/api)

OIDC providers are configured **per organization** via org-admin API endpoints. Client secrets are stored in the database-backed secret store (`secrets` table) and encrypted using `SECRET_STORE_KEYS_JSON` (preferred) or the legacy `SECRET_STORE_KEY`.

Routes (org admin only; session or API key auth; session auth requires MFA):

- `GET /orgs/:orgId/oidc/providers` (alias: `/orgs/:orgId/oidc-providers`) → list configured providers (no secrets)
- `PUT /orgs/:orgId/oidc/providers/:providerId` (alias: `/orgs/:orgId/oidc-providers/:providerId`) → create/update a provider (optionally set `clientSecret`)
- `DELETE /orgs/:orgId/oidc/providers/:providerId` (alias: `/orgs/:orgId/oidc-providers/:providerId`) → deletes the provider and removes the stored secret

Production deployments **must** set `PUBLIC_BASE_URL` so the OIDC login flow can construct redirect URIs without trusting `Host` / `X-Forwarded-*` headers.

The OIDC `/auth/oidc/:orgId/:provider/start` and `/auth/oidc/:orgId/:provider/callback` endpoints are rate limited per IP. When limited, the API returns `429 { error: "too_many_requests" }` and includes a `Retry-After` header (seconds).

```typescript
interface SSOConfig {
  provider: "saml" | "oidc";
  
  // SAML settings
  saml?: {
    entryPoint: string;
    issuer: string;
    cert: string;
    signatureAlgorithm: "sha256" | "sha512";
    digestAlgorithm: "sha256" | "sha512";
    wantAuthnResponseSigned: boolean;
    wantAssertionsSigned: boolean;
    attributeMapping: {
      email: string;
      name: string;
      groups?: string;
    };
  };
  
  // OIDC settings
  oidc?: {
    issuer: string;
    clientId: string;
    clientSecret: string;
    redirectUri: string;
    scopes: string[];
    userInfoEndpoint: string;
  };
}

class SSOAuthenticator {
  async initiateSAMLLogin(config: SSOConfig): Promise<string> {
    const samlRequest = this.buildSAMLRequest(config.saml!);
    const redirectUrl = `${config.saml!.entryPoint}?SAMLRequest=${encodeURIComponent(samlRequest)}`;
    return redirectUrl;
  }
  
  async validateSAMLResponse(response: string, config: SSOConfig): Promise<UserInfo> {
    const assertion = this.parseSAMLResponse(response);
    
    // Validate signature
    if (!this.validateSignature(assertion, config.saml!.cert)) {
      throw new Error("Invalid SAML signature");
    }
    
    // Validate conditions (time, audience)
    this.validateConditions(assertion);
    
    // Extract user info
    return {
      email: this.extractAttribute(assertion, config.saml!.attributeMapping.email),
      name: this.extractAttribute(assertion, config.saml!.attributeMapping.name),
      groups: config.saml!.attributeMapping.groups
        ? this.extractAttributes(assertion, config.saml!.attributeMapping.groups)
        : []
    };
  }
}
```

### SAML 2.0 SSO (cloud backend implementation)

`services/api` implements **per-organization** SAML configuration and a hardened Assertion Consumer Service (ACS) flow.

#### Configuration (per organization)

Stored in `org_saml_providers` and managed via org-admin APIs:

- `GET /orgs/:orgId/saml-providers`
- `PUT /orgs/:orgId/saml-providers/:providerId`
- `DELETE /orgs/:orgId/saml-providers/:providerId`

Each provider stores:

- `idp_entry_point` (SSO URL)
- `idp_issuer`
- `idp_cert_pem` (X.509 PEM used to verify XML signatures)
- `sp_entity_id` (Audience + AuthnRequest Issuer)
- optional `attribute_mapping` (`email`, `name`, `groups`)

Configuration changes emit audit events:

- `org.saml_provider.created`
- `org.saml_provider.updated`
- `org.saml_provider.deleted`

#### Login endpoints

- `GET /auth/saml/:orgId/:provider/start` — builds a SAML `AuthnRequest` and redirects to the IdP (Redirect binding).
- `POST /auth/saml/:orgId/:provider/callback` — ACS endpoint (POST binding). On success, issues the standard session cookie.

`user_identities.provider` is namespaced as `saml:<providerId>` to avoid collisions with other identity types (e.g. OIDC).

The SAML `/start` and `/callback` endpoints are rate limited per IP. When limited, the API returns `429 { error: "too_many_requests" }` and includes a `Retry-After` header (seconds).

#### Assertion validation (defense-in-depth)

The ACS handler validates:

- XML signature on the **Assertion** or the **Response** (must be present)
- issuer (`idp_issuer`)
- audience (`sp_entity_id`)
- destination/recipient (ACS URL)
- `NotBefore` / `NotOnOrAfter` windows with clock skew
- `InResponseTo` tracking via server-side request cache + TTL
- replay prevention via assertion ID cache + TTL

Additionally:

- external entities are rejected (`<!DOCTYPE` / `<!ENTITY`)
- signature wrapping is mitigated by rejecting responses that contain multiple Assertion elements

#### MFA alignment

When `org_settings.require_mfa = true`, SAML logins are accepted only when the assertion’s `AuthnContextClassRef`
indicates MFA (conservative matching on common MFA/“multipleauthn” context values). Otherwise the login is rejected and audited.

### Session Management

```typescript
interface SessionConfig {
  // Session lifetime
  maxAge: number;              // e.g., 24 hours
  slidingExpiration: boolean;  // Reset on activity
  
  // Security
  secure: boolean;             // HTTPS only
  httpOnly: boolean;           // No JS access
  sameSite: "strict" | "lax" | "none";
  
  // Concurrency
  maxConcurrentSessions: number;
  invalidateOnPasswordChange: boolean;
}

class SessionManager {
  async createSession(userId: string, deviceInfo: DeviceInfo): Promise<Session> {
    // Check concurrent session limit
    const existingSessions = await this.getActiveSessions(userId);
    if (existingSessions.length >= this.config.maxConcurrentSessions) {
      // Invalidate oldest session
      await this.invalidateSession(existingSessions[0].id);
    }
    
    const session: Session = {
      id: crypto.randomUUID(),
      userId,
      token: this.generateSecureToken(),
      createdAt: new Date(),
      expiresAt: new Date(Date.now() + this.config.maxAge),
      deviceInfo,
      lastActivity: new Date()
    };
    
    await this.store.set(session.id, session);
    return session;
  }
  
  private generateSecureToken(): string {
    // 256 bits of cryptographically secure randomness
    const bytes = crypto.getRandomValues(new Uint8Array(32));
    return Array.from(bytes).map(b => b.toString(16).padStart(2, "0")).join("");
  }
}
```

---

## Authorization

### Role-Based Access Control (RBAC)

```typescript
type Role = "owner" | "admin" | "editor" | "commenter" | "viewer";

interface Permission {
  resource: ResourceType;
  action: ActionType;
}

type ResourceType = 
  | "document"
  | "sheet"
  | "range"
  | "cell"
  | "comment"
  | "version"
  | "settings";

type ActionType =
  | "create"
  | "read"
  | "update"
  | "delete"
  | "share"
  | "export"
  | "admin";

const ROLE_PERMISSIONS: Record<Role, Permission[]> = {
  owner: [
    { resource: "*", action: "*" }  // All permissions
  ],
  admin: [
    { resource: "document", action: "read" },
    { resource: "document", action: "update" },
    { resource: "document", action: "share" },
    { resource: "sheet", action: "*" },
    { resource: "range", action: "*" },
    { resource: "settings", action: "*" }
  ],
  editor: [
    { resource: "document", action: "read" },
    { resource: "sheet", action: "read" },
    { resource: "sheet", action: "update" },
    { resource: "range", action: "*" },
    { resource: "cell", action: "*" },
    { resource: "comment", action: "*" }
  ],
  commenter: [
    { resource: "document", action: "read" },
    { resource: "sheet", action: "read" },
    { resource: "comment", action: "create" },
    { resource: "comment", action: "read" }
  ],
  viewer: [
    { resource: "document", action: "read" },
    { resource: "sheet", action: "read" },
    { resource: "comment", action: "read" }
  ]
};

class AuthorizationService {
  async checkPermission(
    userId: string,
    documentId: string,
    resource: ResourceType,
    action: ActionType
  ): Promise<boolean> {
    // Get user's role for this document
    const role = await this.getDocumentRole(userId, documentId);
    if (!role) return false;
    
    // Check role permissions
    const permissions = ROLE_PERMISSIONS[role];
    return permissions.some(p => 
      (p.resource === "*" || p.resource === resource) &&
      (p.action === "*" || p.action === action)
    );
  }
  
  async checkCellPermission(
    userId: string,
    documentId: string,
    cell: CellRef,
    action: "read" | "update"
  ): Promise<boolean> {
    // Check document-level permission first
    const docPermission = await this.checkPermission(userId, documentId, "cell", action);
    if (!docPermission) return false;
    
    // Check cell-level restrictions
    const cellRestriction = await this.getCellRestriction(documentId, cell);
    if (cellRestriction) {
      return cellRestriction.allowedUsers.includes(userId);
    }
    
    return true;
  }
}
```

### Cell-Level Permissions

```typescript
interface CellRestriction {
  range: Range;
  allowedUsers: string[];
  allowedGroups: string[];
  permissionType: "read" | "edit";
}

class CellPermissionManager {
  async restrictRange(
    documentId: string,
    range: Range,
    options: {
      allowedUsers?: string[];
      allowedGroups?: string[];
      permissionType: "read" | "edit";
    }
  ): Promise<void> {
    const restriction: CellRestriction = {
      range,
      allowedUsers: options.allowedUsers || [],
      allowedGroups: options.allowedGroups || [],
      permissionType: options.permissionType
    };
    
    await this.store.addRestriction(documentId, restriction);
    
    // Notify affected users
    await this.notifyRestrictionChange(documentId, restriction);
  }
  
  async getVisibleCells(
    userId: string,
    documentId: string,
    requestedRange: Range
  ): Promise<CellData[][]> {
    // NOTE: In the real-time collaboration (Yjs) system, the sync server broadcasts
    // CRDT updates and cannot filter cell-level payloads per connection. That means
    // returning `{ value: "###" }` here is only *UI masking* and does not provide
    // confidentiality by itself.
    //
    // For truly confidential protected ranges, cell contents must be end-to-end
    // encrypted client-side before entering the CRDT (stored under the Yjs `enc`
    // field with per-range keys).
    const restrictions = await this.getRestrictions(documentId);
    const userGroups = await this.getUserGroups(userId);
    
    const result: CellData[][] = [];
    
    for (let row = requestedRange.startRow; row <= requestedRange.endRow; row++) {
      const rowData: CellData[] = [];
      for (let col = requestedRange.startCol; col <= requestedRange.endCol; col++) {
        const canRead = this.checkCellAccess(
          { row, col },
          restrictions,
          userId,
          userGroups,
          "read"
        );
        
        if (canRead) {
          rowData.push(await this.getCell(documentId, row, col));
        } else {
          rowData.push({ value: "###", hidden: true });
        }
      }
      result.push(rowData);
    }
    
    return result;
  }
}
```

---

## Data Classification & DLP

Formula supports attaching **data classifications** to parts of a document (document/sheet/column/range/cell). These classifications are then used by Data Loss Prevention (DLP) policy checks to decide whether operations like sharing, export, clipboard, or AI processing are allowed, blocked, or redacted.

### Selector scopes

Classifications are stored as `document_classifications` records with a `selector` describing the scope:

- `document`: `{ scope: "document", documentId }`
- `sheet`: `{ scope: "sheet", documentId, sheetId }`
- `column`: `{ scope: "column", documentId, sheetId, columnIndex }`
- `range`: `{ scope: "range", documentId, sheetId, range: { start, end } }`
- `cell`: `{ scope: "cell", documentId, sheetId, row, col }`

### Effective classification resolution

`services/api` supports two related (but intentionally different) classification resolution modes:

1) **Selector precedence** (used by the classification resolve endpoint)
   - Cell queries:
     - exact `cell` selector (if present)
     - otherwise the **smallest containing** `range` selector
     - otherwise matching `column` → `sheet` → `document` → default `Public`
   - Range queries:
     - the **smallest containing** `range` selector (if present)
     - if the query is a single column, falls back to matching `column`
     - otherwise matching `sheet` → `document` → default `Public`

2) **Aggregate semantics** (used for DLP evaluation of `range` selectors)
   - When DLP is evaluated for a `range` selector, the backend computes the **maximum** classification level across all selectors that *intersect* the queried range (including cells inside the range) and unions labels across all matches. This is intentionally conservative for bulk operations like copy/export.

### API: resolve endpoint

`services/api` exposes `POST /docs/:docId/classifications/resolve` (requires document read permission) which uses **selector precedence**:

- Request body: `{ selector: <cell|range selector> }`
- Response: `{ classification, source }` where `source` includes the resolved scope and `selectorKey`.

### DLP policy evaluation & enforcement (services/api)

DLP policies are stored at two levels:

- **Org policy**: `org_dlp_policies` (per organization)
- **Document overrides**: `document_dlp_policies` (per document), applied only when the effective org policy allows overrides.

If an organization has no `org_dlp_policies` row, `services/api` falls back to a built-in default policy so DLP remains enforceable even before an explicit org configuration is stored.

The backend exposes a policy evaluation endpoint for clients to understand whether an action is permitted:

- `POST /docs/:docId/dlp/evaluate` (requires document read permission)
  - Body: `{ action, selector?, options? }`
  - Response: `{ decision, reasonCode?, classification, maxAllowed }`

This endpoint returns *decisions*, not full policy configuration, so clients can build UX without needing org-admin access to the raw policy.

#### External share-link enforcement

Public (external) share links are enforced server-side using `DLP_ACTION.SHARE_EXTERNAL_LINK`:

- `POST /docs/:docId/share-links` blocks creation when `visibility = "public"` is disallowed by DLP.
- `POST /share-links/:token/redeem` re-evaluates DLP when the redeemer is not already an org member (treat as “external access”).

To avoid leaking sensitive classification/policy thresholds, the API responds with:

- `403 { error: "dlp_blocked" }`

When an action is blocked, an audit event is written (`event_type = "dlp.blocked"`) including the action, document id, effective classification, and reason code.

## Data Encryption

### Encryption at Rest

#### Desktop (local) encryption

Formula Desktop (Tauri) stores sensitive local state encrypted-at-rest using AES-256-GCM with keys persisted in the OS keychain.

Examples:

- **Power Query credentials** + refresh scheduling metadata: stored as encrypted blobs under the app data directory (key material in keychain).
- **Power Query query-result cache**: persisted in IndexedDB and wrapped in `EncryptedCacheStore` so cached results (including Arrow IPC bytes) are never stored in plaintext.

```typescript
interface EncryptionConfig {
  algorithm: "aes-256-gcm";
  keyManagement: "local" | "aws-kms" | "azure-keyvault" | "gcp-kms";
  keyRotationDays: number;
}

class EncryptionService {
  private currentKey: CryptoKey;
  private keyVersion: number;
  
  async encryptDocument(data: Uint8Array): Promise<EncryptedData> {
    const iv = crypto.getRandomValues(new Uint8Array(12));
    
    const encrypted = await crypto.subtle.encrypt(
      { name: "AES-GCM", iv },
      this.currentKey,
      data
    );
    
    return {
      ciphertext: new Uint8Array(encrypted),
      iv,
      keyVersion: this.keyVersion,
      algorithm: "aes-256-gcm"
    };
  }
  
  async decryptDocument(encrypted: EncryptedData): Promise<Uint8Array> {
    const key = await this.getKeyForVersion(encrypted.keyVersion);
    
    const decrypted = await crypto.subtle.decrypt(
      { name: "AES-GCM", iv: encrypted.iv },
      key,
      encrypted.ciphertext
    );
    
    return new Uint8Array(decrypted);
  }
  
  async rotateKey(): Promise<void> {
    // Generate new key
    const newKey = await crypto.subtle.generateKey(
      { name: "AES-GCM", length: 256 },
      true,
      ["encrypt", "decrypt"]
    );
    
    // Store old key for decryption
    await this.archiveKey(this.currentKey, this.keyVersion);
    
    // Update current key
    this.currentKey = newKey;
    this.keyVersion++;
    
    // Schedule background re-encryption of old documents
    await this.scheduleReEncryption();
  }
}
```

#### Cloud backend implementation (services/api)

The cloud backend implements **envelope encryption** for sensitive database blobs starting with `document_versions.data`.

- **Policy gate:** `org_settings.cloud_encryption_at_rest`
  - `true`: new writes are encrypted; reads transparently decrypt
  - `false`: plaintext writes are allowed; reads still decrypt encrypted rows (mixed-mode rollout)
- **Algorithm:** AES-256-GCM for payload encryption
- **Per-blob metadata (stored in Postgres):**
  - `document_versions.data_ciphertext`, `data_iv`, `data_tag`
  - `document_versions.data_encrypted_dek` (wrapped DEK)
    - **Envelope schema v1:** base64-encoded bytes (legacy)
    - **Envelope schema v2:** JSON string (canonical `packages/security` wrapped-key object)
  - `document_versions.data_kms_provider`, `data_kms_key_id` (provider + key identifier/version for debugging)
  - `document_versions.data_aad` (AAD includes `orgId`, `documentId`, `documentVersionId`, and `envelopeVersion`)
  - `document_versions.data_envelope_version`
    - `1`: legacy services/api envelope format (HKDF local KMS)
    - `2`: canonical `packages/security/crypto/envelope.js` format

##### KMS providers

The canonical KMS provider interface lives in `packages/security/crypto/envelope.js` and uses:

```ts
type EncryptionContext = Record<string, unknown> | null;

interface EnvelopeKmsProvider {
  provider: string; // e.g. "local", "aws"
  wrapKey(args: { plaintextKey: Buffer; encryptionContext?: EncryptionContext }): Promise<unknown>;
  unwrapKey(args: { wrappedKey: unknown; encryptionContext?: EncryptionContext }): Promise<Buffer>;
}
```

- **Local (dev/test):** `kms_provider = 'local'`
  - Canonical implementation: `packages/security/crypto/kms/localKmsProvider.js`
  - Per-org KEK material is persisted + versioned in Postgres (`org_kms_local_state`)
  - Legacy support: `LOCAL_KMS_MASTER_KEY` is only required to decrypt **envelope schema v1** rows
- **AWS (optional):** `kms_provider = 'aws'`
  - Canonical implementation: `packages/security/crypto/kms/providers.js` (`AwsKmsProvider`)
  - Requires `AWS_KMS_ENABLED=true`, `AWS_REGION`, and installing `@aws-sdk/client-kms` in the runtime image
    (typically as a dependency of the service using the provider, e.g. `services/api`)
  - The AWS SDK is loaded lazily; if the dependency is missing, a clear runtime error is thrown
- **GCP / Azure:** stubs exist under the same provider interface but are not implemented in this reference repo

##### Key rotation

`services/api` includes a rotation script for `kms_provider = 'local'` that:

1. Checks each org’s `org_settings.key_rotation_days` against `org_settings.kms_key_rotated_at`
2. Rotates the org’s local KEK version in `org_kms_local_state`
3. **Re-wraps DEKs** in `document_versions.data_encrypted_dek` to the latest KEK version **without re-encrypting ciphertext**

Run with:

```bash
cd services/api
npm run keys:rotate
```

##### Backfilling existing plaintext versions
 
During rollout you may have older `document_versions` rows that still have plaintext `data` populated.
To encrypt those rows in-place for orgs with `cloud_encryption_at_rest = true`, run:

```bash
cd services/api
# Optional: scope to a single org + limit rows per run
ORG_ID="<org-uuid>" BATCH_SIZE=100 npm run versions:encrypt
```

##### Migrating legacy envelope schema v1 rows

Deployments that previously used the legacy (HKDF-based) local KMS model will have
`document_versions` rows with `data_envelope_version = 1`. Those rows remain readable, but they still
require `LOCAL_KMS_MASTER_KEY` for decryption.

To upgrade those rows in-place to the canonical schema v2 representation **without re-encrypting
ciphertext** (DEK re-wrap only), run:

```bash
cd services/api
ORG_ID="<org-uuid>" BATCH_SIZE=100 npm run versions:migrate-legacy
```

Note: migrating rows that were encrypted with the legacy local KMS provider requires `LOCAL_KMS_MASTER_KEY` to be set.

##### Database-backed secret store (`secrets` table)

`services/api` uses a small database-backed secret store (`secrets` table) for sensitive configuration
like per-org OIDC client secrets and SIEM auth tokens. Values are encrypted using **AES-256-GCM**.

As of **secret store v2**, ciphertext is bound to the secret's metadata via **AAD**:

- A fixed context string (`formula-secret-store`)
- The secret `name` (which often includes an org scope like `oidc:<orgId>:<providerId>`)

This means secrets cannot be copied between rows (or renamed) without re-encrypting: decryption will
fail if ciphertext is moved to a different `name`.

###### Configuration / keyring

For rotation, configure multiple keys and choose the "current" key for new writes:

- `SECRET_STORE_KEYS` (recommended): comma-separated list of `<keyId>:<base64>` entries.
  - The **last** entry is used for encryption; **all** entries are accepted for decryption.
  - Example:

    ```bash
    SECRET_STORE_KEYS="v1:BASE64_32_BYTES,v2:BASE64_32_BYTES"
    ```

- `SECRET_STORE_KEYS_JSON`: JSON object containing `{ currentKeyId, keys }` (where `keys` maps keyId -> base64 key).

The legacy `SECRET_STORE_KEY` is still supported for smooth upgrades; it is hashed with SHA-256 to
derive a 32-byte AES-256 key.

To generate a new 32-byte key:

```bash
node -e 'console.log(require("crypto").randomBytes(32).toString("base64"))'
```

###### Rotation / re-encryption

After adding a new key (and making it the last entry / current key), re-encrypt existing rows with:

```bash
cd services/api
npm run secrets:rotate
```

###### Migrating from legacy `SECRET_STORE_KEY` to `SECRET_STORE_KEYS`

If you have existing `v1:` rows encrypted with the legacy passphrase, include the derived key bytes
as one of the entries in `SECRET_STORE_KEYS`. You can compute the base64 key material with:

```bash
node -e 'const crypto=require("crypto"); console.log(crypto.createHash("sha256").update(process.env.SECRET_STORE_KEY,"utf8").digest("base64"))'
```

Then set `SECRET_STORE_KEYS` with that derived key plus your new key(s), deploy, and run
`npm run secrets:rotate` to rewrite existing rows to the latest key.

### Encryption in Transit

```typescript
interface TLSConfig {
  minVersion: "TLSv1.2" | "TLSv1.3";
  cipherSuites: string[];
  certificatePinning: boolean;
  pinnedCertificates?: string[];
  hsts: {
    enabled: boolean;
    maxAge: number;
    includeSubDomains: boolean;
    preload: boolean;
  };
}

const SECURE_TLS_CONFIG: TLSConfig = {
  minVersion: "TLSv1.3",
  cipherSuites: [
    "TLS_AES_256_GCM_SHA384",
    "TLS_CHACHA20_POLY1305_SHA256",
    "TLS_AES_128_GCM_SHA256"
  ],
  certificatePinning: true,
  hsts: {
    enabled: true,
    maxAge: 31536000,  // 1 year
    includeSubDomains: true,
    preload: true
  }
};
```

In this repo, outbound enterprise integrations (starting with SIEM delivery) enforce a **minimum TLS version of TLS 1.3**.
Organizations can optionally enable **certificate pinning** via `org_settings`:

- `certificate_pinning_enabled` (boolean)
- `certificate_pins` (JSON array of SHA-256 certificate fingerprints)

For operational details (how pins are computed, rotation guidance), see `docs/tls-pinning.md`.

---

## Audit Logging

### Audit Event Types

```typescript
type AuditEventType =
  // Authentication
  | "auth.login"
  | "auth.logout"
  | "auth.login_failed"
  | "auth.mfa_enabled"
  | "auth.mfa_disabled"
  | "auth.password_changed"
  | "auth.session_expired"
  
  // Document operations
  | "document.created"
  | "document.opened"
  | "document.modified"
  | "document.deleted"
  | "document.exported"
  | "document.printed"
  
  // Sharing
  | "sharing.added"
  | "sharing.removed"
  | "sharing.modified"
  | "sharing.link_created"
  | "sharing.link_revoked"
  
  // Data access
  | "data.viewed"
  | "data.downloaded"
  | "data.copied"
  
  // Admin
  | "admin.user_created"
  | "admin.user_deleted"
  | "admin.settings_changed"
  | "admin.integration_added";

interface AuditEvent {
  id: string;
  timestamp: Date;
  eventType: AuditEventType;
  
  // Actor
  userId: string;
  userEmail: string;
  ipAddress: string;
  userAgent: string;
  sessionId: string;
  
  // Target
  resourceType: string;
  resourceId: string;
  resourceName?: string;
  
  // Details
  details: Record<string, any>;
  
  // Outcome
  success: boolean;
  errorCode?: string;
  errorMessage?: string;
}

class AuditLogger {
  async log(event: Omit<AuditEvent, "id" | "timestamp">): Promise<void> {
    const fullEvent: AuditEvent = {
      ...event,
      id: crypto.randomUUID(),
      timestamp: new Date()
    };
    
    // Write to primary store
    await this.store.insert(fullEvent);
    
    // Send to SIEM if configured
    if (this.siemConfig) {
      await this.sendToSIEM(fullEvent);
    }
    
    // Check for suspicious activity
    await this.anomalyDetector.analyze(fullEvent);
  }
  
  async query(filters: AuditQueryFilters): Promise<AuditEvent[]> {
    return this.store.query({
      userId: filters.userId,
      eventTypes: filters.eventTypes,
      resourceId: filters.resourceId,
      startTime: filters.startTime,
      endTime: filters.endTime,
      success: filters.success
    });
  }
  
  async generateComplianceReport(
    startDate: Date,
    endDate: Date,
    reportType: "access" | "changes" | "sharing" | "full"
  ): Promise<ComplianceReport> {
    const events = await this.query({
      startTime: startDate,
      endTime: endDate
    });
    
    return {
      period: { start: startDate, end: endDate },
      summary: this.summarizeEvents(events),
      details: this.categorizeEvents(events),
      anomalies: await this.detectAnomalies(events),
      recommendations: this.generateRecommendations(events)
    };
  }
}
```

### SIEM Integration

For concrete export formats (JSON/CEF/LEEF), batching/retry behavior, and per-organization configuration endpoints, see [`docs/siem.md`](./siem.md).

```typescript
// `services/api` implements SIEM as a worker-driven export:
// - Per-org config is stored in Postgres and managed via `GET/PUT/DELETE /orgs/:orgId/siem`.
// - Audit logs can be queried/exported via `GET /orgs/:orgId/audit` and `/orgs/:orgId/audit/export`.
// - A background worker (`SiemExportWorker`) batches + redacts events and POSTs them to the configured SIEM endpoint.

type SiemAuthConfig =
  | { type: "none" }
  | { type: "bearer"; token: string }
  | { type: "basic"; username: string; password: string }
  | { type: "header"; name: string; value: string };

interface SiemEndpointConfig {
  endpointUrl: string;
  format?: "json" | "cef" | "leef";
  batchSize?: number;
  timeoutMs?: number;
  idempotencyKeyHeader?: string | null;
  headers?: Record<string, string>;
  auth?: SiemAuthConfig;
  retry?: {
    maxAttempts?: number;
    baseDelayMs?: number;
    maxDelayMs?: number;
    jitter?: boolean;
  };
}
```

---

## Compliance

### Supported Standards

| Standard | Requirements | Implementation |
|----------|--------------|----------------|
| **SOC 2 Type II** | Security, availability, confidentiality | Audit logging, access control, encryption |
| **ISO 27001** | ISMS framework | Security policies, risk management |
| **GDPR** | Data protection, privacy | Data minimization, consent, erasure |
| **HIPAA** | Healthcare data protection | BAA support, audit controls |
| **CCPA** | California privacy | Data access, deletion rights |
| **FedRAMP** | US government | Enhanced security controls |

### Data Residency

```typescript
interface DataResidencyConfig {
  region: "us" | "eu" | "apac" | "custom";
  customRegions?: string[];
  
  // Storage locations
  primaryStorage: string;
  backupStorage: string;
  
  // Processing restrictions
  allowCrossRegionProcessing: boolean;
  aiProcessingRegion?: string;
}

class DataResidencyManager {
  async validateOperation(
    operation: "store" | "process" | "transfer",
    dataClassification: string,
    targetRegion: string,
    config: DataResidencyConfig
  ): Promise<ValidationResult> {
    // Check if operation is allowed for this data classification
    const rules = await this.getClassificationRules(dataClassification);
    
    if (!rules.allowedRegions.includes(targetRegion)) {
      return {
        allowed: false,
        reason: `Data classification "${dataClassification}" prohibits ${operation} in region "${targetRegion}"`
      };
    }
    
    if (operation === "transfer" && !config.allowCrossRegionProcessing) {
      return {
        allowed: false,
        reason: "Cross-region data transfer is disabled"
      };
    }
    
    return { allowed: true };
  }
}
```

### Data Retention

```typescript
interface RetentionPolicy {
  documentRetention: number;     // Days
  versionRetention: number;      // Days
  auditLogRetention: number;     // Days
  deletedDocumentRetention: number; // Days before permanent deletion
  
  // Exceptions
  legalHoldOverride: boolean;
  regulatoryMinimum?: number;
}

class RetentionManager {
  async applyRetentionPolicy(organizationId: string): Promise<RetentionReport> {
    const policy = await this.getPolicy(organizationId);
    const now = new Date();
    
    const report: RetentionReport = {
      documentsDeleted: 0,
      versionsDeleted: 0,
      auditLogsArchived: 0,
      errors: []
    };
    
    // Check legal holds first
    const legalHolds = await this.getLegalHolds(organizationId);
    const heldDocuments = new Set(legalHolds.map(h => h.documentId));
    
    // Process old versions
    const oldVersions = await this.getVersionsOlderThan(
      organizationId,
      new Date(now.getTime() - policy.versionRetention * 24 * 60 * 60 * 1000)
    );
    
    for (const version of oldVersions) {
      if (heldDocuments.has(version.documentId)) continue;
      
      try {
        await this.deleteVersion(version.id);
        report.versionsDeleted++;
      } catch (error) {
        report.errors.push({ type: "version", id: version.id, error: error.message });
      }
    }
    
    // Archive old audit logs
    const archiveDate = new Date(
      now.getTime() - policy.auditLogRetention * 24 * 60 * 60 * 1000
    );
    report.auditLogsArchived = await this.archiveAuditLogs(organizationId, archiveDate);
    
    return report;
  }
}
```

---

## Enterprise Administration

### Organization Management

```typescript
interface Organization {
  id: string;
  name: string;
  domain: string;
  plan: "free" | "team" | "business" | "enterprise";
  
  settings: OrganizationSettings;
  
  createdAt: Date;
  billingEmail: string;
}

interface OrganizationSettings {
  // Security
  requireMFA: boolean;
  allowedAuthMethods: ("password" | "sso" | "mfa")[];
  sessionTimeout: number;
  ipAllowlist?: string[];
  
  // Sharing
  allowExternalSharing: boolean;
  allowPublicLinks: boolean;
  defaultPermission: "viewer" | "commenter" | "editor";
  
  // Data
  dataResidency: DataResidencyConfig;
  retentionPolicy: RetentionPolicy;
  
  // Integrations
  allowedIntegrations: string[];
  
  // AI
  aiEnabled: boolean;
  aiDataProcessingConsent: boolean;
}

class OrganizationAdmin {
  async updateSettings(
    orgId: string,
    updates: Partial<OrganizationSettings>
  ): Promise<void> {
    // Validate settings
    this.validateSettings(updates);
    
    // Apply updates
    await this.store.updateOrganization(orgId, { settings: updates });
    
    // Log change
    await this.auditLogger.log({
      eventType: "admin.settings_changed",
      userId: this.currentUser.id,
      userEmail: this.currentUser.email,
      resourceType: "organization",
      resourceId: orgId,
      details: { updates },
      success: true
    });
    
    // Notify affected users if needed
    if (updates.requireMFA) {
      await this.notifyMFARequired(orgId);
    }
  }
}
```

### User Provisioning (SCIM)

```typescript
// SCIM 2.0 API for automated user provisioning
class SCIMService {
  // GET /scim/v2/Users
  async listUsers(filter?: string, startIndex?: number, count?: number): Promise<SCIMListResponse> {
    const users = await this.userStore.list({
      filter: this.parseFilter(filter),
      offset: startIndex ? startIndex - 1 : 0,
      limit: count || 100
    });
    
    return {
      schemas: ["urn:ietf:params:scim:api:messages:2.0:ListResponse"],
      totalResults: users.total,
      startIndex: startIndex || 1,
      itemsPerPage: users.items.length,
      Resources: users.items.map(u => this.toSCIMUser(u))
    };
  }
  
  // POST /scim/v2/Users
  async createUser(scimUser: SCIMUser): Promise<SCIMUser> {
    const user = await this.userStore.create({
      email: scimUser.emails.find(e => e.primary)?.value || scimUser.userName,
      name: scimUser.displayName || `${scimUser.name?.givenName} ${scimUser.name?.familyName}`,
      active: scimUser.active !== false
    });
    
    await this.auditLogger.log({
      eventType: "admin.user_created",
      resourceType: "user",
      resourceId: user.id,
      details: { source: "scim", email: user.email }
    });
    
    return this.toSCIMUser(user);
  }
  
  // PATCH /scim/v2/Users/:id
  async patchUser(userId: string, operations: SCIMPatchOp[]): Promise<SCIMUser> {
    const user = await this.userStore.get(userId);
    if (!user) throw new Error("User not found");
    
    for (const op of operations) {
      switch (op.op) {
        case "replace":
          if (op.path === "active") {
            user.active = op.value;
          } else if (op.path === "displayName") {
            user.name = op.value;
          }
          break;
        case "add":
          // Handle add operations
          break;
        case "remove":
          // Handle remove operations
          break;
      }
    }
    
    await this.userStore.update(user);
    return this.toSCIMUser(user);
  }
}
```

---

## Disaster Recovery

### Backup Strategy

```typescript
interface BackupConfig {
  frequency: "hourly" | "daily" | "weekly";
  retention: number;  // Number of backups to keep
  type: "full" | "incremental";
  destination: BackupDestination;
  encryption: boolean;
}

interface BackupDestination {
  type: "s3" | "azure-blob" | "gcs" | "local";
  config: Record<string, string>;
}

class BackupService {
  async performBackup(config: BackupConfig): Promise<BackupResult> {
    const startTime = Date.now();
    
    // Create backup manifest
    const manifest: BackupManifest = {
      id: crypto.randomUUID(),
      timestamp: new Date(),
      type: config.type,
      documents: [],
      metadata: {}
    };
    
    // Get documents to backup
    const documents = config.type === "full"
      ? await this.getAllDocuments()
      : await this.getModifiedDocuments(this.lastBackupTime);
    
    // Backup each document
    for (const doc of documents) {
      const docBackup = await this.backupDocument(doc);
      manifest.documents.push(docBackup);
    }
    
    // Encrypt if configured
    const data = config.encryption
      ? await this.encryption.encrypt(JSON.stringify(manifest))
      : JSON.stringify(manifest);
    
    // Upload to destination
    await this.upload(config.destination, manifest.id, data);
    
    // Clean up old backups
    await this.cleanupOldBackups(config);
    
    return {
      success: true,
      backupId: manifest.id,
      documentCount: manifest.documents.length,
      duration: Date.now() - startTime
    };
  }
  
  async restore(backupId: string): Promise<RestoreResult> {
    // Download backup
    const data = await this.download(backupId);
    
    // Decrypt if needed
    const manifest: BackupManifest = JSON.parse(
      typeof data === "string" ? data : await this.encryption.decrypt(data)
    );
    
    // Restore documents
    const results: DocumentRestoreResult[] = [];
    for (const docBackup of manifest.documents) {
      try {
        await this.restoreDocument(docBackup);
        results.push({ documentId: docBackup.id, success: true });
      } catch (error) {
        results.push({ documentId: docBackup.id, success: false, error: error.message });
      }
    }
    
    return {
      backupId,
      timestamp: manifest.timestamp,
      documentsRestored: results.filter(r => r.success).length,
      documentsFailed: results.filter(r => !r.success).length,
      details: results
    };
  }
}
```

---

## Security Testing

### Penetration Testing Checklist

- [ ] **Authentication**
  - [ ] Password brute force protection
  - [ ] Session fixation
  - [ ] Session hijacking
  - [ ] MFA bypass attempts
  
- [ ] **Authorization**
  - [ ] Horizontal privilege escalation
  - [ ] Vertical privilege escalation
  - [ ] IDOR (Insecure Direct Object Reference)
  - [ ] Cell-level permission bypass
  
- [ ] **Input Validation**
  - [ ] SQL injection (N/A if no SQL)
  - [ ] XSS (stored, reflected, DOM)
  - [ ] Formula injection
  - [ ] File upload validation
  
- [ ] **Data Protection**
  - [ ] Sensitive data exposure
  - [ ] Encryption validation
  - [ ] Key management
  
- [ ] **API Security**
  - [ ] Rate limiting
  - [ ] Authentication bypass
  - [ ] Mass assignment
  
- [ ] **Infrastructure**
  - [ ] TLS configuration
  - [ ] CORS policy
  - [ ] Security headers

### Security Scanning

```typescript
class SecurityScanner {
  async runScan(scope: ScanScope): Promise<ScanReport> {
    const findings: SecurityFinding[] = [];
    
    // Static analysis
    if (scope.includeStatic) {
      findings.push(...await this.staticAnalysis());
    }
    
    // Dependency vulnerabilities
    if (scope.includeDependencies) {
      findings.push(...await this.checkDependencies());
    }
    
    // Configuration review
    if (scope.includeConfig) {
      findings.push(...await this.reviewConfiguration());
    }
    
    // Categorize findings
    const categorized = this.categorizeBySeverity(findings);
    
    return {
      scanDate: new Date(),
      scope,
      totalFindings: findings.length,
      critical: categorized.critical.length,
      high: categorized.high.length,
      medium: categorized.medium.length,
      low: categorized.low.length,
      findings
    };
  }
  
  private async checkDependencies(): Promise<SecurityFinding[]> {
    // Check npm audit, cargo audit, etc.
    const npmAudit = await this.runNpmAudit();
    const cargoAudit = await this.runCargoAudit();
    
    return [...npmAudit, ...cargoAudit].map(vuln => ({
      type: "dependency_vulnerability",
      severity: vuln.severity,
      package: vuln.package,
      version: vuln.version,
      vulnerability: vuln.cve,
      recommendation: vuln.recommendation
    }));
  }
}
```
