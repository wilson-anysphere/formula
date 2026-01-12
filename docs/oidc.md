# OIDC SSO

Formula supports per-organization OpenID Connect (OIDC) SSO via the **Authorization Code** flow with **PKCE**:

- SP → IdP: redirect to the OIDC authorization endpoint (`/start`)
- IdP → SP: redirect back to the API callback (`/callback`)
- API → IdP: token exchange (server-to-server)

## Provider configuration (per org)

OIDC providers are stored in `org_oidc_providers` and managed through org-admin APIs. Client secrets are stored in the encrypted secret store (`secrets` table) under the key `oidc:<orgId>:<providerId>`.

### Fields

- `providerId` (path param): A short identifier for the IdP configuration (e.g. `okta`, `azuread`).
- `issuerUrl`: OIDC issuer URL (must be a valid URL; **HTTPS required in production**; must not include credentials, query params, fragments, or localhost in production). Trailing slashes are stripped.
- `clientId`: OIDC client id.
- `clientSecret`: OIDC client secret (**stored in the secret store**, not returned by list/get endpoints).
- `scopes`: Scopes requested during login (default: `["openid", "email", "profile"]`). The API always ensures `openid` is included.
- `enabled`: Whether this provider can be used for login.

## Admin APIs

All endpoints require an authenticated **org admin**. When using session auth, the admin must have satisfied org MFA (`org_settings.require_mfa`).

- Primary endpoints:
  - `GET /orgs/:orgId/oidc-providers`
  - `GET /orgs/:orgId/oidc-providers/:providerId`
  - `PUT /orgs/:orgId/oidc-providers/:providerId`
  - `DELETE /orgs/:orgId/oidc-providers/:providerId`

- Legacy aliases:
  - `GET /orgs/:orgId/oidc/providers`
  - `GET /orgs/:orgId/oidc/providers/:providerId`
  - `PUT /orgs/:orgId/oidc/providers/:providerId`
  - `DELETE /orgs/:orgId/oidc/providers/:providerId`

The list/get endpoints include `clientSecretConfigured` to indicate whether a client secret exists in the secret store.

## Login flow

### Start login

`GET /auth/oidc/:orgId/:provider/start`

Redirects the browser to the IdP authorization endpoint. The API generates and stores:

- `state` (anti-CSRF)
- `nonce` (binds the id_token to the auth request)
- PKCE verifier/challenge

### Callback

`GET /auth/oidc/:orgId/:provider/callback`

Accepts the OIDC redirect with query params:

- `code` (required)
- `state` (required)
- `error` / `error_description` (optional)

On success, the API:

1. Validates and consumes the `state` (single-use, TTL-backed).
2. Exchanges the authorization code for tokens using the stored `clientSecret`.
3. Verifies the `id_token`:
   - issuer (`issuerUrl`)
   - audience (`clientId`)
   - signature (JWKS)
   - `nonce`
4. Extracts identity:
   - `subject = sub`
   - email from `email` / `preferred_username` / `upn`
5. Links identity in `user_identities` (`provider = providerId`, `subject = sub`, `org_id = orgId`).
6. Provisions the user + org membership if needed.
7. Issues a session cookie and writes an `auth.login` audit event (`details: { method: "oidc", provider }`).

### `allowed_auth_methods`

OIDC login is only allowed if the org has `"oidc"` present in `org_settings.allowed_auth_methods`.

## Deployment notes

### PUBLIC_BASE_URL (recommended)

Set `PUBLIC_BASE_URL` (example: `https://api.example.com`) in production so the API can build the callback URL without relying on request headers.

### Rate limiting

The OIDC `/start` and `/callback` endpoints are rate limited per IP. When rate limited, the API returns `429 { error: "too_many_requests" }` and includes a `Retry-After` header (seconds).
