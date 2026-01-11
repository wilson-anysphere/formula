# SAML 2.0 SSO

Formula supports per-organization SAML 2.0 SSO via IdP-initiated **HTTP-POST** bindings.

## Provider configuration (per org)

SAML providers are stored in `org_saml_providers` and managed through org-admin APIs.

### Fields

- `providerId` (path param): A short identifier for the IdP configuration (e.g. `okta`, `azuread`).
- `entryPoint`: IdP SSO URL (must be a valid URL; **HTTPS required in production**).
- `issuer`: Service Provider issuer / Entity ID. This is also used as the SAML **audience** when validating assertions.
- `idpCertPem`: IdP signing certificate in PEM format (the public cert used to validate XML signatures).
- `wantAssertionsSigned`: Require `<Assertion>` signatures (default `true`).
- `wantResponseSigned`: Require `<Response>` signatures (default `true`).
- `attributeMapping`: Attribute names used to extract user identity:
  - `email` (required): attribute containing the user email.
  - `name` (required): attribute containing the display name.
  - `groups` (optional): attribute containing groups (currently not persisted).
- `enabled`: Whether this provider can be used for login.

### Certificate format

`idpCertPem` should be a PEM-encoded X.509 certificate, for example:

```
-----BEGIN CERTIFICATE-----
MIIC...
-----END CERTIFICATE-----
```

## Admin APIs

All endpoints require an authenticated **org admin**.

- `GET /orgs/:orgId/saml/providers`
- `PUT /orgs/:orgId/saml/providers/:providerId`
- `DELETE /orgs/:orgId/saml/providers/:providerId`

Changes emit audit events:

- `admin.integration_added`
- `admin.integration_updated`
- `admin.integration_removed`

with `details: { type: "saml", providerId }`.

## Login flow

### Start login

`GET /auth/saml/:orgId/:provider/start`

Redirects the browser to the IdP `entryPoint` with a `SAMLRequest` query param.

### Callback (ACS)

`POST /auth/saml/:orgId/:provider/callback`

Accepts `application/x-www-form-urlencoded` POSTs containing:

- `SAMLResponse` (required)
- `RelayState` (optional)

On success, Formula:

1. Validates response/assertion signatures, issuer/audience, and time conditions.
2. Extracts identity via `attributeMapping` and normalizes email.
3. Links the identity in `user_identities` (`provider = providerId`, `subject = NameID`, `org_id = orgId`).
4. Provisions the user + org membership if needed.
5. Issues a session cookie and writes an `auth.login` audit event (`details: { method: "saml", provider }`).

### `allowed_auth_methods`

SAML login is only allowed if the org has `"saml"` present in `org_settings.allowed_auth_methods`.

## Deployment notes

### PUBLIC_BASE_URL (recommended)

Set `PUBLIC_BASE_URL` (example: `https://app.example.com`) in production so the API can build the ACS URL without relying on request headers.

### Rate limiting

The SAML callback endpoint is rate limited per IP to reduce the impact of replay/brute-force attempts.

