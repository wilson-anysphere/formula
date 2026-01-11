# SCIM 2.0 Provisioning (Users)

Formula supports a minimal subset of **SCIM 2.0** user provisioning for enterprise IdPs (Okta, Azure AD) via an **org-scoped bearer token**.

Only the **Users** resource is implemented (enough for basic provisioning/deprovisioning).

## Token management (org admin)

SCIM tokens are managed by org admins (owner/admin) via the cloud API:

- `POST /orgs/:orgId/scim/tokens` → create a token (returned once)
- `GET /orgs/:orgId/scim/tokens` → list token metadata (no raw token)
- `DELETE /orgs/:orgId/scim/tokens/:tokenId` → revoke a token

Tokens are stored **hashed** in the database (salted SHA-256); the raw token is only returned at creation time.

Example:

```bash
curl -X POST "https://api.example.com/orgs/$ORG_ID/scim/tokens" \
  -H "Cookie: formula_session=..." \
  -H "Content-Type: application/json" \
  -d '{"name":"okta"}'
```

Response:

```json
{ "id": "…", "name": "okta", "token": "scim_<uuid>.<secret>" }
```

## SCIM authentication

All SCIM endpoints require a SCIM bearer token:

```http
Authorization: Bearer scim_<uuid>.<secret>
```

Requests and responses use `Content-Type: application/scim+json`.

## SCIM base URL

All SCIM routes are mounted under:

`/scim/v2`

## Supported endpoints

### Users

- `GET /scim/v2/Users`
  - Supports `startIndex` (1-based), `count`
  - Minimal filter support: `filter=userName eq "user@example.com"`
- `GET /scim/v2/Users/:id`
- `POST /scim/v2/Users`
  - Creates the user if missing (by `userName` email)
  - Adds org membership
- `PATCH /scim/v2/Users/:id`
  - Supports `active` plus basic name/email updates
  - `active=false` deprovisions by removing org membership
  - `active=true` re-adds org membership
- `DELETE /scim/v2/Users/:id`
  - Removes org membership (does not delete the user record)

All operations are **org-scoped** by the SCIM token; cross-org access is denied.

## Data mapping

- SCIM `id` → `users.id` (UUID)
- SCIM `userName` → `users.email`
- SCIM `name.formatted` (or `givenName` + `familyName`) → `users.name`
- SCIM `active` → org membership (presence of `org_members` row)

## Audit logging

- `org.scim_token.created` / `org.scim_token.revoked`
- `scim.user.created`
- `scim.user.deactivated` / `scim.user.reactivated`
- `scim.user.removed_from_org`

All SCIM provisioning events include `details.source = "scim"`.
