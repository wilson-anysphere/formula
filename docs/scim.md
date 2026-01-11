# SCIM 2.0 Provisioning (Enterprise)

Formula supports a minimal SCIM 2.0 subset for automated user provisioning into an organization.

## Authentication (SCIM token)

SCIM endpoints **only** accept a SCIM bearer token (sessions / API keys are not valid).

### Create / rotate token (org admin)

`POST /orgs/:orgId/scim/token`

- Requires an org admin (owner/admin) session.
- Returns the **plaintext token once**; it is never persisted server-side.
- Calling this endpoint again rotates the token (old token becomes invalid).

Response:

```json
{ "token": "scim_<orgId>.<secret>" }
```

### Revoke token (org admin)

`DELETE /orgs/:orgId/scim/token`

Revokes the current token immediately.

## SCIM base URL

All SCIM routes live under:

`/scim/v2`

Use the token as:

`Authorization: Bearer scim_<orgId>.<secret>`

## Supported endpoints

### List users

`GET /scim/v2/Users`

Supports:

- `startIndex` (1-based)
- `count`
- `filter` subset: `userName eq "user@example.com"`

Only users that are **active members** of the org are returned.

### Create user (idempotent by email)

`POST /scim/v2/Users`

- Creates a new `users` row if the email does not exist.
- Ensures `org_members` exists when `active !== false`.

### Get user

`GET /scim/v2/Users/:id`

Returns the user only if they are an active member of the org.

### Patch user

`PATCH /scim/v2/Users/:id`

Supported updates:

- `active` (boolean)
  - `active=false` removes org membership (`org_members` row deleted)
  - `active=true` adds org membership (`org_members` row inserted)
- `displayName` (string) updates `users.name`

## Data mapping

- SCIM `id` → `users.id` (UUID)
- SCIM `userName` + primary email → `users.email`
- SCIM `displayName` → `users.name`
- SCIM `active` → org membership (presence of `org_members` row)

## Audit logging

Provisioning actions are recorded in `audit_log`:

- `org.scim.token_created` / `org.scim.token_revoked`
- `admin.user_created`
- `admin.user_deactivated` / `admin.user_reactivated`

All provisioning events include `details.source = "scim"`.

