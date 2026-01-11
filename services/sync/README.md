# `services/sync` (deprecated)

`services/sync` was an early placeholder WebSocket gateway that validated API-issued sync tokens and then **echoed** messages. Formula collaboration now uses the real Yjs sync server in [`services/sync-server`](../sync-server) (the `y-websocket` protocol) and this legacy service is no longer part of the local stack.

## What to use instead

- Local dev sync server: `pnpm dev:sync` (runs `services/sync-server`)
- Sync server tests: `pnpm test:sync` (runs `services/sync-server` integration tests)
- Docker: `docker-compose.yml` runs `sync-server` on port `1234`

This package remains only for historical context and may be removed in a future cleanup.

