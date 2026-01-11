# ADR-0002: Engine execution model (Tauri invoke vs WASM Worker)

- **Status:** Accepted
- **Date:** 2026-01-10

## Context

We need an execution model that:

- keeps UI responsive (no long-running calc on the main thread),
- works in both **Tauri** and **web** environments,
- and does not force the UI layer to depend on a single IPC mechanism.

Two realistic paths exist:

1. **WASM in a Worker**: UI ↔ Worker (postMessage/RPC). Worker hosts the WASM engine.
2. **Tauri `invoke` commands**: UI ↔ Rust backend via Tauri IPC. Rust backend hosts the engine natively (non-WASM).

Both are useful, but we need a clear rule for when to use which so the architecture doesn’t drift.

## Decision

### 1) Default execution: **WASM engine in a Worker**

The default engine execution model is:

- **Engine runs in a Worker**
- **Engine is compiled to WASM**
- UI communicates via a small typed RPC protocol

This is the portability baseline and is the only option in the browser. Desktop uses the same model to maximize behavioral parity and reduce “works on desktop only” regressions.

### 2) Desktop-only capability path: **Tauri `invoke`**

Tauri `invoke` commands are used for:

- privileged OS integrations (open/save dialogs, filesystem access),
- persistence layers that are desktop-only (SQLite via `rusqlite`, local encryption, keychain),
- background services (auto-update, crash reporting, telemetry),
- and (optionally) performance-critical or thread-heavy computation that cannot be reliably expressed in WASM across all targets.

Critically: **Tauri `invoke` is not a substitute for the engine API.** The UI should call the engine through the same high-level interface on every platform; platform-specific work should be surfaced as “host services”.

## Consequences

- We maintain a single “engine client” abstraction with two implementations:
  - Web: Worker-based engine client
  - Desktop: Worker-based engine client + additional host services exposed via `invoke`
- Shared UI code stays free of direct Tauri imports.
- Features that require desktop privileges must be isolated so web can:
  - degrade gracefully,
  - or be explicitly disabled behind capability checks.

## Current implementation pointers

- Worker engine client boundary (web + desktop): `packages/engine/` (instantiated by `apps/web/`)
- Worker RPC + wasm module loader: `packages/engine/src/worker/EngineWorker.ts` + `packages/engine/src/engine.worker.ts`
- Rust/WASM engine crate (wasm-bindgen): `crates/formula-wasm/`
