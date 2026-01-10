# ADR-0001: Platform target and portability strategy

- **Status:** Accepted
- **Date:** 2026-01-10

## Context

Formula is being built as an Excel-compatible spreadsheet with a high-performance calculation engine and a modern UI stack (TypeScript/React + Canvas grid). We need to pick an explicit platform strategy early to avoid:

- locking core UI/engine code to a single host (e.g. desktop-only APIs),
- accidentally coupling the engine to Tauri IPC semantics,
- or creating two divergent implementations (desktop vs web) that drift over time.

At the same time, the product requirements strongly favor a native-feeling desktop app (offline-first, large-file performance, OS integrations, installers, auto-update).

## Decision

### 1) Primary platform: **Tauri desktop**

Formula is **desktop-first**. The primary product target is a **Tauri** application for Windows/macOS/Linux.

Reasons:

- native windowing and OS integrations (file dialogs, clipboard, system permissions),
- small bundle and memory profile vs Electron,
- Rust backend for storage/import/export and privileged operations,
- a clear path to distribution and auto-update.

### 2) Secondary platform: **Web build (optional, kept green)**

We will maintain an **optional web build target** that:

- runs in a standard browser,
- is built in CI (`pnpm build:web`),
- is used for development velocity, demos, and long-term optional deployment.

The web target is not required to be feature-complete vs desktop, but it must remain a first-class portability check for the core UI + engine boundary.

### 3) Engine boundary: **Rust → WASM**

The core spreadsheet engine (parse/evaluate, dependency graph, cell storage primitives) is treated as a **platform-agnostic Rust library** that is compiled to **WASM** and executed off the UI thread.

Implications:

- Shared UI code must not import Tauri APIs directly.
- Platform-specific capabilities (filesystem, SQLite, networking permissions) live behind host adapters.
- The UI talks to the engine via a message/RPC boundary rather than direct Rust bindings.

## Long-term plan

1. **Keep the engine portable:** the default execution target is WASM so that web and desktop can share the same deterministic behavior.
2. **Allow native fast-paths on desktop:** desktop may add optional native services (via Tauri `invoke`) for features that are not available or performant enough in the browser.
3. **Keep the web target continuously buildable:** CI must always be able to produce a working `apps/web` build as a guardrail against platform lock-in.

## Consequences

- Shared packages (e.g. grid/components, engine client) must be host-agnostic.
- New features must explicitly decide whether they belong in:
  - the engine (portable, pure, deterministic),
  - the UI (portable),
  - or the host adapter (platform-specific).
- The web target will constrain some early design choices (e.g. avoid synchronous filesystem APIs).

## Current implementation pointers

- Desktop app (Tauri host): `apps/desktop/`
- Web app (Vite/React): `apps/web/` (CI builds via `pnpm build:web`)
- Shared grid renderer package: `packages/grid/`
- Worker-based engine client boundary: `packages/engine/` (initial stub; long-term will load the Rust/WASM engine)
- Rust/WASM engine crate (wasm-bindgen): `crates/formula-wasm/`
