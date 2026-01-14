# Contributing

Thanks for your interest in contributing to Formula!

This repository is early-stage and evolving quickly. Please open an issue before starting large work to avoid duplicate effort.

## Development setup

### Prerequisites

- Node.js (see `.nvmrc` / `.node-version` / `mise.toml`; matches CI/release)
- `pnpm` (see `packageManager` in the root `package.json`)
- Rust toolchain (install via rustup; pinned in `rust-toolchain.toml`)

### Install

```bash
pnpm install
```

### Run the web target

```bash
pnpm dev:web
```

### Desktop perf (startup, memory, size)

To measure desktop shell performance locally (Tauri binary + real WebView), run from the repo root:

```bash
pnpm perf:desktop-startup
pnpm perf:desktop-memory
pnpm perf:desktop-size
```

These commands use an isolated, repo-local HOME (`target/perf-home`) so they don't touch your real user profile.
For details (metrics, tuning knobs, and CI gating env vars), see:

- [`docs/11-desktop-shell.md`](./docs/11-desktop-shell.md)
- [`docs/16-performance-targets.md`](./docs/16-performance-targets.md)

## Repository principles

- **Desktop-first, web-kept-green:** The desktop app (Tauri) is the primary target, but the web build must stay buildable in CI.
- **No platform leaks:** Shared packages must not depend on Tauri-only APIs.
- **Engine runs off the UI thread:** Long-running work belongs in the engine (Worker/WASM), not in React components.

## Pull requests

- Keep PRs focused (one logical change).
- Include screenshots for UI changes.
- Add/update documentation when changing architecture (`docs/adr/*`).

## Code of Conduct

This project follows the Contributor Covenant. See [`CODE_OF_CONDUCT.md`](./CODE_OF_CONDUCT.md).
