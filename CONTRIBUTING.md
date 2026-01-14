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

## Formula engine generators (common gotchas)

If you add/rename a built-in Excel-like function in `crates/formula-engine`, there are a few
generated artifacts that must stay in sync:

- Function catalog (`shared/functionCatalog.json`):
  - `pnpm generate:function-catalog`
- Locale function-name translation TSVs (`crates/formula-engine/src/locale/data/*.tsv`):
  - Normalize locale sources (omits identity mappings + enforces stable casing):
    - `pnpm normalize:locale-function-sources`
    - `pnpm check:locale-function-sources`
  - `pnpm generate:locale-function-tsv`
  - `pnpm check:locale-function-tsv`
  - For Spanish (`es-ES`), locale sources must come from a full-catalog Excel extraction (do not
    replace with partial online translation tables); see
    [`crates/formula-engine/src/locale/data/README.md`](./crates/formula-engine/src/locale/data/README.md).

## Code of Conduct

This project follows the Contributor Covenant. See [`CODE_OF_CONDUCT.md`](./CODE_OF_CONDUCT.md).
