# Formula — AI-Native Spreadsheet (Excel-Compatible)

Formula is a next-generation spreadsheet with a **desktop-first** product strategy (Tauri) and an **optional web target** used to keep the core engine/UI portable.

## Platform strategy (high level)

- **Primary target:** Tauri desktop app (Windows/macOS/Linux).
- **Secondary target:** Web build for development, demos, and long-term optional deployment.
- **Core principle:** The calculation engine runs behind a **WASM boundary** and is executed off the UI thread (Worker). Platform-specific integrations live in thin host adapters.

For details, see:
- [`docs/adr/ADR-0001-platform-target.md`](./docs/adr/ADR-0001-platform-target.md)
- [`docs/adr/ADR-0002-engine-execution-model.md`](./docs/adr/ADR-0002-engine-execution-model.md)

## Development

### Web (preview target)

```bash
pnpm install
pnpm dev:web
```

Build:

```bash
pnpm build:web
```

The web build currently renders a placeholder grid and initializes a Worker-based engine stub. The Worker boundary is where the Rust/WASM engine will execute.

## License

Licensed under the Apache License 2.0. See [`LICENSE`](./LICENSE).

