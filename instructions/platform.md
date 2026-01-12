# Workstream F: Platform

> **⛔ STOP. READ [`AGENTS.md`](../AGENTS.md) FIRST. FOLLOW IT COMPLETELY. THIS IS NOT OPTIONAL. ⛔**
>
> This document is supplementary to AGENTS.md. All rules, constraints, and guidelines in AGENTS.md apply to you at all times. Memory limits, build commands, design philosophy—everything.

---

## Mission

Build the native desktop application shell using **Tauri**. Handle system integration, distribution, auto-updates, and platform-specific features.

**The goal:** 10x smaller than Electron, 4-8x less memory, <500ms startup.

---

## Scope

### Your Code

| Location | Purpose |
|----------|---------|
| `apps/desktop/src-tauri/` | Rust backend, Tauri commands |
| `apps/desktop/src-tauri/capabilities/` | Tauri v2 capability allowlists (permissions for core APIs/plugins) |
| `apps/desktop/src/tauri/` | TypeScript Tauri bindings |
| `apps/desktop/` | Desktop app entrypoint, Vite config |
| `apps/web/` | Web build target (secondary) |

### Your Documentation

- **Primary:** [`docs/11-desktop-shell.md`](../docs/11-desktop-shell.md) — Tauri integration, native features, distribution

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│  DESKTOP APPLICATION                                            │
├─────────────────────────────────────────────────────────────────┤
│  ┌───────────────────────────────────────────────────────────┐  │
│  │  WEBVIEW (System WebView)                                  │  │
│  │  ├── React UI Components                                   │  │
│  │  ├── Canvas Grid Renderer                                  │  │
│  │  └── TypeScript Application Logic                          │  │
│  └─────────────────────────┬─────────────────────────────────┘  │
│                            │ Tauri IPC                          │
│  ┌─────────────────────────▼─────────────────────────────────┐  │
│  │  RUST BACKEND                                              │  │
│  │  ├── Calculation Engine (WASM in Worker)                   │  │
│  │  ├── File I/O (async)                                      │  │
│  │  ├── SQLite Database                                       │  │
│  │  ├── System Integration                                    │  │
│  │  └── Native Dialogs                                        │  │
│  └───────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────┘
```

---

## Key Requirements

### System Integration

1. **File associations:** Double-click `.xlsx` opens Formula
2. **Native dialogs:** Open/save file dialogs
3. **Clipboard:** Rich clipboard support (HTML, RTF, images)
   - Frontend entry point: `apps/desktop/src/clipboard/platform/provider.js`
   - Desktop prefers custom Rust IPC commands `clipboard_read` / `clipboard_write` (multi-format),
     with fallbacks to `navigator.clipboard` and legacy `globalThis.__TAURI__.clipboard.readText` / `writeText`.
   - Supported formats: `text/plain`, `text/html`, `text/rtf`, `image/png`.
   - PNG over IPC is transported as `pngBase64` (raw base64, no `data:image/png;base64,` prefix).
4. **Drag and drop:** Files and data
5. **System tray:** Background sync indicator
6. **Global shortcuts:** Capture shortcuts even when unfocused
7. **Notifications:** Native system notifications

### Tauri v2 Capabilities & Permissions

This repo is on **Tauri v2**. Config lives in `apps/desktop/src-tauri/tauri.conf.json`.

Key config fields you'll touch most often:

- `build.devUrl`: URL the desktop WebView loads in dev (Vite server)
- `build.frontendDist`: path to built frontend assets for production builds
- `app.security.headers`: COOP/COEP headers (required for `crossOriginIsolated` / `SharedArrayBuffer`)
- `app.security.csp`: Content Security Policy for the desktop WebView
- `plugins.*`: plugin configuration (e.g. updater)

`app.security.headers` is especially important for the desktop app because the Pyodide-based Python runtime
prefers running in a Worker with `SharedArrayBuffer` (requires `crossOriginIsolated === true`).
See `docs/11-desktop-shell.md` for details.

Tauri v2 permissions are granted via **capabilities**:

- `apps/desktop/src-tauri/capabilities/*.json`
- capability files scope themselves to window labels via the capability file’s `"windows": [...]` list (matches `app.windows[].label` in `apps/desktop/src-tauri/tauri.conf.json`)
- windows opt into capabilities via `app.windows[].capabilities` (list of capability identifiers; the main window uses `["main"]`)

Example excerpt (see `apps/desktop/src-tauri/capabilities/main.json` for the full allowlists):

```jsonc
{
  "$schema": "../gen/schemas/desktop-schema.json",
  "identifier": "main",
  "description": "Permissions for the main desktop window (explicit IPC allowlist).",
  "local": true,
  "windows": ["main"],
  "permissions": [
    "core:default",
    { "identifier": "core:allow-invoke", "allow": ["open_workbook", "save_workbook"] },
    // ...
    { "identifier": "event:allow-listen", "allow": [{ "event": "open-file" }] },
    { "identifier": "event:allow-emit", "allow": [{ "event": "open-file-ready" }] },
    // ...
    "dialog:allow-open",
    "dialog:allow-save",
    "window:allow-hide",
    "window:allow-show",
    "window:allow-set-focus",
    "window:allow-close",
    "clipboard:allow-read-text",
    "clipboard:allow-write-text",
    "shell:allow-open",
    "updater:allow-check",
    "updater:allow-download",
    "updater:allow-install"
  ]
}
```

Note: `clipboard:allow-read-text` / `clipboard:allow-write-text` only grant access to the legacy plain-text
clipboard helpers (`globalThis.__TAURI__.clipboard.readText` / `writeText`). Rich clipboard formats
(HTML/RTF/PNG) are handled via custom Rust commands and must be added to the `core:allow-invoke` allowlist
when used.

If you add new desktop IPC surface area, you must update the capability allowlists:

- new `#[tauri::command]` names → `core:allow-invoke`
- new frontend↔backend events → `event:allow-listen` / `event:allow-emit`

We keep guardrail tests to ensure we don't accidentally broaden the desktop IPC surface:

- **Event allowlists**: enforce the **exact** `event:allow-listen` / `event:allow-emit` sets (no wildcard / allow-all):
  - `apps/desktop/src/tauri/__tests__/eventPermissions.vitest.ts`
- **Invoke allowlist**: ensure `core:allow-invoke` matches the Rust `generate_handler![...]` registration list:
  - `apps/desktop/src-tauri/tests/tauri_ipc_allowlist.rs`
- **Updater permissions**: ensure the frontend restart/install flow has the minimum updater permissions:
  - `apps/desktop/src/tauri/__tests__/capabilitiesPermissions.vitest.ts`

Filesystem access for Power Query is handled via **custom Rust commands** (e.g. `read_text_file`, `list_dir`)
instead of the optional Tauri FS plugin. Those commands enforce an explicit scope:

- `$HOME/**`
- `$DOCUMENT/**`

The scope check uses canonicalization to normalize paths and prevent symlink-based escapes.

### Auto-Update

- Check for updates on startup
- Background download
- User approval before install
- Rollback capability
  - Tauri does not provide an automatic “revert to previous version” after a successful upgrade.
  - Formula supports a clear **manual downgrade path** via the GitHub Releases page (in-app via the
    updater dialog’s “Open release page” / “Download manually” action).
  - Rollback depends on keeping older release assets available (don’t delete prior releases). See
    `docs/11-desktop-shell.md` and `docs/release.md`.

### Distribution

- **macOS:** `.dmg`, notarized and stapled
- **Windows:** `.msi` and `.exe` installers, signed
- **Linux:** `.AppImage`, `.deb`, `.rpm`

---

## Performance Targets

| Metric | Target |
|--------|--------|
| Cold start | <500ms to window visible |
| Time to interactive | <1 second |
| Memory (idle) | <100MB |
| Bundle size | <50MB |

---

## Build & Run

### Development

```bash
# Initialize safe defaults (required for agents)
source scripts/agent-init.sh

# Install dependencies
pnpm install

# Build WASM engine first
pnpm build:wasm
```

Run the desktop frontend (Vite):

```bash
# Vite dev server (matches `build.devUrl` in `apps/desktop/src-tauri/tauri.conf.json`)
pnpm dev:desktop
```

Run the native desktop shell (Tauri):

```bash
# ALWAYS use the cargo wrapper (see `AGENTS.md`)
cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri dev
```

> Tip: depending on `build.beforeDevCommand`, `cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri dev` may start Vite for you.
> Avoid running two dev servers on the same port.

### Production Build

```bash
# Initialize safe defaults (required for agents)
source scripts/agent-init.sh

# Build web assets
pnpm build:desktop
```

Build the native app:

```bash
# ALWAYS use the cargo wrapper (see `AGENTS.md`)
cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build
```

### Headless Testing

```bash
# Tauri can build without display
cd apps/desktop && bash ../../scripts/cargo_agent.sh tauri build

# Dev server needs virtual display
cd apps/desktop && xvfb-run --auto-servernum bash ../../scripts/cargo_agent.sh tauri dev
```

---

## Tauri Commands (IPC)

TypeScript ↔ Rust communication:

```rust
// Rust side (src-tauri/src/lib.rs)
#[tauri::command]
async fn read_file(path: String) -> Result<Vec<u8>, String> {
    std::fs::read(&path).map_err(|e| e.to_string())
}

#[tauri::command]
async fn save_file(path: String, data: Vec<u8>) -> Result<(), String> {
    std::fs::write(&path, data).map_err(|e| e.to_string())
}
```

```typescript
// TypeScript side
type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;
const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
if (!invoke) throw new Error("Tauri invoke API not available");

const data = await invoke("read_file", { path: "/path/to/file.xlsx" });
await invoke("save_file", { path: "/path/to/file.xlsx", data: newData });
```

---

## Coordination Points

- **UI Team:** Window management, native dialogs, menu bar
- **File I/O Team:** Native file system access
- **Collaboration Team:** Background sync, system tray status

---

## Platform-Specific Notes

### macOS

- Use `entitlements.plist` for sandboxing
- Notarize for Gatekeeper
- Support Apple Silicon (aarch64) and Intel (x86_64)
- Use `.icns` icon format

### Windows

- Sign with code signing certificate
- Support both x64 and arm64
- Use `.ico` icon format
- Handle UAC elevation if needed

### Linux

- AppImage for universal compatibility
- Respect XDG directories
- Handle Wayland and X11

---

## Reference

- Tauri documentation: https://tauri.app/
- Tauri v2: https://tauri.app/v2/
- wry (WebView): https://github.com/nicklockwood/wry
