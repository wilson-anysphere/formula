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
4. **Drag and drop:** Files and data
5. **System tray:** Background sync indicator
6. **Global shortcuts:** Capture shortcuts even when unfocused
7. **Notifications:** Native system notifications

### Tauri Security & Permissions (Tauri v2)

This repo is on **Tauri v2**. Config lives in `apps/desktop/src-tauri/tauri.conf.json`.

Key config fields you'll touch most often:

- `build.devUrl`: URL the desktop WebView loads in dev (Vite server)
- `build.frontendDist`: path to built frontend assets for production builds
- `app.security.csp`: Content Security Policy for the desktop WebView
- `plugins.*`: plugin configuration (e.g. updater)

> Note on permissions: Tauri v1 used a `tauri.allowlist` section. **Tauri v2 uses capabilities**
> to grant access to core APIs and plugins. Once we wire capabilities into this repo, the
> definitions should live under `apps/desktop/src-tauri/capabilities/` (Tauri's default location).

### Auto-Update

- Check for updates on startup
- Background download
- User approval before install
- Rollback capability

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

> Tip: depending on `build.beforeDevCommand`, `tauri dev` may start Vite for you.
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
