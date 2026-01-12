# Desktop Shell (Tauri v2)

The desktop app is a **Tauri v2.9** shell around the standard web UI. The goal of the Tauri layer is to:

- host the Vite-built UI in a system WebView
- provide native integration (tray, global shortcuts, file open via drag/drop, auto-update)
- expose a small, explicit Rust IPC surface for privileged operations

This document is a “what’s real in the repo” reference for contributors.

## Where the desktop code lives

- **Frontend (TypeScript/Vite):** `apps/desktop/src/`
  - Entry point + desktop host wiring: `apps/desktop/src/main.ts`
- **Tauri (Rust):** `apps/desktop/src-tauri/`
  - Tauri config: `apps/desktop/src-tauri/tauri.conf.json`
  - Entry point: `apps/desktop/src-tauri/src/main.rs`
  - IPC commands: `apps/desktop/src-tauri/src/commands.rs`
  - “Open file” path normalization: `apps/desktop/src-tauri/src/open_file.rs`
  - Tray: `apps/desktop/src-tauri/src/tray.rs`
  - Tray status (icon + tooltip updates): `apps/desktop/src-tauri/src/tray_status.rs`
  - Global shortcuts: `apps/desktop/src-tauri/src/shortcuts.rs`
  - Updater integration: `apps/desktop/src-tauri/src/updater.rs`

---

## Tauri configuration (v2)

The desktop configuration lives in `apps/desktop/src-tauri/tauri.conf.json` (Tauri v2 format).

Key sections you’ll most commonly touch:

### App identity (name/id/version)

Top-level keys in `tauri.conf.json` define the packaged app identity:

- `productName`: human-readable app name
- `identifier`: reverse-DNS bundle identifier (`app.formula.desktop`)
- `version`: desktop app version used by the updater / release tooling (see `docs/release.md`)
- `mainBinaryName`: the Rust binary name Tauri expects to launch (matches `[[bin]].name` in `apps/desktop/src-tauri/Cargo.toml`)

### `build.*` (frontend dev/build + Cargo feature flags)

- `build.beforeDevCommand`: `pnpm dev` (runs Vite)
- `build.beforeBuildCommand`: `pnpm build` (builds `../dist`)
- `build.devUrl`: `http://localhost:4174` (matches `apps/desktop/package.json`)
- `build.frontendDist`: `../dist`
- `build.features: ["desktop"]` enables the Cargo feature gate for the real desktop binary (see “Cargo feature gating” below)

### `app.security.csp`

The CSP is set in `app.security.csp`.

Current policy allows:

- `script-src 'wasm-unsafe-eval' 'unsafe-eval'` (WASM + JS evaluation used by scripting/macro tooling)
- `worker-src 'self' blob:` and `child-src 'self' blob:` (Workers; some bundlers bootstrap via `blob:` URLs)
- `connect-src 'self'` (tight by default; expand intentionally when adding network access)

### `app.windows[].capabilities` (Tauri permissions)

Tauri v2 permissions are granted via **capabilities**, attached to windows in `tauri.conf.json`:

- `apps/desktop/src-tauri/tauri.conf.json` → `app.windows[].capabilities`
- `apps/desktop/src-tauri/capabilities/*.json` (for example: `capabilities/main.json`)

### Cross-origin isolation (COOP/COEP) for Pyodide / `SharedArrayBuffer`

The Pyodide-based Python runtime prefers running in a **Worker** with a `SharedArrayBuffer + Atomics` bridge.
In Chromium/WebView2, that requires a **cross-origin isolated** browsing context:

- `globalThis.crossOriginIsolated === true`
- `typeof SharedArrayBuffer !== "undefined"`

How this is (currently) handled in the repo:

- **Dev / preview (Vite):** `apps/desktop/vite.config.ts` sets COOP/COEP headers on dev/preview responses.
- **Packaged Tauri builds:** COOP/COEP are set via `app.security.headers` in `apps/desktop/src-tauri/tauri.conf.json`,
  which Tauri applies to its built-in `tauri://…` protocol responses.
  If this is missing in a production desktop build, the UI logs an error and shows a long-lived toast (see
  `warnIfMissingCrossOriginIsolationInTauriProd()` in `apps/desktop/src/main.ts`).

Quick verification guidance lives in `apps/desktop/README.md` (“Production/Tauri: `crossOriginIsolated` check”),
including an automated smoke check:

```bash
pnpm -C apps/desktop check:coi
```

Practical warning: with `Cross-Origin-Embedder-Policy: require-corp`, *every* subresource must be same-origin or explicitly opt-in via CORS/CORP.
In Tauri production it’s common to load icons/fonts/images via `asset:`/`asset://…`, which may require adding a `Cross-Origin-Resource-Policy`
header for those responses (or ensuring assets are served from the same origin as the main document).

### `bundle.*` (packaging)

Notable keys:

- `bundle.fileAssociations` registers `.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.csv` with the OS.
- `bundle.linux.deb.depends` documents runtime deps for Linux packaging (e.g. `libwebkit2gtk-4.1-0`, `libgtk-3-0t64 | libgtk-3-0`,
  appindicator, `librsvg2-2`, `libssl3t64 | libssl3`).
- `bundle.macOS.entitlements` / signing keys and `bundle.windows.timestampUrl`.

### `plugins.updater`

Auto-update is configured under `plugins.updater` (Tauri v2 plugin config). In this repo it is enabled, but the public key is intentionally a placeholder:

- `plugins.updater.pubkey` → set this for real releases (see `docs/release.md`)
- `plugins.updater.endpoints` → update JSON endpoint(s)
- `plugins.updater.dialog: false` → the Rust host emits an event instead of showing a built-in dialog

Minimal excerpt (not copy/pasteable; see the full file for everything):

```jsonc
// apps/desktop/src-tauri/tauri.conf.json
{
  "build": {
    "devUrl": "http://localhost:4174",
    "frontendDist": "../dist",
    "features": ["desktop"]
  },
  "app": {
    "security": {
      "csp": "default-src 'self'; ...; worker-src 'self' blob:; ...;"
    },
    "windows": [
      { "label": "main", "title": "Formula", "width": 1280, "height": 800, "dragDropEnabled": true, "capabilities": ["main"] }
    ]
  },
  "bundle": {
    "fileAssociations": [{ "ext": ["xlsx"], "name": "Excel Spreadsheet", "role": "Editor" }]
  },
  "plugins": {
    "updater": { "active": true, "dialog": false, "pubkey": "REPLACE_WITH_TAURI_UPDATER_PUBLIC_KEY" }
  }
}
```

---

## Rust host (Tauri backend)

### Entry point: `apps/desktop/src-tauri/src/main.rs`

`main.rs` wires together:

- **state** (`SharedAppState`) + **macro trust store** (`SharedMacroTrustStore`)
- Tauri plugins:
  - `tauri_plugin_global_shortcut` (registers accelerators + emits app events)
  - `tauri_plugin_updater` (update checks)
  - `tauri_plugin_single_instance` (forward argv/cwd from subsequent launches into the running instance)
- `invoke_handler(...)` mapping commands in `commands.rs`
- window/tray event forwarding to the frontend via `app.emit(...)` / `window.emit(...)`

#### Close flow (hide vs quit)

The desktop app deliberately **does not exit on window close** so the tray remains available.

The window-close sequence is:

1. Rust receives `WindowEvent::CloseRequested` and calls `api.prevent_close()`.
2. Rust emits `close-prep` with a random token.
3. Frontend (in `apps/desktop/src/main.ts`) flushes pending workbook sync and calls `set_macro_ui_context`, then emits `close-prep-done` with the same token.
4. Rust runs a best-effort `Workbook_BeforeClose` macro (if trusted) and collects any cell updates.
5. Rust emits `close-requested` with `{ token, updates }`.
6. Frontend applies any macro cell updates, prompts for unsaved changes if needed, then either:
   - hides the window (default behavior; app keeps running in the tray), or
   - keeps the window open if the user cancels the close (e.g. cancels the unsaved-changes prompt)
7. Frontend emits `close-handled` with the token so Rust can clear its “close in flight” guard.

Implementation detail: `main.rs` uses an `AtomicBool` (`CLOSE_REQUEST_IN_FLIGHT`) to prevent overlapping close flows if the user clicks close repeatedly while a prompt is still open.

#### Drag & drop → open file

When a file is dropped onto the window, `main.rs` listens for `WindowEvent::DragDrop` and emits:

- `file-dropped` with `Vec<String>` of filesystem paths

The frontend listens for this event and queues an open via `queueOpenWorkbook(...)` (so opens are serialized).

#### Open-with / file associations / CLI args

In addition to drag & drop, the desktop shell supports opening workbooks via:

- “Open with…” / Finder / Explorer (file associations configured in `bundle.fileAssociations` in `tauri.conf.json`)
- passing a path on the command line (cold start)
- launching the app again while an instance is already running (warm start)

Implementation notes:

- `apps/desktop/src-tauri/src/open_file.rs` extracts supported spreadsheet paths from argv-style inputs (and also supports `file://...` URLs used by macOS open-document events).
- `main.rs` uses a small in-memory queue (`OpenFileState`) so open-file requests received *before* the frontend installs its listeners aren’t lost.
  - Backend emits: `open-file` (payload: `string[]` paths)
  - Frontend emits: `open-file-ready` once its `listen("open-file", ...)` handler is installed, which flushes any queued paths.
- When an open-file request is handled, `main.rs` **shows + focuses** the main window before emitting `open-file` so the request is visible to the user.
- On macOS, `tauri::RunEvent::Opened { urls, .. }` is routed through the same pipeline so opening a document in Finder reaches the running instance.

#### Tray + global shortcuts

- Tray menu and click behavior are implemented in `apps/desktop/src-tauri/src/tray.rs`.
  - Emits: `tray-new`, `tray-open`, `tray-quit`
  - “Check for Updates” runs an update check (`updater::spawn_update_check(..., UpdateCheckSource::Manual)`)
- In release builds, `main.rs` also runs a lightweight update check on startup (`updater::spawn_update_check(..., UpdateCheckSource::Startup)`, behind `#[cfg(not(debug_assertions))]`).
- Global shortcuts are registered in `apps/desktop/src-tauri/src/shortcuts.rs`.
  - Accelerators: `CmdOrCtrl+Shift+O`, `CmdOrCtrl+Shift+P`
  - The plugin handler in `main.rs` emits: `shortcut-quick-open`, `shortcut-command-palette`

Note on quitting from the tray:

- The Rust host emits `tray-quit`, but it does **not** hard-exit immediately.
- The frontend handles `tray-quit` by running its quit flow (best-effort `Workbook_BeforeClose`, unsaved changes prompt) and finally invoking the `quit_app` command to exit the process.

---

## Frontend host wiring (`apps/desktop/src/main.ts`)

The desktop UI intentionally avoids a hard dependency on `@tauri-apps/api` and instead uses the injected runtime object:

- `globalThis.__TAURI__.core.invoke` for `#[tauri::command]` calls
- `globalThis.__TAURI__.event.listen` / `emit` for events
- `globalThis.__TAURI__.window.*` for hiding the window
- `globalThis.__TAURI__.dialog.open/save` for file open/save prompts

Desktop-specific listeners are set up near the bottom of `apps/desktop/src/main.ts`:

- `close-prep` → flush pending workbook sync + call `set_macro_ui_context` → emit `close-prep-done`
- `close-requested` → run `handleCloseRequest(...)` (unsaved changes prompt + hide vs quit) → emit `close-handled`
- `open-file` → queue workbook opens; then emits `open-file-ready` once the handler is installed (flushes any queued open-file requests on the Rust side)
- `file-dropped` → open the first dropped path
- `tray-open` / `tray-new` / `tray-quit` → open dialog/new workbook/quit flow
- `shortcut-quick-open` / `shortcut-command-palette` → open dialog/palette

Important implementation detail: invoke calls are serialized via `queueBackendOp(...)` / `pendingBackendSync` so that bulk edits (workbook sync) don’t race with open/save/close.

---

## Desktop IPC surface

### Commands (`#[tauri::command]` in `apps/desktop/src-tauri/src/commands.rs`)

The command list is large; below are the “core” ones most contributors will interact with (not exhaustive):

- **Workbook lifecycle**
  - `open_workbook`, `new_workbook`, `save_workbook`, `mark_saved`, `add_sheet`
- **Cells / ranges / recalculation**
  - `get_cell`, `set_cell`, `get_range`, `set_range`, `recalculate`, `undo`, `redo`
  - Dependency inspection: `get_precedents`, `get_dependents`
  - Sheet bounds: `get_sheet_used_range`
- **Workbook metadata (used by UI + Power Query + AI tooling)**
  - `get_workbook_theme_palette`, `list_defined_names`, `list_tables`
- **Pivot tables**
  - `create_pivot_table`, `refresh_pivot_table`, `list_pivot_tables`
- **Printing / export**
  - `get_sheet_print_settings`, `set_sheet_page_setup`, `set_sheet_print_area`, `export_sheet_range_pdf`
- **Local file access for Power Query sources (instead of Tauri FS plugin)**
  - `read_text_file`, `read_binary_file`, `read_binary_file_range`, `stat_file`, `list_dir`
- **Power Query secure storage + refresh state**
  - `power_query_cache_key_get_or_create`
  - `power_query_credential_get|set|delete|list`
  - `power_query_refresh_state_get|set`
  - `power_query_state_get|set`
- **SQL (connectors / queries)**
  - `sql_query`, `sql_get_schema`
- **Macros + scripting**
  - Macro inspection/security: `get_vba_project`, `list_macros`, `get_macro_security_status`, `set_macro_trust`
  - Execution/context: `set_macro_ui_context`, `run_macro`, `validate_vba_migration`
  - Python: `run_python_script`
  - VBA event hooks: `fire_workbook_open`, `fire_workbook_before_close`, `fire_worksheet_change`, `fire_selection_change`
- **Lifecycle**
  - `quit_app` (hard-exits the process; used by the tray quit flow)
  - `exit_process` (hard-exits with a given code; used by `pnpm -C apps/desktop check:coi`)
  - `report_cross_origin_isolation` (logs `crossOriginIsolated` + `SharedArrayBuffer` status for the COOP/COEP smoke check)
- **Tray integration**
  - `set_tray_status` (update tray icon + tooltip for simple statuses: `idle`, `syncing`, `error`)

### Backend → frontend events

Events emitted by the Rust host (see `main.rs`, `tray.rs`, `updater.rs`):

- Window lifecycle:
  - `close-prep` (payload: token `string`)
  - `close-requested` (payload: `{ token: string, updates: CellUpdate[] }`)
  - `open-file` (payload: `string[]` paths)
  - `file-dropped` (payload: `string[]` paths)
- Tray:
  - `tray-new`, `tray-open`, `tray-quit`
- Shortcuts:
  - `shortcut-quick-open`, `shortcut-command-palette`
- Updates:
  - `update-check-started` (payload: `{ source }`)
  - `update-not-available` (payload: `{ source }`)
  - `update-check-error` (payload: `{ source, message }`)
  - `update-available` (payload: `{ source, version, body }`)

Note: at the time of writing, `apps/desktop/src/main.ts` does not yet listen for these updater events (they are emitted by the Rust host, but there is no UI flow wired up yet).

Related frontend → backend events used as acknowledgements during close:

- `close-prep-done` (token)
- `close-handled` (token)
- `open-file-ready` (signals that the frontend’s `open-file` listener is installed; causes the Rust host to flush queued open requests)

---

## Cargo feature gating (`desktop`)

The Tauri **binary** is feature-gated so that backend unit tests can run without system WebView
dependencies (notably GTK/WebKit on Linux).

Where it’s defined:

- `apps/desktop/src-tauri/Cargo.toml`
  - The desktop binary (`[[bin]]`) has `required-features = ["desktop"]`.
  - The `desktop` feature enables the optional deps: `tauri`, `tauri-build`, and the desktop-only Tauri plugins
    (currently `tauri-plugin-global-shortcut`, `tauri-plugin-updater`, `tauri-plugin-single-instance`).
- `apps/desktop/src-tauri/tauri.conf.json`
  - `build.features: ["desktop"]` ensures `tauri dev` / `tauri build` compiles the real desktop binary with the correct feature set.

Practical effect:

- `bash scripts/cargo_agent.sh test -p desktop` can run in CI without installing WebView toolchains.
- Building/running the app uses the `desktop` feature and therefore requires the platform WebView dependencies.

Note: most `#[tauri::command]` functions in `apps/desktop/src-tauri/src/commands.rs` are also `#[cfg(feature = "desktop")]`, so the
backend library can still compile (and be tested) without linking Tauri or system WebView components.

---

## Permissions: from “allowlist” (v1) to “capabilities” (v2)

Tauri v1 used an `allowlist` section in config. **Tauri v2 replaces this with window-scoped “capabilities”**:

- Capabilities define which **core APIs / plugin APIs** are available to which windows.
- They are intended to be **narrow and explicit** (e.g. “main window can show file-open dialog” vs “all APIs enabled”).

This repo’s desktop shell currently relies primarily on:

- explicit Rust `#[tauri::command]` functions in `apps/desktop/src-tauri/src/commands.rs` for privileged operations
- a small set of runtime JS APIs (`__TAURI__.event`, `__TAURI__.window`, `__TAURI__.dialog`)

Capabilities live alongside the Tauri app under:

- `apps/desktop/src-tauri/capabilities/` (see `main.json`)

The `tauri.conf.json` window config references capabilities via `app.windows[].capabilities`.

---

## Release, signing, updater keys

The updater config (`plugins.updater.*`) is in `apps/desktop/src-tauri/tauri.conf.json`.

For the actual release workflow, signing, and updater key management, see:

- `docs/release.md`
