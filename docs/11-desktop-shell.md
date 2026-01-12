# Desktop Application Shell

## Overview

The desktop application uses **Tauri** for a native shell with a Rust backend. This provides 10x smaller bundles than Electron, 4-8x less memory usage, and sub-500ms startup times while maintaining cross-platform compatibility.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  DESKTOP APPLICATION                                                        │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│  ┌─────────────────────────────────────────────────────────────────────┐   │
│  │  WEBVIEW (System WebView)                                            │   │
│  │  ├── React UI Components                                             │   │
│  │  ├── Canvas Grid Renderer                                            │   │
│  │  └── TypeScript Application Logic                                    │   │
│  └─────────────────────────────┬───────────────────────────────────────┘   │
│                                │                                            │
│                                │ Tauri IPC                                  │
│                                │                                            │
│  ┌─────────────────────────────▼───────────────────────────────────────┐   │
│  │  RUST BACKEND                                                        │   │
│  │  ├── Calculation Engine (multi-threaded)                            │   │
│  │  ├── File I/O (async)                                                │   │
│  │  ├── SQLite Database                                                 │   │
│  │  ├── System Integration                                              │   │
│  │  └── Native Dialogs                                                  │   │
│  └─────────────────────────────────────────────────────────────────────┘   │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## Tauri Configuration

### tauri.conf.json

```json
{
  "build": {
    "beforeBuildCommand": "npm run build",
    "beforeDevCommand": "npm run dev",
    "devPath": "http://localhost:5173",
    "distDir": "../dist"
  },
  
  "package": {
    "productName": "Formula",
    "version": "1.0.0"
  },
  
  "tauri": {
    "allowlist": {
      "all": false,
      "shell": {
        "all": false,
        "open": true
      },
      "dialog": {
        "all": true
      },
      "fs": {
        "all": false,
        "readFile": true,
        "writeFile": true,
        "readDir": true,
        "exists": true,
        "scope": ["$DOCUMENT/**", "$DOWNLOAD/**", "$HOME/**"]
      },
      "path": {
        "all": true
      },
      "clipboard": {
        "all": true
      },
      "window": {
        "all": true
      },
      "globalShortcut": {
        "all": true
      },
      "notification": {
        "all": true
      }
    },
    
    "bundle": {
      "active": true,
      "icon": [
        "icons/32x32.png",
        "icons/128x128.png",
        "icons/128x128@2x.png",
        "icons/icon.icns",
        "icons/icon.ico"
      ],
      "identifier": "app.formula.desktop",
      "targets": "all",
      "category": "Productivity",
      
      "macOS": {
        "entitlements": "entitlements.plist",
        "signingIdentity": "Developer ID Application",
        "minimumSystemVersion": "10.15"
      },
      
      "windows": {
        "certificateThumbprint": null,
        "timestampUrl": "http://timestamp.digicert.com"
      },
      
      "linux": {
        "deb": {
          "depends": [
            "libwebkit2gtk-4.1-0",
            "libgtk-3-0",
            "libayatana-appindicator3-1 | libappindicator3-1",
            "librsvg2-2",
            "libssl3"
          ]
        }
      }
    },
    
    "security": {
      "dangerousDisableAssetCspModification": false,
      "csp": "default-src 'self'; base-uri 'self'; object-src 'none'; frame-ancestors 'none'; script-src 'self' 'wasm-unsafe-eval' 'unsafe-eval'; worker-src 'self' blob:; child-src 'self' blob:; style-src 'self' 'unsafe-inline'; img-src 'self' asset: data: https:; connect-src 'self' https://api.formula.app wss://sync.formula.app"
    },
    
    "windows": [
      {
        "title": "Formula",
        "width": 1280,
        "height": 800,
        "minWidth": 800,
        "minHeight": 600,
        "resizable": true,
        "fullscreen": false,
        "decorations": true,
        "transparent": false,
        "center": true,
        "fileDropEnabled": true
      }
    ],
    
    "systemTray": {
      "iconPath": "icons/tray.png",
      "iconAsTemplate": true
    }
  }
}
```

#### CSP notes (WASM engine + Workers)

- The Rust engine runs as **WebAssembly inside a module Worker**, so CSP must allow:
  - `script-src 'wasm-unsafe-eval'` for WASM compilation/instantiation.
  - `worker-src 'self' blob:` for module workers (Vite may use `blob:` URLs for worker bootstrapping).
- Some WebKit-based WebViews historically require `script-src 'unsafe-eval'` for WASM.
  We also rely on `unsafe-eval` for the TypeScript scripting sandbox (it evaluates compiled code via `new Function` inside a Worker).
- Older WebKit versions may still gate worker creation behind `child-src`, so we include `child-src 'self' blob:` as a fallback.

---

## Rust Backend

### Main Entry Point

```rust
// src-tauri/src/main.rs

#![cfg_attr(
    all(not(debug_assertions), target_os = "windows"),
    windows_subsystem = "windows"
)]

mod calc_engine;
mod commands;
mod file_io;
mod state;

use state::AppState;
use std::sync::Mutex;
use tauri::Manager;

fn main() {
    tauri::Builder::default()
        .manage(Mutex::new(AppState::new()))
        .invoke_handler(tauri::generate_handler![
            commands::open_workbook,
            commands::save_workbook,
            commands::get_cell,
            commands::set_cell,
            commands::get_range,
            commands::set_range,
            commands::recalculate,
            commands::undo,
            commands::redo,
            commands::copy,
            commands::paste,
            commands::find_replace,
        ])
        .setup(|app| {
            // Initialize logging
            env_logger::init();
            
            // Set up file associations
            #[cfg(target_os = "macos")]
            {
                use tauri::api::path::document_dir;
                // Handle file open events
            }
            
            // Set up global shortcuts
            let handle = app.handle();
            tauri::async_runtime::spawn(async move {
                setup_shortcuts(handle).await;
            });
            
            Ok(())
         })
         .on_window_event(|event| match event.event() {
             tauri::WindowEvent::CloseRequested { api, .. } => {
                // Delegate close handling to the frontend (Workbook_BeforeClose,
                // unsaved changes prompt, hide vs keep open).
                api.prevent_close();
                event.window().emit("close-requested", ()).unwrap();
             }
             tauri::WindowEvent::FileDrop(event) => {
                 if let tauri::FileDropEvent::Dropped(paths) = event {
                     for path in paths {
                        event.window().emit("file-dropped", path).unwrap();
                    }
                }
            }
            _ => {}
        })
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
```

### Tauri Commands

```rust
// src-tauri/src/commands.rs

use crate::calc_engine::CalcEngine;
use crate::file_io::{read_xlsx, write_xlsx};
use crate::state::AppState;
use serde::{Deserialize, Serialize};
use std::sync::Mutex;
use tauri::State;

#[derive(Serialize, Deserialize)]
pub struct CellValue {
    value: Option<serde_json::Value>,
    formula: Option<String>,
    display_value: String,
}

#[derive(Serialize, Deserialize)]
pub struct RangeData {
    values: Vec<Vec<CellValue>>,
    start_row: usize,
    start_col: usize,
}

#[tauri::command]
pub async fn open_workbook(
    path: String,
    state: State<'_, Mutex<AppState>>,
) -> Result<WorkbookInfo, String> {
    let workbook = read_xlsx(&path).await.map_err(|e| e.to_string())?;
    
    let mut state = state.lock().unwrap();
    let info = state.load_workbook(workbook);
    
    Ok(info)
}

#[tauri::command]
pub async fn save_workbook(
    path: Option<String>,
    state: State<'_, Mutex<AppState>>,
) -> Result<(), String> {
    let state = state.lock().unwrap();
    let workbook = state.get_workbook().ok_or("No workbook loaded")?;
    
    let save_path = path.unwrap_or_else(|| {
        workbook.path.clone().expect("Workbook has no path")
    });
    
    write_xlsx(&save_path, workbook).await.map_err(|e| e.to_string())
}

#[tauri::command]
pub fn get_cell(
    sheet_id: String,
    row: usize,
    col: usize,
    state: State<'_, Mutex<AppState>>,
) -> Result<CellValue, String> {
    let state = state.lock().unwrap();
    let cell = state.get_cell(&sheet_id, row, col)?;
    
    Ok(CellValue {
        value: cell.value.clone(),
        formula: cell.formula.clone(),
        display_value: cell.format_display_value(),
    })
}

#[tauri::command]
pub fn set_cell(
    sheet_id: String,
    row: usize,
    col: usize,
    value: Option<serde_json::Value>,
    formula: Option<String>,
    state: State<'_, Mutex<AppState>>,
) -> Result<Vec<CellUpdate>, String> {
    let mut state = state.lock().unwrap();
    
    // Set cell value
    state.set_cell(&sheet_id, row, col, value, formula)?;
    
    // Recalculate affected cells
    let updates = state.recalculate_from(&sheet_id, row, col)?;
    
    Ok(updates)
}

#[tauri::command]
pub fn get_range(
    sheet_id: String,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    state: State<'_, Mutex<AppState>>,
) -> Result<RangeData, String> {
    let state = state.lock().unwrap();
    let cells = state.get_range(&sheet_id, start_row, start_col, end_row, end_col)?;
    
    Ok(RangeData {
        values: cells,
        start_row,
        start_col,
    })
}

#[tauri::command]
pub async fn recalculate(
    state: State<'_, Mutex<AppState>>,
) -> Result<Vec<CellUpdate>, String> {
    // Spawn blocking task for CPU-intensive calculation
    let state_clone = state.inner().clone();
    
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = state_clone.lock().unwrap();
        state.recalculate_all()
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
pub fn undo(state: State<'_, Mutex<AppState>>) -> Result<UndoResult, String> {
    let mut state = state.lock().unwrap();
    state.undo()
}

#[tauri::command]
pub fn redo(state: State<'_, Mutex<AppState>>) -> Result<UndoResult, String> {
    let mut state = state.lock().unwrap();
    state.redo()
}
```

### Calculation Engine (Rust)

```rust
// src-tauri/src/calc_engine.rs

use rayon::prelude::*;
use std::collections::{HashMap, HashSet};

pub struct CalcEngine {
    cells: HashMap<CellId, Cell>,
    dependency_graph: DependencyGraph,
    dirty_cells: HashSet<CellId>,
}

impl CalcEngine {
    pub fn new() -> Self {
        Self {
            cells: HashMap::new(),
            dependency_graph: DependencyGraph::new(),
            dirty_cells: HashSet::new(),
        }
    }
    
    pub fn set_cell(&mut self, cell_id: CellId, formula: Option<String>, value: CellValue) {
        // Parse formula and update dependencies
        if let Some(ref formula_str) = formula {
            let ast = self.parse_formula(formula_str);
            let deps = self.extract_dependencies(&ast);
            self.dependency_graph.update(cell_id, deps);
        }
        
        // Store cell
        let cell = Cell { formula, value, ast: None };
        self.cells.insert(cell_id, cell);
        
        // Mark dependents dirty
        self.mark_dirty(cell_id);
    }
    
    pub fn recalculate(&mut self) -> Vec<CellUpdate> {
        // Get calculation order
        let calc_order = self.dependency_graph.topological_sort(&self.dirty_cells);
        
        // Partition into independent subgraphs for parallel execution
        let subgraphs = self.partition_independent_subgraphs(&calc_order);
        
        // Calculate in parallel
        let updates: Vec<CellUpdate> = subgraphs
            .into_par_iter()
            .flat_map(|subgraph| {
                self.calculate_subgraph(subgraph)
            })
            .collect();
        
        // Clear dirty set
        self.dirty_cells.clear();
        
        updates
    }
    
    fn calculate_subgraph(&self, cells: Vec<CellId>) -> Vec<CellUpdate> {
        let mut updates = Vec::new();
        
        for cell_id in cells {
            if let Some(cell) = self.cells.get(&cell_id) {
                if let Some(ref formula) = cell.formula {
                    let result = self.evaluate_formula(formula, cell_id);
                    updates.push(CellUpdate {
                        cell_id,
                        new_value: result,
                    });
                }
            }
        }
        
        updates
    }
    
    fn evaluate_formula(&self, formula: &str, context: CellId) -> CellValue {
        let ast = self.parse_formula(formula);
        self.evaluate_ast(&ast, context)
    }
    
    fn evaluate_ast(&self, node: &AstNode, context: CellId) -> CellValue {
        match node {
            AstNode::Number(n) => CellValue::Number(*n),
            AstNode::String(s) => CellValue::String(s.clone()),
            AstNode::Boolean(b) => CellValue::Boolean(*b),
            AstNode::CellRef(cell_ref) => {
                let resolved = self.resolve_cell_ref(cell_ref, context);
                self.cells.get(&resolved)
                    .map(|c| c.value.clone())
                    .unwrap_or(CellValue::Empty)
            }
            AstNode::BinaryOp { op, left, right } => {
                let left_val = self.evaluate_ast(left, context);
                let right_val = self.evaluate_ast(right, context);
                self.apply_binary_op(op, left_val, right_val)
            }
            AstNode::FunctionCall { name, args } => {
                let arg_values: Vec<CellValue> = args
                    .iter()
                    .map(|a| self.evaluate_ast(a, context))
                    .collect();
                self.call_function(name, arg_values)
            }
            // ... more node types
        }
    }
}
```

---

## Frontend Integration

### Tauri IPC Bridge

```typescript
// src/lib/tauri-bridge.ts

import { invoke } from "@tauri-apps/api/tauri";
import { listen } from "@tauri-apps/api/event";
import { open, save } from "@tauri-apps/api/dialog";
import { readTextFile, writeTextFile } from "@tauri-apps/api/fs";

export class TauriBridge {
  // Workbook operations
  async openWorkbook(path?: string): Promise<WorkbookInfo> {
    const filePath = path || await open({
      filters: [
        { name: "Excel", extensions: ["xlsx", "xlsm", "xls"] },
        { name: "CSV", extensions: ["csv"] },
        { name: "All Files", extensions: ["*"] }
      ]
    });
    
    if (!filePath) throw new Error("No file selected");
    
    return invoke<WorkbookInfo>("open_workbook", { path: filePath });
  }
  
  async saveWorkbook(path?: string): Promise<void> {
    let savePath = path;
    
    if (!savePath) {
      savePath = await save({
        filters: [
          { name: "Excel", extensions: ["xlsx"] }
        ]
      });
    }
    
    if (!savePath) throw new Error("No save path selected");
    
    return invoke("save_workbook", { path: savePath });
  }
  
  // Cell operations
  async getCell(sheetId: string, row: number, col: number): Promise<CellValue> {
    return invoke<CellValue>("get_cell", { sheetId, row, col });
  }
  
  async setCell(
    sheetId: string,
    row: number,
    col: number,
    value?: any,
    formula?: string
  ): Promise<CellUpdate[]> {
    return invoke<CellUpdate[]>("set_cell", {
      sheetId,
      row,
      col,
      value,
      formula
    });
  }
  
  async getRange(
    sheetId: string,
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number
  ): Promise<RangeData> {
    return invoke<RangeData>("get_range", {
      sheetId,
      startRow,
      startCol,
      endRow,
      endCol
    });
  }
  
  // Event listeners
  async onFileDrop(callback: (paths: string[]) => void): Promise<void> {
    await listen<string[]>("file-dropped", (event) => {
      callback(event.payload);
    });
  }
  
  async onCloseRequested(callback: () => void): Promise<void> {
    await listen("close-requested", () => {
      callback();
    });
  }
}

export const tauriBridge = new TauriBridge();
```

### Native Dialogs

```typescript
// src/lib/dialogs.ts

import { message, ask, confirm } from "@tauri-apps/api/dialog";

export async function showError(title: string, message: string): Promise<void> {
  await message(message, { title, type: "error" });
}

export async function showWarning(title: string, message: string): Promise<void> {
  await message(message, { title, type: "warning" });
}

export async function askSaveChanges(): Promise<"save" | "discard" | "cancel"> {
  const result = await ask(
    "You have unsaved changes. Do you want to save them?",
    {
      title: "Save Changes",
      type: "warning",
      okLabel: "Save",
      cancelLabel: "Discard"
    }
  );
  
  if (result === null) return "cancel";
  return result ? "save" : "discard";
}

export async function confirmDelete(itemName: string): Promise<boolean> {
  return confirm(
    `Are you sure you want to delete "${itemName}"?`,
    { title: "Confirm Delete", type: "warning" }
  );
}
```

---

## Native Features

### System Tray

```rust
// src-tauri/src/tray.rs

use tauri::{
    CustomMenuItem, SystemTray, SystemTrayMenu, SystemTrayMenuItem, SystemTrayEvent,
    Manager,
};

pub fn create_tray() -> SystemTray {
    let new = CustomMenuItem::new("new", "New Workbook");
    let open = CustomMenuItem::new("open", "Open...");
    let recent = CustomMenuItem::new("recent", "Recent Files");
    let quit = CustomMenuItem::new("quit", "Quit");
    
    let menu = SystemTrayMenu::new()
        .add_item(new)
        .add_item(open)
        .add_item(recent)
        .add_native_item(SystemTrayMenuItem::Separator)
        .add_item(quit);
    
    SystemTray::new().with_menu(menu)
}

pub fn handle_tray_event(app: &tauri::AppHandle, event: SystemTrayEvent) {
    match event {
        SystemTrayEvent::LeftClick { .. } => {
            // Show main window
            if let Some(window) = app.get_window("main") {
                window.show().unwrap();
                window.set_focus().unwrap();
            }
        }
        SystemTrayEvent::MenuItemClick { id, .. } => {
            match id.as_str() {
                "new" => {
                    app.emit_all("tray-new", ()).unwrap();
                }
                "open" => {
                    app.emit_all("tray-open", ()).unwrap();
                }
                "quit" => {
                    // Delegate quit-handling to the frontend so it can:
                    // - fire `Workbook_BeforeClose` macros
                    // - prompt for unsaved changes
                    // - decide whether to exit or keep running
                    app.emit_all("tray-quit", ()).unwrap();
                }
                _ => {}
            }
        }
        _ => {}
    }
}
```

### Global Shortcuts

```rust
// src-tauri/src/shortcuts.rs

use tauri::{AppHandle, GlobalShortcutManager, Manager};

pub async fn setup_shortcuts(app: AppHandle) {
    let mut manager = app.global_shortcut_manager();
    
    // Quick open
    manager
        .register("CmdOrCtrl+Shift+O", move || {
            app.emit_all("shortcut-quick-open", ()).unwrap();
        })
        .unwrap();
    
    // Command palette
    manager
        .register("CmdOrCtrl+Shift+P", move || {
            app.emit_all("shortcut-command-palette", ()).unwrap();
        })
        .unwrap();
}
```

### File Associations

#### macOS (Info.plist)

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleDocumentTypes</key>
    <array>
        <dict>
            <key>CFBundleTypeName</key>
            <string>Excel Spreadsheet</string>
            <key>CFBundleTypeRole</key>
            <string>Editor</string>
            <key>LSHandlerRank</key>
            <string>Alternate</string>
            <key>LSItemContentTypes</key>
            <array>
                <string>org.openxmlformats.spreadsheetml.sheet</string>
                <string>com.microsoft.excel.xls</string>
            </array>
        </dict>
        <dict>
            <key>CFBundleTypeName</key>
            <string>CSV File</string>
            <key>CFBundleTypeRole</key>
            <string>Editor</string>
            <key>LSItemContentTypes</key>
            <array>
                <string>public.comma-separated-values-text</string>
            </array>
        </dict>
    </array>
</dict>
</plist>
```

#### Windows (Installer)

```nsis
; File associations in NSIS installer
!macro RegisterFileType ext description icon
    WriteRegStr HKCR ".${ext}" "" "Formula.${ext}"
    WriteRegStr HKCR "Formula.${ext}" "" "${description}"
    WriteRegStr HKCR "Formula.${ext}\DefaultIcon" "" "$INSTDIR\${icon}"
    WriteRegStr HKCR "Formula.${ext}\shell\open\command" "" '"$INSTDIR\Formula.exe" "%1"'
!macroend

Section "File Associations"
    !insertmacro RegisterFileType "xlsx" "Excel Spreadsheet" "xlsx.ico"
    !insertmacro RegisterFileType "xls" "Excel 97-2003 Spreadsheet" "xls.ico"
    !insertmacro RegisterFileType "csv" "CSV File" "csv.ico"
SectionEnd
```

---

## Auto-Update

### Update Configuration

```rust
// src-tauri/src/updater.rs

use tauri::api::dialog::{MessageDialogBuilder, MessageDialogKind};
use tauri::updater::UpdateBuilder;

pub async fn check_for_updates(app: &tauri::AppHandle) -> Result<(), tauri::Error> {
    let update = app.updater().check().await?;
    
    if let Some(update) = update {
        // Show update dialog
        let should_update = MessageDialogBuilder::new(
            "Update Available",
            &format!(
                "Version {} is available. Would you like to update now?\n\n{}",
                update.version,
                update.body.as_deref().unwrap_or("")
            ),
        )
        .kind(MessageDialogKind::Info)
        .show();
        
        if should_update {
            update.download_and_install().await?;
        }
    }
    
    Ok(())
}
```

### tauri.conf.json Update Settings

```json
{
  "tauri": {
    "updater": {
      "active": true,
      "endpoints": [
        "https://releases.formula.app/{{target}}/{{arch}}/{{current_version}}"
      ],
      "dialog": true,
      "pubkey": "dW50cnVzdGVkIGNvbW1lbnQ6...",
      "windows": {
        "installMode": "passive"
      }
    }
  }
}
```

---

## Performance Optimization

### Startup Optimization

```rust
// src-tauri/src/main.rs

fn main() {
    // Pre-allocate calculation engine
    let engine = CalcEngine::with_capacity(100_000);
    
    tauri::Builder::default()
        .manage(Mutex::new(engine))
        // Lazy load non-critical features
        .setup(|app| {
            // Start background initialization
            tauri::async_runtime::spawn(async {
                // Load recent files list
                // Initialize AI models
                // Pre-warm caches
            });
            
            Ok(())
        })
        .run(tauri::generate_context!())
        .expect("error running application");
}
```

### Memory Management

```rust
// Efficient cell storage
pub struct CellStore {
    // Use sparse storage - only store non-empty cells
    cells: HashMap<CellId, Cell>,
    
    // Use arena allocator for formula strings
    formula_arena: bumpalo::Bump,
    
    // Cache frequently accessed ranges
    range_cache: LruCache<RangeKey, Vec<CellValue>>,
}

impl CellStore {
    pub fn get_range(&mut self, key: RangeKey) -> Vec<CellValue> {
        // Check cache first
        if let Some(cached) = self.range_cache.get(&key) {
            return cached.clone();
        }
        
        // Compute and cache
        let values = self.compute_range(&key);
        self.range_cache.put(key, values.clone());
        values
    }
}
```

---

## Cross-Platform Considerations

### WebView Differences

| Platform | WebView | Notes |
|----------|---------|-------|
| macOS | WebKit | Best performance, full feature support |
| Windows | WebView2 (Chromium) | Requires runtime install |
| Linux | WebKitGTK | May vary by distro |

### Platform-Specific Code

```rust
#[cfg(target_os = "macos")]
fn setup_macos(app: &tauri::App) {
    use cocoa::appkit::{NSApp, NSApplication};
    
    // Enable native macOS features
    unsafe {
        let ns_app = NSApp();
        ns_app.setActivationPolicy_(cocoa::appkit::NSApplicationActivationPolicyRegular);
    }
}

#[cfg(target_os = "windows")]
fn setup_windows(app: &tauri::App) {
    // Windows-specific setup
    // e.g., jumplist, taskbar integration
}

#[cfg(target_os = "linux")]
fn setup_linux(app: &tauri::App) {
    // Linux-specific setup
    // e.g., desktop file, icon theme
}
```

---

## Distribution

### Build Pipeline

```yaml
# .github/workflows/release.yml
name: Release

on:
  push:
    tags:
      - 'v*'

jobs:
  build:
    strategy:
      matrix:
        platform: [macos-latest, ubuntu-22.04, windows-latest]
    
    runs-on: ${{ matrix.platform }}
    
    steps:
      - uses: actions/checkout@v4
      
      - name: Setup Node
        uses: actions/setup-node@v4
        with:
          node-version: 20
      
      - name: Setup Rust
        uses: dtolnay/rust-toolchain@stable
      
      - name: Install dependencies (Linux)
        if: matrix.platform == 'ubuntu-22.04'
        run: |
          sudo apt-get update
          sudo apt-get install -y libgtk-3-dev libwebkit2gtk-4.0-dev
      
      - name: Install frontend dependencies
        run: npm ci
      
      - name: Build
        uses: tauri-apps/tauri-action@v0
        env:
          GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
          TAURI_PRIVATE_KEY: ${{ secrets.TAURI_PRIVATE_KEY }}
          TAURI_KEY_PASSWORD: ${{ secrets.TAURI_KEY_PASSWORD }}
        with:
          tagName: v__VERSION__
          releaseName: 'Formula v__VERSION__'
          releaseBody: 'See the assets below for download links.'
          releaseDraft: true
          prerelease: false
```

### Code Signing

```bash
# macOS
codesign --deep --force --options runtime \
  --sign "Developer ID Application: Your Company" \
  --entitlements entitlements.plist \
  Formula.app

# Notarize
xcrun notarytool submit Formula.dmg \
  --apple-id "$APPLE_ID" \
  --password "$NOTARIZE_PASSWORD" \
  --team-id "$TEAM_ID" \
  --wait

# Windows (signtool)
signtool sign /f certificate.pfx /p $PASSWORD \
  /t http://timestamp.digicert.com \
  Formula-Setup.exe
```
