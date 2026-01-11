# Extensibility & Plugin Architecture

## Overview

A vibrant extension ecosystem is critical for long-term success. We follow VS Code's model: extensions run in isolated processes with well-defined APIs, enabling powerful customization without compromising stability or security.

## Desktop UX (current)

The desktop app wires `@formula/extension-host` contribution points into the UI so extensions are actually usable without devtools:

- **Extensions panel**: open via the status bar **Extensions** button. It lists installed extensions, their contributed commands, and contributed panels.
- **Executing commands**: clicking a command routes to `BrowserExtensionHost.executeCommand(commandId, ...args)`. Errors surface as a toast.
- **Panels / webviews**:
  - Contributed panels (`contributes.panels`) are registered in the panel registry so they can be persisted in layouts.
  - Panels created programmatically (`formula.ui.createPanel`) are opened automatically in the layout when created.
- **Notifications + prompts**:
  - `formula.ui.showMessage` shows a **toast** in the desktop UI.
  - `formula.ui.showQuickPick` / `formula.ui.showInputBox` are implemented via native `<dialog>` prompts.

## Webview sandbox model (desktop)

Extension panels are rendered as a sandboxed `<iframe>` using `srcdoc`:

- `sandbox="allow-scripts"` (no `allow-same-origin`)
- No top navigation / popups enabled
- Communication is **postMessage-only**:
  - Webview → extension: `window.parent.postMessage(message, "*")`
  - Extension → webview: `panel.webview.postMessage(message)` delivered to the iframe via `postMessage`

This is a **best-effort** browser sandbox. The primary security boundary remains the Node host sandbox (vm + module restrictions).

## `when` clauses + context keys (subset)

The desktop implements a small, VS Code-inspired subset of `when` syntax for menus/keybindings:

- Operators: `&&`, `||`, `!`, parentheses
- Identifiers (context keys): `cellHasValue`, `hasSelection`, `sheetName`, …
- Equality: `==` / `!=` against string/number/boolean literals

Unknown/invalid clauses fail closed (treated as `false`).

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  MAIN PROCESS                                                               │
│  ├── Core Application                                                       │
│  ├── Extension Host Manager                                                 │
│  └── IPC Router                                                             │
├──────────────────────────────────┬──────────────────────────────────────────┤
│  EXTENSION HOST 1                │  EXTENSION HOST 2                        │
│  (Isolated Process)              │  (Isolated Process)                      │
│  ├── Extension A                 │  ├── Extension C                         │
│  ├── Extension B                 │  └── Extension D                         │
│  └── API Proxy                   │                                          │
├──────────────────────────────────┴──────────────────────────────────────────┤
│  EXTENSION API                                                              │
│  ├── Cell/Range Operations                                                  │
│  ├── UI Extensions (Commands, Panels, Menus)                               │
│  ├── Event Subscriptions                                                    │
│  ├── Custom Functions                                                       │
│  └── Data Connectors                                                        │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Desktop runtime flow (current implementation)

The desktop app wires marketplace installs into the Node extension host runtime:

1. **Install/update/uninstall** via the marketplace UI (`apps/desktop/src/panels/marketplace/MarketplacePanel.js`)
2. **Filesystem + state** via `ExtensionManager`
   - Packages are extracted to `extensionsDir/<publisher>.<name>/`
   - The installed set is tracked in `ExtensionManager.statePath` (JSON)
3. **Runtime loading** via `ExtensionHostManager` (`apps/desktop/src/extensions/ExtensionHostManager.js`)
    - Reads `statePath` and calls `ExtensionHost.loadExtension(...)` for each installed extension
    - Exposes `executeCommand`, `invokeCustomFunction`, and `listContributions` for the app to route UI actions
    - Verifies extracted extension integrity using the file manifest stored in `statePath` both on startup
      and during subsequent syncs; corrupted/tampered installs are quarantined and require reinstall
    - Emits `onDidChangeContributions()` events after reload/unload/sync so the app can refresh command palettes,
      menus, and panels in response to extension changes
4. **Execution + contributions** via `ExtensionHost` (`packages/extension-host`)
    - Spawns an isolated worker per extension
    - Registers contributions (commands, panels, menus, keybindings, custom functions, data connectors)
5. **Hot reload + runtime syncing**
    - `ExtensionHostManager.syncInstalledExtensions()` reconciles the runtime with `statePath`:
      loads new installs, reloads updated versions, and unloads removed extensions.
    - The runtime can **bind to installer change events** (`ExtensionManager.onDidChange`) so installs outside
      the marketplace panel still hot-reload automatically.
    - For out-of-process changes (another process writing `statePath`), the runtime can also watch the state file
      on disk via `ExtensionHostManager.watchStateFile()`.
    - `ExtensionHostManager` serializes host operations so `syncInstalledExtensions()`, reload/unload, and command
      execution do not race each other.
    - `syncInstalledExtensions()` is serialized/coalesced so multiple rapid install/update events coalesce into a
      single sync run.
    - Uninstall removes contributions and terminates the worker (via `ExtensionHost.unloadExtension`).

---

## Extension Manifest

```json
{
  "name": "my-extension",
  "displayName": "My Extension",
  "version": "1.0.0",
  "description": "A sample extension",
  "publisher": "mycompany",
  "license": "MIT",
  
  "engines": {
    "formula": "^1.0.0"
  },
  
  "main": "./dist/extension.js",
  "browser": "./dist/extension.mjs",
  "module": "./dist/extension.mjs",
  
  "activationEvents": [
    "onStartupFinished",
    "onCommand:myExtension.run",
    "onView:myExtension.panel",
    "onCustomFunction:MYFUNCTION",
    "onDataConnector:myExtension.connector"
  ],
  
  "contributes": {
    "commands": [
      {
        "command": "myExtension.run",
        "title": "Run My Extension",
        "category": "My Extension",
        "icon": "$(play)"
      }
    ],
    
    "menus": {
      "cell/context": [
        {
          "command": "myExtension.processCell",
          "when": "cellHasValue"
        }
      ],
      "toolbar": [
        {
          "command": "myExtension.run",
          "group": "extensions"
        }
      ]
    },
    
    "keybindings": [
      {
        "command": "myExtension.run",
        "key": "ctrl+shift+m",
        "mac": "cmd+shift+m"
      }
    ],
    
    "panels": [
      {
        "id": "myExtension.panel",
        "title": "My Panel",
        "icon": "$(graph)"
      }
    ],
    
    "customFunctions": [
      {
        "name": "MYFUNCTION",
        "description": "My custom function",
        "parameters": [
          {
            "name": "value",
            "type": "number",
            "description": "Input value"
          }
        ],
        "result": {
          "type": "number"
        }
      }
    ],
    
    "dataConnectors": [
      {
        "id": "myExtension.connector",
        "name": "My Data Source",
        "icon": "$(database)"
      }
    ],
    
    "configuration": {
      "title": "My Extension",
      "properties": {
        "myExtension.apiKey": {
          "type": "string",
          "description": "API key for external service"
        },
        "myExtension.enabled": {
          "type": "boolean",
          "default": true,
          "description": "Enable extension features"
        }
      }
    }
  },
  
  "permissions": [
    "network",
    "clipboard",
    "storage"
  ],
  
  "repository": {
    "type": "git",
    "url": "https://github.com/mycompany/my-extension"
  }
}
```

### Manifest validation + entrypoints

Formula validates extension manifests using **one shared validator** across:

- the **Node/desktop extension host**
- the **browser extension host**
- **marketplace publish-time** checks
- the **extension publisher** (so local packaging fails fast)

This means the marketplace rejects extension packages whose manifests would be rejected at runtime
(e.g. invalid permissions, malformed `contributes` blocks, or `activationEvents` that reference
unknown commands/panels/custom functions/data connectors).

Entrypoint fields:

- `main` (**required**) — CommonJS entrypoint used by the Node/desktop extension host. The file must
  exist in the published package. (Must end in `.js` or `.cjs`.)
- `browser` (optional) — browser-first entrypoint (ESM). If present, the file must exist in the
  published package. (Must end in `.js` or `.mjs`.)
- `module` (optional) — module entrypoint (ESM). If present, the file must exist in the published
  package. (Must end in `.js` or `.mjs`.)

The browser extension host loads `browser` → `module` → `main` (first defined wins). The Node
extension host always uses `main`.

### `engines.formula` semver range syntax

The extension host validates `engines.formula` using a small semver range implementation that is
shared across the marketplace and both the Node and browser extension hosts.

Supported range forms:

- `*` (any version)
- Exact versions: `1.2.3`
- Caret / tilde: `^1.2.3`, `~1.2.3`
- Comparators: `>=1.0.0`, `>1.0.0`, `<=2.0.0`, `<2.0.0`
- AND (whitespace-separated): `>=1.0.0 <2.0.0`
- OR (optional): `<1.0.0 || >=2.0.0`

Notes:

- Pre-release ordering follows semver precedence rules (e.g. `1.0.0-alpha < 1.0.0`).
- Build metadata (`+build.123`) is ignored for ordering.

---

## Extension API

### Core API

```typescript
// @formula/extension-api

// Source of truth: `packages/extension-api/index.d.ts`
//
// The current API is intentionally small and async-first: calls go through the host
// (worker_threads in Node, WebWorker in the desktop/webview runtime).
//
// `@formula/extension-api` ships both a CommonJS and an ESM entrypoint:
// - Node / CJS: `const formula = require("@formula/extension-api")`
// - Browser / ESM: `import * as formula from "@formula/extension-api"`
//
// They are **behaviorally identical** and both match the `index.d.ts` contract
// (including `Workbook.save/saveAs/close`, `Sheet.getRange/setRange/activate/rename`,
// and `Range.address/formulas`).

type CellValue = string | number | boolean | null;

interface Disposable {
  dispose(): void;
}

interface Workbook {
  readonly name: string;
  readonly path?: string | null;
  readonly sheets: Sheet[];
  readonly activeSheet: Sheet;
  save(): Promise<void>;
  saveAs(path: string): Promise<void>;
  close(): Promise<void>;
}

interface Sheet {
  readonly id: string;
  readonly name: string;
  getRange(ref: string): Promise<Range>;
  setRange(ref: string, values: CellValue[][]): Promise<void>;
  activate(): Promise<Sheet>;
  rename(name: string): Promise<Sheet>;
}

interface Range {
  readonly startRow: number;
  readonly startCol: number;
  readonly endRow: number;
  readonly endCol: number;
  readonly address: string;
  readonly values: CellValue[][];
  readonly formulas: (string | null)[][];
}

interface PanelWebview {
  html: string;
  setHtml(html: string): Promise<void>;
  postMessage(message: any): Promise<void>;
  onDidReceiveMessage(handler: (message: any) => void): Disposable;
}

interface Panel extends Disposable {
  readonly id: string;
  readonly webview: PanelWebview;
}

declare namespace formula {
  export namespace workbook {
    function getActiveWorkbook(): Promise<Workbook>;
    function openWorkbook(path: string): Promise<Workbook>;
    function createWorkbook(): Promise<Workbook>;
    function save(): Promise<void>;
    function saveAs(path: string): Promise<void>;
    function close(): Promise<void>;
  }

  export namespace sheets {
    function getActiveSheet(): Promise<Sheet>;
    function getSheet(name: string): Promise<Sheet | undefined>;
    function activateSheet(name: string): Promise<Sheet>;
    function createSheet(name: string): Promise<Sheet>;
    function renameSheet(oldName: string, newName: string): Promise<void>;
    function deleteSheet(name: string): Promise<void>;
  }

  export namespace cells {
    function getSelection(): Promise<Range>;
    function getRange(ref: string): Promise<Range>;
    function getCell(row: number, col: number): Promise<CellValue>;
    function setCell(row: number, col: number, value: CellValue): Promise<void>;
    function setRange(ref: string, values: CellValue[][]): Promise<void>;
  }

  export namespace commands {
    function registerCommand(
      id: string,
      handler: (...args: any[]) => any | Promise<any>
    ): Promise<Disposable>;
    function executeCommand(id: string, ...args: any[]): Promise<any>;
  }

  export namespace functions {
    function register(
      name: string,
      def: {
        description?: string;
        parameters?: Array<{ name: string; type: string; description?: string }>;
        result?: { type: string };
        isAsync?: boolean;
        returnsArray?: boolean;
        handler: (...args: any[]) => any | Promise<any>;
      }
    ): Promise<Disposable>;
  }

  export interface DataConnectorQueryResult {
    columns: string[];
    rows: any[][];
  }

  export interface DataConnectorImplementation {
    browse(config: any, path?: string | null): Promise<any>;
    query(config: any, query: any): Promise<DataConnectorQueryResult>;
    getConnectionConfig?: (...args: any[]) => Promise<any>;
    testConnection?: (...args: any[]) => Promise<any>;
    getQueryBuilder?: (...args: any[]) => Promise<any>;
  }

  export namespace dataConnectors {
    function register(connectorId: string, impl: DataConnectorImplementation): Promise<Disposable>;
  }

  export namespace network {
    function fetch(url: string, init?: any): Promise<{
      readonly ok: boolean;
      readonly status: number;
      readonly statusText: string;
      readonly url: string;
      readonly headers: { get(name: string): string | undefined };
      text(): Promise<string>;
      json<T = any>(): Promise<T>;
    }>;
  }

  export namespace clipboard {
    function readText(): Promise<string>;
    function writeText(text: string): Promise<void>;
  }

  export namespace ui {
    type MessageType = "info" | "warning" | "error";

    function showMessage(message: string, type?: MessageType): Promise<void>;
    function showInputBox(options: { prompt?: string; value?: string; placeHolder?: string }): Promise<
      string | undefined
    >;
    function showQuickPick<T>(
      items: Array<{ label: string; value: T; description?: string; detail?: string }>,
      options?: { placeHolder?: string }
    ): Promise<T | undefined>;
    function createPanel(
      id: string,
      options: { title: string; icon?: string; position?: "left" | "right" | "bottom" }
    ): Promise<Panel>;
    function registerContextMenu(
      menuId: string,
      items: Array<{ command: string; when?: string; group?: string }>
    ): Promise<Disposable>;
  }

  export const storage: {
    get<T = unknown>(key: string): Promise<T | undefined>;
    set<T = unknown>(key: string, value: T): Promise<void>;
    delete(key: string): Promise<void>;
  };

  export namespace config {
    function get<T = unknown>(key: string): Promise<T | undefined>;
    function update(key: string, value: any): Promise<void>;
    function onDidChange(callback: (e: { key: string; value: any }) => void): Disposable;
  }

  export namespace events {
    function onSelectionChanged(callback: (e: { selection: Range }) => void): Disposable;
    function onCellChanged(callback: (e: { row: number; col: number; value: CellValue }) => void): Disposable;
    function onSheetActivated(callback: (e: { sheet: Sheet }) => void): Disposable;
    function onWorkbookOpened(callback: (e: { workbook: Workbook }) => void): Disposable;
    function onBeforeSave(callback: (e: { workbook: Workbook }) => void): Disposable;
    function onViewActivated(callback: (e: { viewId: string }) => void): Disposable;
  }

  export namespace context {
    const extensionId: string;
    const extensionPath: string;
    const extensionUri: string;
    const globalStoragePath: string;
    const workspaceStoragePath: string;
  }
}
```

### Custom Functions API

```typescript
// Register a custom function
export function activate(context: ExtensionContext) {
  // Simple function
  const myFunc = formula.functions.register("MYFUNCTION", {
    description: "Doubles the input value",
    parameters: [
      { name: "value", type: "number", description: "Value to double" }
    ],
    result: { type: "number" },
    
    handler: (value: number) => {
      return value * 2;
    }
  });
  
  // Async function (for external data)
  const fetchFunc = formula.functions.register("FETCHDATA", {
    description: "Fetches data from API",
    parameters: [
      { name: "endpoint", type: "string" },
      { name: "field", type: "string" }
    ],
    result: { type: "any" },
    isAsync: true,
    
    handler: async (endpoint: string, field: string) => {
      const response = await fetch(endpoint);
      const data = await response.json();
      return data[field];
    }
  });
  
  // Array-returning function
  const splitFunc = formula.functions.register("SPLITALL", {
    description: "Splits text into array",
    parameters: [
      { name: "text", type: "string" },
      { name: "delimiter", type: "string" }
    ],
    result: { type: "array" },
    returnsArray: true,
    
    handler: (text: string, delimiter: string) => {
      return text.split(delimiter).map(s => [s]);
    }
  });
  
  context.subscriptions.push(myFunc, fetchFunc, splitFunc);
}
```

### Panel API (Webviews)

```typescript
export function activate(context: ExtensionContext) {
  // Create panel
  const panel = formula.ui.createPanel("myExtension.panel", {
    title: "My Panel",
    icon: "$(graph)",
    position: "right"
  });
  
  // Set HTML content
  panel.webview.html = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: var(--vscode-font-family); }
        button { background: var(--vscode-button-background); color: var(--vscode-button-foreground); }
      </style>
    </head>
    <body>
      <h1>My Panel</h1>
      <button id="analyze">Analyze Selection</button>
      <div id="results"></div>
      
      <script>
        const vscode = acquireVsCodeApi();
        
        document.getElementById('analyze').addEventListener('click', () => {
          vscode.postMessage({ command: 'analyze' });
        });
        
        window.addEventListener('message', event => {
          const message = event.data;
          if (message.command === 'results') {
            document.getElementById('results').innerHTML = message.html;
          }
        });
      </script>
    </body>
    </html>
  `;
  
  // Handle messages from webview
  panel.webview.onDidReceiveMessage(async (message) => {
    if (message.command === 'analyze') {
      const selection = formula.cells.getSelection();
      const values = selection.values;
      
      // Process values
      const sum = values.flat().reduce((a, b) => a + (typeof b === 'number' ? b : 0), 0);
      
      // Send results back
      panel.webview.postMessage({
        command: 'results',
        html: `<p>Sum: ${sum}</p>`
      });
    }
  });
  
  context.subscriptions.push(panel);
}
```

### Data Connector API

```typescript
export function activate(context: ExtensionContext) {
  // The connector's metadata (name/icon) is declared in the extension manifest via
  // `contributes.dataConnectors`. Runtime registration only supplies implementation.
  const connector = await formula.dataConnectors.register("myExtension.salesforce", {
    // Connection configuration
    getConnectionConfig: async () => {
      return [
        { name: "instanceUrl", label: "Instance URL", type: "string", required: true },
        { name: "username", label: "Username", type: "string", required: true },
        { name: "password", label: "Password", type: "password", required: true },
        { name: "securityToken", label: "Security Token", type: "password", required: true }
      ];
    },
    
    // Test connection
    testConnection: async (config) => {
      try {
        const sf = new Salesforce(config);
        await sf.login();
        return { success: true };
      } catch (error) {
        return { success: false, error: error.message };
      }
    },
    
    // Browse available data
    browse: async (config, path) => {
      const sf = new Salesforce(config);
      
      if (!path) {
        // Root level - show objects
        const objects = await sf.describeGlobal();
        return objects.map(obj => ({
          id: obj.name,
          name: obj.label,
          type: "table",
          children: true
        }));
      } else {
        // Object level - show fields
        const describe = await sf.describe(path);
        return describe.fields.map(field => ({
          id: field.name,
          name: field.label,
          type: "column",
          dataType: field.type
        }));
      }
    },
    
    // Execute query
    query: async (config, query) => {
      const sf = new Salesforce(config);
      const result = await sf.query(query.soql);
      
      return {
        columns: query.fields,
        rows: result.records.map(r => query.fields.map(f => r[f]))
      };
    },
    
    // Build query UI
    getQueryBuilder: () => ({
      type: "soql",
      placeholder: "SELECT Id, Name FROM Account LIMIT 10"
    })
  });
  
  context.subscriptions.push(connector);
}
```

---

## Extension Host

### Process Isolation

```typescript
class ExtensionHostManager {
  private hosts: Map<string, ExtensionHost> = new Map();
  
  async startHost(extensions: Extension[]): Promise<ExtensionHost> {
    const hostId = crypto.randomUUID();
    
    // Spawn worker process
    const worker = new Worker("extension-host-worker.js", {
      name: `ExtensionHost-${hostId}`
    });
    
    // Set up IPC
    const host = new ExtensionHost(hostId, worker);
    
    // Initialize extensions
    for (const ext of extensions) {
      await host.loadExtension(ext);
    }
    
    this.hosts.set(hostId, host);
    return host;
  }
  
  async terminateHost(hostId: string): Promise<void> {
    const host = this.hosts.get(hostId);
    if (host) {
      await host.dispose();
      this.hosts.delete(hostId);
    }
  }
}

class ExtensionHost {
  constructor(
    private id: string,
    private worker: Worker
  ) {
    this.setupIPC();
  }
  
  private setupIPC(): void {
    this.worker.onmessage = (event) => {
      this.handleMessage(event.data);
    };
    
    this.worker.onerror = (error) => {
      console.error(`Extension host error:`, error);
      this.handleCrash();
    };
  }
  
  private handleMessage(message: HostMessage): void {
    switch (message.type) {
      case "api_call":
        this.handleAPICall(message);
        break;
      case "event":
        this.handleEvent(message);
        break;
      case "log":
        console.log(`[Extension ${message.extensionId}]`, message.data);
        break;
    }
  }
  
  private async handleAPICall(message: APICallMessage): Promise<void> {
    const { id, namespace, method, args } = message;
    
    try {
      // Validate permission
      if (!this.hasPermission(message.extensionId, namespace, method)) {
        throw new Error(`Permission denied: ${namespace}.${method}`);
      }
      
      // Execute API call
      const result = await this.executeAPICall(namespace, method, args);
      
      // Send result back
      this.worker.postMessage({
        type: "api_result",
        id,
        result
      });
    } catch (error) {
      this.worker.postMessage({
        type: "api_error",
        id,
        error: error.message
      });
    }
  }
}
```

### Execution Guardrails

Extensions are treated as untrusted code. The host applies **best-effort** guardrails so a single
misbehaving extension cannot hang or OOM the main app:

- **Activation timeout**: extension activation is bounded (default: 5s). If activation exceeds the
  limit, the worker is terminated and the activation promise rejects with a timeout error.
- **Command/custom function timeouts**: command handlers and custom function invocations are also
  bounded (default: 5s). Timeouts terminate the worker and reject any other in-flight requests for
  that worker to avoid leaks.
- **Configurable**: hosts may override defaults via `activationTimeoutMs`, `commandTimeoutMs`,
  `customFunctionTimeoutMs`, and `dataConnectorTimeoutMs` when constructing the extension host.
- **Memory caps (best-effort)**: extension workers are started with `worker_threads` `resourceLimits`
  based on a per-host `memoryMb` setting (default: 256MB, Node host only). This caps the V8 heap, but
  does not cover all native/external allocations.
- **Crash/restart**: on crash/timeout, the extension is marked inactive and the next activation will
  spawn a fresh worker. Hosts can also explicitly recycle an extension via `reloadExtension(id)`.

### Node extension sandbox (security boundary)

For the desktop/Node runtime, extensions execute inside a hardened `vm` context with a minimal
CommonJS loader. This is the primary security boundary that makes the permission system enforceable.

Key properties:

- **No Node builtins**: `require("node:fs")`, `require("fs")`, `import("node:fs")`, etc. are blocked.
- **No ESM dynamic import**: `import(...)` is rejected (prevents bypassing CommonJS loaders).
- **No `process` escape hatches**: extensions do not receive the real `process` object;
  `process.binding(...)` is blocked.
- **No string codegen**: `eval` / `new Function(...)` are disabled via
  `codeGeneration: { strings: false, wasm: false }`.
- **Symlink-safe resolution**: all filesystem loads are validated using `realpath` so extensions
  cannot escape the extension root via symlinks (including when Node is started with
  `--preserve-symlinks`).
- **Filesystem/network access is API-only**: extensions must use Formula APIs (e.g.
  `formula.network.fetch`, `formula.storage.*`) which are permission-gated by the host.
- **Module restrictions**:
  - allowed: `require("@formula/extension-api")` (or `require("formula")`)
  - allowed: relative `./` / `../` requires that resolve inside the extension folder
  - blocked: any other specifier (including `node_modules` dependencies)

Implication: extensions should be shipped as a bundled CommonJS entrypoint (and can include relative
chunks inside the extension folder if needed).

#### Recommended build pipeline (repo supported)

This repo ships an esbuild-based bundler at `tools/extension-builder/` that produces **two**
entrypoints for each extension:

- **Node (CommonJS)**: `manifest.main` (defaults to `./dist/extension.js`)
- **Browser (ESM)**: `manifest.browser` or `manifest.module` (defaults to `./dist/extension.mjs`)

The builder bundles your extension source into those outputs and validates that the resulting code
conforms to sandbox restrictions:

- no Node builtins (`fs`, `node:fs`, etc)
- no remote/URL imports (`https://...`)
- no `import(...)` (dynamic import)
- (strict mode) no `eval` / `new Function`

In this repo you can run it via the root `pnpm` scripts:

```bash
pnpm extension:build <extensionDir>
pnpm extension:check <extensionDir>
```

When published, it is also available as a binary: `formula-extension-builder`.

#### Packaging + publishing

Once built, extensions are distributed as a signed `.fextpkg` archive:

```bash
# Build (writes dist entrypoints)
pnpm extension:build extensions/my-extension

# Pack into a signed archive (generates an ephemeral keypair if you don't pass one)
pnpm extension:pack extensions/my-extension --out ./my-extension.fextpkg --private-key ./publisher-private.pem

# Verify / inspect the resulting package
pnpm extension:verify ./my-extension.fextpkg --pubkey ./publisher-public.pem
pnpm extension:inspect ./my-extension.fextpkg
```

To publish to a marketplace instance, use the repo tool:

```bash
node tools/extension-publisher/src/cli.js publish extensions/my-extension \
  --marketplace https://marketplace.example.com \
  --token <publisher-token> \
  --private-key ./publisher-private.pem
```

### Browser extension sandbox (best-effort)

In the browser runtime, extensions run inside a Web Worker. JavaScript does not provide the same
process-level or `vm` isolation primitives as Node, so the browser host applies **best-effort**
guardrails:

- Replace `fetch` and `WebSocket` with permission-gated wrappers and lock them down on `globalThis`
  (and the prototype chain when possible).
- Disable other obvious network primitives such as `XMLHttpRequest`, `EventSource`, WebTransport, and
  WebRTC (`RTCPeerConnection`).
- Disable nested script-loading/execution primitives (`importScripts`, `Worker`, `SharedWorker`) to
  avoid spawning a fresh worker with pristine globals.
- **Strict import policy (best-effort)**: before activating an extension, the worker fetches the
  entrypoint module and its static dependency graph and rejects:
  - any static import specifier that is not relative (`./` / `../`) or `@formula/extension-api`
    (optionally `formula` if an import map/alias is provided)
  - any dynamic `import(...)` usage
  - any module URL that resolves outside the extension base URL (including redirects)
  - implication: browser extensions must bundle third-party dependencies and reference any split
    chunks via relative imports inside the extension package.
- **Code generation lockdown (best-effort)**: `eval`, `Function` (and related constructors), and
  string timer callbacks (`setTimeout("...")`) are disabled by default.
  - configurable via `new BrowserExtensionHost({ sandbox: { strictImports, disableEval } })`

Limitations:

- These guardrails are not a complete security boundary. Browsers do not expose a hardened `vm`
  equivalent, and some escape hatches (and engine bugs) may still exist.
- Production deployments should pair the worker guardrails with a restrictive CSP and extension
  packaging policies that prevent loading untrusted remote scripts.

### Permission System

Permissions are **declared** in an extension manifest and **granted** at runtime by the host. Grants are
persisted per-extension and can be inspected/revoked by the application.

#### Declaring permissions

Legacy (string) declarations are still supported:

```json
{
  "permissions": ["ui.commands", "network", "clipboard"]
}
```

The manifest validator also accepts a future object form (currently treated as declaring the same
top-level permission key):

```json
{
  "permissions": [
    "ui.commands",
    { "network": { "mode": "allowlist", "hosts": ["api.example.com"] } }
  ]
}
```

#### Stored grant format (v2)

On disk (Node) or in `localStorage` (browser), the host stores grants per extension as:

```json
{
  "publisher.name": {
    "cells.read": true,
    "ui.commands": true,
    "network": { "mode": "allowlist", "hosts": ["api.example.com", "*.example.org"] }
  }
}
```

Network modes:

- `full`: allow all outbound network access
- `allowlist`: allow only `hosts` that match the request (exact host, `*.wildcard`, or full origin like
  `https://api.example.com`)
- `deny`: deny all outbound network access

Both `formula.network.fetch(url)` and permission-gated `WebSocket(url)` connections are checked against
the effective network policy.

#### Host introspection + revocation APIs

Both the Node and browser hosts expose permission management helpers:

```ts
await host.getGrantedPermissions("publisher.name");
await host.revokePermissions("publisher.name", ["network"]); // omit list to revoke all
await host.resetAllPermissions(); // clears all extensions
```

#### Backwards compatibility / migration

Existing persisted permission data stored as string arrays is automatically migrated on load. Legacy
`"network"` grants are upgraded to `{ "mode": "full" }` to preserve behavior for already-trusted
extensions.

---  

## Extension Marketplace

Marketplace implementation details (HTTP endpoints, publish/download headers, caching) are documented in
[`docs/marketplace.md`](./marketplace.md).
### Web runtime install + update flow

The browser/web runtime uses a different installation model than the desktop/Node runtime:

- Extensions are downloaded as signed v2 `.fextpkg` blobs from the Marketplace.
- The client verifies the package **in the browser** (tar parsing + SHA-256 checksums + Ed25519
  signature verification) before persisting anything.
- Verified package bytes + metadata are stored in IndexedDB, and the extension is loaded into
  `BrowserExtensionHost` via a `blob:` module URL (no remote module graph imports).
- Updates replace the stored `{id, version}` atomically and reload the extension in the host if it is
  currently loaded.

Because `blob:` module URLs cannot resolve relative imports, `manifest.browser` should point at a
**single-file ESM bundle**.

### Extension Discovery

```typescript
interface MarketplaceExtension {
  id: string;
  name: string;
  displayName: string;
  publisher: string;
  version: string;
  description: string;
  categories: string[];
  tags: string[];
  rating: number;
  reviewCount: number;
  downloadCount: number;
  icon: string;
  screenshots: string[];
  readme: string;
  changelog: string;
  lastUpdated: Date;
  
  // Verification
  verified: boolean;
  featured: boolean;
}

class ExtensionMarketplace {
  private apiEndpoint = "https://marketplace.formula.app/api";
  
  async search(query: string, options?: SearchOptions): Promise<SearchResult> {
    const params = new URLSearchParams({
      q: query,
      category: options?.category || "",
      sortBy: options?.sortBy || "relevance",
      page: String(options?.page || 1),
      pageSize: String(options?.pageSize || 20)
    });
    
    const response = await fetch(`${this.apiEndpoint}/search?${params}`);
    return response.json();
  }
  
  async getExtension(id: string): Promise<MarketplaceExtension> {
    const response = await fetch(`${this.apiEndpoint}/extensions/${id}`);
    return response.json();
  }
  
  async install(id: string, version?: string): Promise<void> {
    // Get extension package
    const ext = await this.getExtension(id);
    const packageUrl = await this.getPackageUrl(id, version || ext.version);
    
    // Download
    const response = await fetch(packageUrl);
    const buffer = await response.arrayBuffer();
    
    // Verify signature
    const verified = await this.verifySignature(buffer, ext.publisher);
    if (!verified) {
      throw new Error("Extension signature verification failed");
    }
    
    // Extract and install
    await this.extractAndInstall(buffer);
    
    // Activate
    await this.activateExtension(id);
  }
  
  async uninstall(id: string): Promise<void> {
    // Deactivate
    await this.deactivateExtension(id);
    
    // Remove files
    await this.removeExtensionFiles(id);
    
    // Clean up storage
    await this.cleanupStorage(id);
  }
}
```

### Extension Publishing

```typescript
// CLI tool for publishing extensions
class ExtensionPublisher {
  async publish(extensionPath: string, options: PublishOptions): Promise<void> {
    // Validate package.json
    const manifest = await this.loadManifest(extensionPath);
    this.validateManifest(manifest);
    
    // Build extension
    await this.build(extensionPath);
    
    // Package
    const vsix = await this.package(extensionPath);
    
    // Sign
    const signed = await this.sign(vsix, options.privateKey);
    
    // Upload
    const response = await fetch(`${this.apiEndpoint}/publish`, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${options.accessToken}`,
        "Content-Type": "application/octet-stream"
      },
      body: signed
    });
    
    if (!response.ok) {
      throw new Error(`Publish failed: ${await response.text()}`);
    }
    
    console.log(`Successfully published ${manifest.name}@${manifest.version}`);
  }
  
  private validateManifest(manifest: ExtensionManifest): void {
    // Publishing uses the same core manifest validation rules as the extension hosts.
    //
    // Note: clients enforce `engines.formula` compatibility at runtime; marketplace publish-time
    // validation checks the manifest structure but does not gate on a specific engine version.
    validateExtensionManifest(manifest, { enforceEngine: false });
  }
}
```

---

## Installed extension integrity (tamper detection + quarantine)

Extensions are treated as **immutable** after installation.

### What we guarantee

1. **Package signature verification at install time**
   - When installing from the marketplace, the desktop app verifies the extension package signature
     using the publisher public key.
   - The installer persists verification metadata (package SHA-256, signature, and per-file
     SHA-256 + size).

2. **On-disk integrity verification before execution**
   - Before loading an installed extension into the runtime, the desktop app verifies the installed
     extension directory matches the recorded per-file hashes.
   - If any expected file is missing, modified, or if unexpected files are present, the extension is
     considered **corrupted**.

### Quarantine behavior (“corrupted”)

If integrity verification fails, the extension is **quarantined**:

- The extension is marked `corrupted: true` in the desktop marketplace state along with a timestamp
  and human-readable reason.
- The extension host refuses to load/execute the extension and surfaces an **“integrity check
  failed”** error.

This protects users from accidental edits to the extension folder and from malware that tampers with
installed extension code.

### Recovery (“repair”)

Users recover by reinstalling the extension from the marketplace:

- **Repair** downloads and verifies the package again, then re-extracts it atomically over the
  existing installation.
- On success, the `corrupted` flag is cleared and the extension can be loaded again.

--- 

## Extension Examples

### Repo sample extension (runnable)

This repository includes a runnable reference extension at `extensions/sample-hello/` which is used
by integration tests and marketplace packaging tests. It demonstrates:

- Command registration + execution
- Panels + view activation + webview messaging
- Permission-gated APIs (network, clipboard, cells)
- Custom functions (`SAMPLEHELLO_DOUBLE`)
- Data connectors (`sampleHello.connector`)

`extensions/sample-hello/src/extension.js` is the source of truth, and
`extensions/sample-hello/dist/extension.js` is the built entrypoint referenced by the extension
manifest for the Node extension host. The build also emits `dist/extension.mjs` for the browser
extension host. To regenerate:

```bash
pnpm extension:build extensions/sample-hello
```

### 1. Custom Visualization Extension

```typescript
// extension.ts
import * as formula from "@formula/extension-api";

export function activate(context: formula.ExtensionContext) {
  // Register command
  const cmd = formula.commands.registerCommand("chartViz.createChart", async () => {
    const selection = formula.cells.getSelection();
    const values = selection.values;
    
    // Create visualization panel
    const panel = formula.ui.createPanel("chartViz.panel", {
      title: "Chart Visualization",
      position: "right"
    });
    
    // Generate chart
    const chartConfig = inferChartType(values);
    
    panel.webview.html = `
      <!DOCTYPE html>
      <html>
      <head>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
      </head>
      <body>
        <canvas id="chart"></canvas>
        <script>
          const ctx = document.getElementById('chart').getContext('2d');
          new Chart(ctx, ${JSON.stringify(chartConfig)});
        </script>
      </body>
      </html>
    `;
  });
  
  context.subscriptions.push(cmd);
}

function inferChartType(data: CellValue[][]): ChartConfiguration {
  // Analyze data to determine best chart type
  const hasLabels = typeof data[0][0] === "string";
  const numericColumns = data[0].filter(v => typeof v === "number").length;
  
  if (numericColumns === 1) {
    return { type: "bar", data: formatBarData(data) };
  } else if (numericColumns === 2) {
    return { type: "scatter", data: formatScatterData(data) };
  } else {
    return { type: "line", data: formatLineData(data) };
  }
}
```

### 2. External Data Connector Extension

```typescript
// extension.ts
import * as formula from "@formula/extension-api";

export function activate(context: formula.ExtensionContext) {
  // Register custom function
  const stockFunc = formula.functions.register("STOCK", {
    description: "Get stock price",
    parameters: [
      { name: "symbol", type: "string", description: "Stock symbol" }
    ],
    result: { type: "number" },
    isAsync: true,
    
    handler: async (symbol: string) => {
      const apiKey = formula.config.get<string>("stockData.apiKey");
      const response = await fetch(
        `https://api.stockdata.com/v1/quote?symbol=${symbol}&api_key=${apiKey}`
      );
      const data = await response.json();
      return data.price;
    }
  });
  
  // Register data connector (metadata is declared in the manifest)
  const connector = await formula.dataConnectors.register("stockData.connector", {
    browse: async (config) => {
      return [
        { id: "quotes", name: "Real-time Quotes", type: "table" },
        { id: "historical", name: "Historical Data", type: "table" },
        { id: "fundamentals", name: "Fundamentals", type: "table" }
      ];
    },
    
    query: async (config, query) => {
      const apiKey = await formula.config.get<string>("stockData.apiKey");
      // Execute query...
      return { columns: [], rows: [] };
    }
  });
  
  context.subscriptions.push(stockFunc, connector);
}
```

---

## Testing Extensions

```js
// Example: run an extension under the Node extension host harness.
//
// In this repo we use the built-in `node:test` runner and spin up an
// `ExtensionHost` instance to load and exercise an extension end-to-end.
//
// Run with: `node --test`
const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");

const { ExtensionHost } = require("@formula/extension-host");

test("myExtension.run updates the workbook", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-test-"));

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(path.resolve("path/to/extension"));

  host.spreadsheet.setCell(0, 0, 1);
  await host.executeCommand("myExtension.run");

  assert.equal(host.spreadsheet.getCell(0, 0), 2);
});
```
