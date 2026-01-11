# Extensibility & Plugin Architecture

## Overview

A vibrant extension ecosystem is critical for long-term success. We follow VS Code's model: extensions run in isolated processes with well-defined APIs, enabling powerful customization without compromising stability or security.

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
4. **Execution + contributions** via `ExtensionHost` (`packages/extension-host`)
   - Spawns an isolated worker per extension
   - Registers contributions (commands, panels, menus, keybindings, custom functions, data connectors)
5. **Hot reload on change**
   - Install/update calls `ExtensionHostManager.reloadExtension(id)`
   - Uninstall calls `ExtensionHostManager.unloadExtension(id)` (removes contributions and terminates worker)

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
    "onCustomFunction:MYFUNCTION"
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

---

## Extension API

### Core API

```typescript
// @formula/extension-api

// Source of truth: `packages/extension-api/index.d.ts`
//
// The current API is intentionally small and async-first: calls go through the host
// (worker_threads in Node, WebWorker in the desktop/webview runtime).

type CellValue = string | number | boolean | null;

interface Disposable {
  dispose(): void;
}

interface Workbook {
  readonly name: string;
  readonly path?: string | null;
}

interface Sheet {
  readonly id: string;
  readonly name: string;
}

interface Range {
  readonly startRow: number;
  readonly startCol: number;
  readonly endRow: number;
  readonly endCol: number;
  readonly values: CellValue[][];
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
  const connector = formula.dataConnectors.register("myExtension.salesforce", {
    id: "salesforce",
    name: "Salesforce",
    icon: "$(cloud)",
    
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
- **Configurable**: hosts may override defaults via `activationTimeoutMs`, `commandTimeoutMs`, and
  `customFunctionTimeoutMs` when constructing the extension host.
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

### Browser extension sandbox (best-effort)

In the browser runtime, extensions run inside a Web Worker. JavaScript does not provide the same
process-level or `vm` isolation primitives as Node, so the browser host applies **best-effort**
guardrails:

- Replace `fetch` and `WebSocket` with permission-gated wrappers and lock them down on `globalThis`
  (and the prototype chain when possible).
- Disable other obvious network primitives such as `XMLHttpRequest`.
- Disable nested script-loading/execution primitives (`importScripts`, `Worker`, `SharedWorker`) to
  avoid spawning a fresh worker with pristine globals.

Limitations:

- The ESM loader in browsers can still fetch module graphs via `import`/`import()` in ways that are
  not interceptable from inside a standard worker. Production deployments should pair the worker
  guardrails with CSP / extension packaging policies that prevent loading untrusted remote scripts.

### Permission System

```typescript
type Permission = 
  | "cells.read"
  | "cells.write"
  | "sheets.manage"
  | "workbook.manage"
  | "network"
  | "clipboard"
  | "storage"
  | "ui.panels"
  | "ui.commands"
  | "ui.menus";

const API_PERMISSIONS: Record<string, Permission[]> = {
  "cells.getCell": ["cells.read"],
  "cells.getRange": ["cells.read"],
  "cells.setCell": ["cells.write"],
  "cells.setRange": ["cells.write"],
  "sheets.createSheet": ["sheets.manage"],
  "sheets.deleteSheet": ["sheets.manage"],
  "ui.createPanel": ["ui.panels"],
  "commands.registerCommand": ["ui.commands"],
  // ... more mappings
};

class PermissionManager {
  private grantedPermissions: Map<string, Set<Permission>> = new Map();
  
  async requestPermissions(
    extensionId: string,
    permissions: Permission[]
  ): Promise<boolean> {
    // Check if already granted
    const existing = this.grantedPermissions.get(extensionId) || new Set();
    const needed = permissions.filter(p => !existing.has(p));
    
    if (needed.length === 0) return true;
    
    // Prompt user
    const granted = await this.promptUser(extensionId, needed);
    
    if (granted) {
      const all = new Set([...existing, ...needed]);
      this.grantedPermissions.set(extensionId, all);
    }
    
    return granted;
  }
  
  hasPermission(extensionId: string, apiCall: string): boolean {
    const required = API_PERMISSIONS[apiCall] || [];
    const granted = this.grantedPermissions.get(extensionId) || new Set();
    
    return required.every(p => granted.has(p));
  }
}
```

---

## Extension Marketplace

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
    const required = ["name", "version", "publisher", "main", "engines"];
    for (const field of required) {
      if (!manifest[field]) {
        throw new Error(`Missing required field: ${field}`);
      }
    }
    
    // Validate version format
    if (!semver.valid(manifest.version)) {
      throw new Error(`Invalid version: ${manifest.version}`);
    }
    
    // Validate permissions
    if (manifest.permissions) {
      for (const perm of manifest.permissions) {
        if (!VALID_PERMISSIONS.includes(perm)) {
          throw new Error(`Invalid permission: ${perm}`);
        }
      }
    }
  }
}
```

---

## Extension Examples

### Repo sample extension (runnable)

This repository includes a runnable reference extension at `extensions/sample-hello/` which is used
by integration tests and marketplace packaging tests. It demonstrates:

- Command registration + execution
- Panels + view activation + webview messaging
- Permission-gated APIs (network, clipboard, cells)
- Custom functions (`SAMPLEHELLO_DOUBLE`)

`extensions/sample-hello/src/extension.js` is the source of truth, and
`extensions/sample-hello/dist/extension.js` is the built entrypoint referenced by the extension
manifest for the Node extension host. The build also emits `dist/extension.mjs` for the browser
extension host. Run `node extensions/sample-hello/build.js` to regenerate the dist files; CI
enforces that the dist entrypoints stay in sync with `src/extension.js`.

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
  
  // Register data connector
  const connector = formula.dataConnectors.register("stockData.connector", {
    name: "Stock Market Data",
    
    browse: async (config) => {
      return [
        { id: "quotes", name: "Real-time Quotes", type: "table" },
        { id: "historical", name: "Historical Data", type: "table" },
        { id: "fundamentals", name: "Fundamentals", type: "table" }
      ];
    },
    
    query: async (config, query) => {
      const apiKey = formula.config.get<string>("stockData.apiKey");
      // Execute query...
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
