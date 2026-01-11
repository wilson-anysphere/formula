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

declare namespace formula {
  // Workbook operations
  export namespace workbook {
    export function getActiveWorkbook(): Workbook;
    export function openWorkbook(path: string): Promise<Workbook>;
    export function createWorkbook(): Workbook;
  }
  
  // Sheet operations
  export namespace sheets {
    export function getActiveSheet(): Sheet;
    export function getSheet(name: string): Sheet | undefined;
    export function createSheet(name: string): Sheet;
    export function deleteSheet(name: string): void;
  }
  
  // Cell operations
  export namespace cells {
    export function getCell(row: number, col: number): Cell;
    export function getRange(ref: string): Range;
    export function getSelection(): Range;
    
    export function setCell(row: number, col: number, value: CellValue): void;
    export function setRange(ref: string, values: CellValue[][]): void;
  }
  
  // Events
  export namespace events {
    export function onCellChanged(callback: (e: CellChangeEvent) => void): Disposable;
    export function onSelectionChanged(callback: (e: SelectionChangeEvent) => void): Disposable;
    export function onSheetActivated(callback: (e: SheetEvent) => void): Disposable;
    export function onWorkbookOpened(callback: (e: WorkbookEvent) => void): Disposable;
    export function onBeforeSave(callback: (e: SaveEvent) => void): Disposable;
  }
  
  // Commands
  export namespace commands {
    export function registerCommand(id: string, handler: (...args: any[]) => any): Disposable;
    export function executeCommand(id: string, ...args: any[]): Promise<any>;
  }
  
  // UI
  export namespace ui {
    export function showMessage(message: string, type?: "info" | "warning" | "error"): void;
    export function showInputBox(options: InputBoxOptions): Promise<string | undefined>;
    export function showQuickPick<T>(items: QuickPickItem<T>[], options?: QuickPickOptions): Promise<T | undefined>;
    export function createPanel(id: string, options: PanelOptions): Panel;
    export function registerContextMenu(menuId: string, items: MenuItem[]): Disposable;
  }
  
  // Storage
  export namespace storage {
    export function get<T>(key: string): T | undefined;
    export function set<T>(key: string, value: T): void;
    export function delete(key: string): void;
  }
  
  // Configuration
  export namespace config {
    export function get<T>(key: string): T | undefined;
    export function update(key: string, value: any): Promise<void>;
    export function onDidChange(callback: (e: ConfigChangeEvent) => void): Disposable;
  }
  
  // Context
  export namespace context {
    export const extensionPath: string;
    export const extensionUri: Uri;
    export const globalStoragePath: string;
    export const workspaceStoragePath: string;
  }
}

// Interfaces
interface Workbook {
  readonly name: string;
  readonly path: string | undefined;
  readonly sheets: Sheet[];
  readonly activeSheet: Sheet;
  
  save(): Promise<void>;
  saveAs(path: string): Promise<void>;
  close(): void;
}

interface Sheet {
  readonly name: string;
  readonly id: string;
  readonly usedRange: Range;
  
  getCell(row: number, col: number): Cell;
  getRange(ref: string): Range;
  
  activate(): void;
  rename(name: string): void;
}

interface Cell {
  readonly row: number;
  readonly col: number;
  readonly address: string;
  
  value: CellValue;
  formula: string | null;
  readonly displayValue: string;
  
  readonly style: CellStyle;
  readonly comment: Comment | null;
  
  clear(): void;
}

interface Range {
  readonly startRow: number;
  readonly startCol: number;
  readonly endRow: number;
  readonly endCol: number;
  readonly address: string;
  
  values: CellValue[][];
  formulas: (string | null)[][];
  
  getCells(): Cell[];
  clear(): void;
  applyStyle(style: Partial<CellStyle>): void;
}

interface Disposable {
  dispose(): void;
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
- **Memory caps (best-effort)**: extension workers are started with `worker_threads` `resourceLimits`
  based on a per-host `memoryMb` setting (default: 256MB). This caps the V8 heap, but does not cover
  all native/external allocations.
- **Crash/restart**: on crash/timeout, the extension is marked inactive and the next activation will
  spawn a fresh worker. Hosts can also explicitly recycle an extension via `reloadExtension(id)`.

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
import * as formula from "@formula/api";

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
import * as formula from "@formula/api";

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

```typescript
// Extension testing framework
import { TestRunner, TestSuite } from "@formula/test";

const suite: TestSuite = {
  name: "My Extension Tests",
  
  beforeEach: async (ctx) => {
    // Set up test workbook
    ctx.workbook = await formula.workbook.createWorkbook();
    ctx.sheet = ctx.workbook.sheets[0];
  },
  
  afterEach: async (ctx) => {
    ctx.workbook.close();
  },
  
  tests: [
    {
      name: "custom function returns correct value",
      test: async (ctx) => {
        ctx.sheet.getCell(0, 0).value = 5;
        ctx.sheet.getCell(0, 1).formula = "=MYFUNCTION(A1)";
        
        await formula.calculation.waitForRecalc();
        
        const result = ctx.sheet.getCell(0, 1).value;
        expect(result).toBe(10);
      }
    },
    
    {
      name: "panel displays correctly",
      test: async (ctx) => {
        await formula.commands.executeCommand("myExtension.openPanel");
        
        const panel = formula.ui.getPanel("myExtension.panel");
        expect(panel).toBeDefined();
        expect(panel.visible).toBe(true);
      }
    }
  ]
};

TestRunner.run(suite);
```
