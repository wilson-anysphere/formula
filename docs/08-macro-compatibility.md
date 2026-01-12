# Macro & Scripting Compatibility

## Overview

VBA macros represent decades of business logic investment. We must **preserve and eventually execute** VBA while offering a **modern scripting alternative** with Python and TypeScript. The transition path should be gradual and supported by AI-assisted migration tools.

---

## VBA Handling Strategy

### Compatibility Levels

| Level | Capability | Priority |
|-------|------------|----------|
| **L1: Preserve** | Round-trip vbaProject.bin unchanged | P0 |
| **L2: Parse** | Read and display VBA code | P0 |
| **L3: Analyze** | Understand dependencies and references | P1 |
| **L4: Execute** | Run VBA code with full compatibility | P2 |
| **L5: Migrate** | Automatically convert VBA to Python/TS | P2 |

### vbaProject.bin Structure

VBA is stored as an OLE compound document embedded in the xlsx:

```
vbaProject.bin (OLE container)
├── VBA/
│   ├── _VBA_PROJECT      # Project metadata, version info
│   ├── dir               # Module directory (compressed)
│   ├── Module1           # Module source (compressed)
│   ├── Module2           # ...
│   ├── ThisWorkbook      # Workbook code module
│   ├── Sheet1            # Sheet code module
│   └── UserForm1         # UserForm code (if any)
├── PROJECT               # Project properties
├── PROJECTwm             # Project web module
└── [Digital Signature]   # If signed
```

### Preservation Strategy

```typescript
class VBAPreserver {
  private vbaProjectBin: Uint8Array | null = null;
  private vbaProjectPath = "xl/vbaProject.bin";
  
  async loadFromXLSX(archive: ZipArchive): Promise<void> {
    if (archive.hasEntry(this.vbaProjectPath)) {
      this.vbaProjectBin = await archive.getEntry(this.vbaProjectPath);
    }
  }
  
  async saveToXLSX(archive: ZipArchive): Promise<void> {
    if (this.vbaProjectBin) {
      // Preserve byte-for-byte
      archive.setEntry(this.vbaProjectPath, this.vbaProjectBin);
      
      // Update content types
      this.updateContentTypes(archive);
      
      // Update relationships
      this.updateRelationships(archive);
    }
  }
  
  private updateContentTypes(archive: ZipArchive): void {
    const contentTypes = archive.getXml("[Content_Types].xml");
    
    // Add VBA content type if not present
    if (!contentTypes.includes("vbaProject")) {
      // Add: <Override PartName="/xl/vbaProject.bin" 
      //       ContentType="application/vnd.ms-office.vbaProject"/>
    }
  }
}
```

### VBA Parser

Parse VBA code for display and analysis (following MS-OVBA specification):

```typescript
interface VBAProject {
  name: string;
  modules: VBAModule[];
  references: VBAReference[];
  constants: Record<string, string>;
}

interface VBAModule {
  name: string;
  type: "standard" | "class" | "document" | "userform";
  code: string;
  attributes: Record<string, string>;
}

interface VBAReference {
  name: string;
  guid: string;
  major: number;
  minor: number;
  path?: string;
}

class VBAParser {
  parse(vbaProjectBin: Uint8Array): VBAProject {
    const ole = new OLEReader(vbaProjectBin);
    
    // Read dir stream (compressed)
    const dirStream = ole.getStream("VBA/dir");
    const decompressed = this.decompress(dirStream);
    const directory = this.parseDirectory(decompressed);
    
    // Read each module
    const modules: VBAModule[] = [];
    for (const moduleInfo of directory.modules) {
      const moduleStream = ole.getStream(`VBA/${moduleInfo.name}`);
      const code = this.decompressModule(moduleStream, moduleInfo.offset);
      
      modules.push({
        name: moduleInfo.name,
        type: moduleInfo.type,
        code,
        attributes: this.parseAttributes(code)
      });
    }
    
    return {
      name: directory.projectName,
      modules,
      references: directory.references,
      constants: directory.constants
    };
  }
  
  private decompress(data: Uint8Array): Uint8Array {
    // MS-OVBA compression algorithm
    // Each chunk: signature byte (0x01) + compressed data
    const chunks: Uint8Array[] = [];
    let offset = 0;
    
    while (offset < data.length) {
      const signature = data[offset];
      if (signature !== 0x01) throw new Error("Invalid VBA compression");
      
      const chunkSize = (data[offset + 1] | (data[offset + 2] << 8)) & 0x0FFF;
      const isCompressed = (data[offset + 1] >> 4) & 0x01;
      
      if (isCompressed) {
        chunks.push(this.decompressChunk(data.slice(offset + 3, offset + 3 + chunkSize)));
      } else {
        chunks.push(data.slice(offset + 3, offset + 3 + chunkSize));
      }
      
      offset += 3 + chunkSize;
    }
    
    return this.concatenate(chunks);
  }
}
```

### VBA Viewer UI

```typescript
class VBAViewer {
  private project: VBAProject | null = null;
  
  render(): React.ReactNode {
    if (!this.project) {
      return <div>No VBA project in this workbook</div>;
    }
    
    return (
      <div className="vba-viewer">
        <div className="project-tree">
          <TreeNode label={this.project.name} expanded>
            <TreeNode label="References">
              {this.project.references.map(ref => (
                <TreeNode key={ref.name} label={ref.name} icon="library" />
              ))}
            </TreeNode>
            <TreeNode label="Modules">
              {this.project.modules.map(mod => (
                <TreeNode 
                  key={mod.name} 
                  label={mod.name}
                  icon={this.getModuleIcon(mod.type)}
                  onClick={() => this.selectModule(mod)}
                />
              ))}
            </TreeNode>
          </TreeNode>
        </div>
        
        <div className="code-editor">
          <CodeEditor
            language="vba"
            value={this.selectedModule?.code || ""}
            readOnly={true}
            options={{
              minimap: { enabled: false },
              lineNumbers: "on",
              folding: true
            }}
          />
        </div>
      </div>
    );
  }
}
```

---

## VBA Execution (desktop - current)

The desktop (Tauri) app includes a first-pass VBA runtime integration. VBA execution is handled in
the Rust backend (so the webview stays responsive), and cell edits produced by macros are applied
back into the UI.

### Automatic event macros

The desktop app automatically fires the following Excel/VBA event macros (subject to Trust Center
policy and sandbox permissions):

- `Workbook_Open`
- `Worksheet_Change` (debounced; per-sheet bounding box)
- `Worksheet_SelectionChange` (debounced)
- `Workbook_BeforeClose` (window close + tray quit)

Important behavioral notes (current implementation):

- Macro-driven cell updates are applied to the `DocumentController` as **external deltas** (not
  undoable) and are tagged so the desktop workbook sync bridge does **not** echo them back into the
  backend.
- While applying macro updates, `Worksheet_Change` event macros are suppressed to prevent runaway
  recursion (infinite loops when event handlers write cells).

---

## VBA Execution (Future)

### Execution Approaches

| Approach | Pros | Cons |
|----------|------|------|
| **Interpreter** | Full control, no external deps | Complex to build, slow |
| **Transpilation** | Leverage existing runtimes | Semantic gaps |
| **COM Interop** | Native Excel compatibility | Windows only, security |
| **WASM VBA Runtime** | Cross-platform | Doesn't exist yet |

### Interpreter Architecture (Recommended)

```typescript
interface VBAInterpreter {
  // Parse VBA into AST
  parse(code: string): VBAProgram;
  
  // Execute with spreadsheet context
  execute(program: VBAProgram, context: SpreadsheetContext): ExecutionResult;
  
  // Event handlers
  onWorkbook_Open?: () => void;
  onWorkbook_BeforeClose?: (cancel: { value: boolean }) => void;
  onWorksheet_Change?: (target: Range) => void;
  onWorksheet_SelectionChange?: (target: Range) => void;
}

class VBAExecutionEngine implements VBAInterpreter {
  private globalScope: Scope;
  private modules: Map<string, CompiledModule>;
  private spreadsheet: SpreadsheetAPI;
  
  async execute(program: VBAProgram, context: SpreadsheetContext): Promise<ExecutionResult> {
    this.spreadsheet = context.api;
    
    // Initialize global scope with VBA built-ins
    this.globalScope = new Scope();
    this.registerBuiltIns();
    
    // Compile all modules
    for (const module of program.modules) {
      this.modules.set(module.name, this.compileModule(module));
    }
    
    // Execute entry point (if any)
    if (context.entryPoint) {
      return this.callProcedure(context.entryPoint, context.args || []);
    }
    
    return { success: true };
  }
  
  private registerBuiltIns(): void {
    // Register VBA built-in functions
    this.globalScope.set("MsgBox", this.builtIn_MsgBox.bind(this));
    this.globalScope.set("InputBox", this.builtIn_InputBox.bind(this));
    this.globalScope.set("Range", this.builtIn_Range.bind(this));
    this.globalScope.set("Cells", this.builtIn_Cells.bind(this));
    this.globalScope.set("ActiveSheet", () => this.spreadsheet.activeSheet);
    this.globalScope.set("ActiveCell", () => this.spreadsheet.activeCell);
    // ... many more
  }
  
  private async builtIn_MsgBox(
    prompt: string,
    buttons?: number,
    title?: string
  ): Promise<number> {
    // Show dialog and return button clicked
    return await this.ui.showMessageBox({
      message: prompt,
      buttons: this.parseButtons(buttons || 0),
      title: title || "Microsoft Excel"
    });
  }
  
  private builtIn_Range(ref: string): VBARange {
    const range = this.spreadsheet.parseReference(ref);
    return new VBARange(this.spreadsheet, range);
  }
}
```

### VBA Object Model

```typescript
// Implement Excel object model for VBA compatibility
class VBARange {
  constructor(
    private api: SpreadsheetAPI,
    private range: Range
  ) {}
  
  get Value(): any {
    if (this.isSingleCell) {
      return this.api.getCellValue(this.range.startRow, this.range.startCol);
    }
    return this.api.getRangeValues(this.range);
  }
  
  set Value(value: any) {
    if (this.isSingleCell) {
      this.api.setCellValue(this.range.startRow, this.range.startCol, value);
    } else if (Array.isArray(value)) {
      this.api.setRangeValues(this.range, value);
    } else {
      // Fill all cells with value
      this.api.fillRange(this.range, value);
    }
  }
  
  get Formula(): string {
    return this.api.getCellFormula(this.range.startRow, this.range.startCol);
  }
  
  set Formula(formula: string) {
    this.api.setCellFormula(this.range.startRow, this.range.startCol, formula);
  }
  
  get Rows(): VBARange {
    return new VBARange(this.api, { ...this.range, type: "rows" });
  }
  
  get Columns(): VBARange {
    return new VBARange(this.api, { ...this.range, type: "columns" });
  }
  
  Select(): void {
    this.api.setSelection(this.range);
  }
  
  Copy(destination?: VBARange): void {
    if (destination) {
      this.api.copyRange(this.range, destination.range);
    } else {
      this.api.copyToClipboard(this.range);
    }
  }
  
  AutoFill(destination: VBARange, type?: number): void {
    this.api.autoFill(this.range, destination.range, type);
  }
}

class VBAWorksheet {
  constructor(
    private api: SpreadsheetAPI,
    private sheetId: string
  ) {}
  
  get Name(): string {
    return this.api.getSheetName(this.sheetId);
  }
  
  set Name(value: string) {
    this.api.renameSheet(this.sheetId, value);
  }
  
  Range(ref: string): VBARange {
    return new VBARange(this.api, this.api.parseReference(ref, this.sheetId));
  }
  
  Cells(row: number, col: number): VBARange {
    return new VBARange(this.api, { 
      startRow: row - 1,  // VBA is 1-indexed
      startCol: col - 1,
      endRow: row - 1,
      endCol: col - 1
    });
  }
  
  Activate(): void {
    this.api.activateSheet(this.sheetId);
  }
}
```

---

## Modern Scripting: Python

### Python Integration Architecture

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  USER SCRIPT                                                                │
│  import formula                                                             │
│  sheet = formula.active_sheet                                               │
│  sheet["A1"] = "Hello"                                                      │
│  df = sheet["A1:D100"].to_dataframe()                                       │
└─────────────────────────────┬───────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  FORMULA PYTHON PACKAGE                                                     │
│  ├── Spreadsheet API bindings                                               │
│  ├── Pandas/NumPy integration                                               │
│  └── IPC to main application                                                │
└─────────────────────────────┬───────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  PYTHON RUNTIME                                                             │
│  ├── Pyodide (WASM) - in-browser                                           │
│  ├── Native Python - desktop                                                │
│  └── Sandboxed execution environment                                        │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Python API

```python
# formula/api.py

from typing import Any, List, Optional, Union
import pandas as pd
import numpy as np

class Sheet:
    """Represents a worksheet in the spreadsheet."""
    
    def __init__(self, sheet_id: str, api: 'SpreadsheetAPI'):
        self._id = sheet_id
        self._api = api
    
    @property
    def name(self) -> str:
        return self._api.get_sheet_name(self._id)
    
    @name.setter
    def name(self, value: str):
        self._api.rename_sheet(self._id, value)
    
    def __getitem__(self, key: str) -> 'Range':
        """Access cells via sheet["A1"] or sheet["A1:B10"]."""
        return Range(self._api.parse_reference(key, self._id), self._api)
    
    def __setitem__(self, key: str, value: Any):
        """Set cell values via sheet["A1"] = value."""
        range_ref = self._api.parse_reference(key, self._id)
        if isinstance(value, pd.DataFrame):
            self._api.set_dataframe(range_ref, value)
        elif isinstance(value, (list, np.ndarray)):
            self._api.set_range_values(range_ref, value)
        else:
            self._api.set_cell_value(range_ref, value)


class Range:
    """Represents a range of cells."""
    
    def __init__(self, range_ref: dict, api: 'SpreadsheetAPI'):
        self._ref = range_ref
        self._api = api
    
    @property
    def value(self) -> Any:
        """Get the value(s) in this range."""
        return self._api.get_range_values(self._ref)
    
    @value.setter
    def value(self, val: Any):
        """Set the value(s) in this range."""
        self._api.set_range_values(self._ref, val)
    
    @property
    def formula(self) -> str:
        """Get the formula (for single cell)."""
        return self._api.get_cell_formula(self._ref)
    
    @formula.setter
    def formula(self, val: str):
        """Set the formula."""
        self._api.set_cell_formula(self._ref, val)
    
    def to_dataframe(self, header: bool = True) -> pd.DataFrame:
        """Convert range to pandas DataFrame."""
        data = self._api.get_range_values(self._ref)
        if header:
            return pd.DataFrame(data[1:], columns=data[0])
        return pd.DataFrame(data)
    
    def from_dataframe(self, df: pd.DataFrame, include_header: bool = True):
        """Write DataFrame to this range."""
        self._api.set_dataframe(self._ref, df, include_header)
    
    def clear(self):
        """Clear the range contents."""
        self._api.clear_range(self._ref)
    
    def apply_formula(self, formula_template: str):
        """Apply a formula pattern to each row."""
        self._api.apply_formula_column(self._ref, formula_template)


# Global convenience functions
def get_active_sheet() -> Sheet:
    """Get the currently active sheet."""
    return _workbook.active_sheet

def get_sheet(name: str) -> Sheet:
    """Get a sheet by name."""
    return _workbook.get_sheet(name)

def create_sheet(name: str, index: Optional[int] = None) -> Sheet:
    """Create a new sheet (optionally inserting at a specific 0-based index)."""
    return _workbook.create_sheet(name, index=index)

# Decorators for custom functions
def custom_function(func):
    """Register a Python function as a spreadsheet function."""
    _function_registry.register(func)
    return func

@custom_function
def PYSUM(range_values: List[List[float]]) -> float:
    """Example custom function: sum using numpy."""
    return np.sum(range_values)
```

### Python Runtime Options

#### Option 1: Pyodide (In-Browser)

Pyodide can run in two modes:

- **Worker-backed (preferred)**: keeps the UI responsive, but requires a cross-origin isolated context (COOP/COEP) so `SharedArrayBuffer` is available.
- **Main-thread fallback**: works without COOP/COEP, but Python execution will block the UI thread (acceptable as a degraded mode for embedded/webview deployments).

```typescript
class PyodideRuntime {
  private pyodide: any;
  
  async initialize(): Promise<void> {
    this.pyodide = await loadPyodide({
      indexURL: "https://cdn.jsdelivr.net/pyodide/v0.24.0/full/"
    });
    
    // Load common packages
    await this.pyodide.loadPackage(["numpy", "pandas"]);
    
    // Install formula package
    await this.pyodide.runPythonAsync(`
      import micropip
      await micropip.install('formula-api')
    `);
  }
  
  async runScript(code: string, context: ScriptContext): Promise<ScriptResult> {
    // Set up API bridge
    this.pyodide.registerJsModule("formula_bridge", {
      get_cell: (row: number, col: number) => context.api.getCell(row, col),
      set_cell: (row: number, col: number, value: any) => context.api.setCell(row, col, value),
      // ... more API methods
    });
    
    // Run the script
    try {
      const result = await this.pyodide.runPythonAsync(code);
      return { success: true, result };
    } catch (error) {
      return { success: false, error: error.message };
    }
  }
}
```

#### Option 2: Native Python (Desktop)

```typescript
class NativePythonRuntime {
  private pythonProcess: ChildProcess | null = null;
  
  async initialize(): Promise<void> {
    // Find Python installation
    const pythonPath = await this.findPython();
    
    // Start Python subprocess with our API server
    this.pythonProcess = spawn(pythonPath, [
      "-m", "formula.runtime",
      "--port", String(this.port)
    ]);
    
    // Wait for server to start
    await this.waitForServer();
  }
  
  async runScript(code: string, context: ScriptContext): Promise<ScriptResult> {
    const response = await fetch(`http://localhost:${this.port}/execute`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        code,
        context: this.serializeContext(context)
      })
    });
    
    return response.json();
  }
}
```

---

## Modern Scripting: TypeScript

### TypeScript API

```typescript
// @formula/scripting

export interface Sheet {
  readonly name: string;
  readonly id: string;
  
  cell(row: number, col: number): Cell;
  range(ref: string): Range;
  
  getUsedRange(): Range;
  clear(): void;
}

export interface Cell {
  value: CellValue;
  formula: string | null;
  
  readonly row: number;
  readonly col: number;
  readonly address: string;
  
  clear(): void;
}

export interface Range {
  values: CellValue[][];
  formulas: (string | null)[][];
  
  readonly rowCount: number;
  readonly colCount: number;
  
  clear(): void;
  fill(value: CellValue): void;
  applyFormula(template: string): void;
  
  toJSON(): any[];
  toCSV(): string;
}

// Script context provided to user scripts
export interface ScriptContext {
  workbook: Workbook;
  activeSheet: Sheet;
  selection: Range;
  
  // UI helpers
  alert(message: string): Promise<void>;
  confirm(message: string): Promise<boolean>;
  prompt(message: string, defaultValue?: string): Promise<string | null>;
  
  // Utilities
  fetch: typeof fetch;
  console: Console;
}

// User script example
export default async function main(ctx: ScriptContext) {
  const sheet = ctx.activeSheet;
  
  // Read data
  const data = sheet.range("A1:D100").values;
  
  // Process
  const processed = data.map(row => ({
    name: row[0],
    total: Number(row[1]) + Number(row[2]) + Number(row[3])
  }));
  
  // Write results
  sheet.range("F1").value = "Name";
  sheet.range("G1").value = "Total";
  
  processed.forEach((item, i) => {
    sheet.cell(i + 2, 6).value = item.name;
    sheet.cell(i + 2, 7).value = item.total;
  });
}
```

### Script Editor

```typescript
class ScriptEditor {
  private editor: monaco.editor.IStandaloneCodeEditor;
  private runtime: ScriptRuntime;
  
  constructor(container: HTMLElement) {
    this.editor = monaco.editor.create(container, {
      language: "typescript",
      theme: "vs-dark",
      automaticLayout: true,
      minimap: { enabled: false }
    });
    
    // Add Formula API types
    this.addTypeDefinitions();
    
    // Set up auto-complete
    this.setupAutoComplete();
  }
  
  private addTypeDefinitions(): void {
    monaco.languages.typescript.typescriptDefaults.addExtraLib(
      FORMULA_API_TYPES,
      "formula.d.ts"
    );
  }
  
  async runScript(): Promise<void> {
    const code = this.editor.getValue();
    
    try {
      // Compile TypeScript
      const compiled = await this.compile(code);
      
      // Execute in sandbox
      const result = await this.runtime.execute(compiled);
      
      if (result.error) {
        this.showError(result.error);
      } else {
        this.showSuccess("Script completed successfully");
      }
    } catch (error) {
      this.showError(error);
    }
  }
}
```

---

## Macro Recorder

### Recording User Actions

```typescript
interface RecordedAction {
  type: string;
  timestamp: number;
  params: any;
}

class MacroRecorder {
  private recording = false;
  private actions: RecordedAction[] = [];
  
  startRecording(): void {
    this.recording = true;
    this.actions = [];
    
    // Hook into spreadsheet events
    this.spreadsheet.on("cellChanged", this.onCellChanged.bind(this));
    this.spreadsheet.on("selectionChanged", this.onSelectionChanged.bind(this));
    this.spreadsheet.on("rangeFormatted", this.onRangeFormatted.bind(this));
    // ... more events
  }
  
  stopRecording(): RecordedAction[] {
    this.recording = false;
    return this.optimizeActions(this.actions);
  }
  
  private onCellChanged(event: CellChangeEvent): void {
    if (!this.recording) return;
    
    this.actions.push({
      type: "setCellValue",
      timestamp: Date.now(),
      params: {
        cell: cellToAddress(event.cell),
        value: event.newValue,
        formula: event.formula
      }
    });
  }
  
  private optimizeActions(actions: RecordedAction[]): RecordedAction[] {
    // Combine consecutive similar actions
    // e.g., multiple cell changes in same range → setRangeValues
    const optimized: RecordedAction[] = [];
    
    for (let i = 0; i < actions.length; i++) {
      const action = actions[i];
      
      // Look for consecutive setCellValue actions in adjacent cells
      if (action.type === "setCellValue") {
        const batch = this.findBatch(actions, i);
        if (batch.length > 1) {
          optimized.push(this.combineToBatchAction(batch));
          i += batch.length - 1;
          continue;
        }
      }
      
      optimized.push(action);
    }
    
    return optimized;
  }
  
  generatePython(actions: RecordedAction[]): string {
    const lines: string[] = [
      "import formula",
      "",
      "def main():",
      "    sheet = formula.active_sheet",
      ""
    ];
    
    for (const action of actions) {
      lines.push(`    ${this.actionToPython(action)}`);
    }
    
    lines.push("");
    lines.push("if __name__ == '__main__':");
    lines.push("    main()");
    
    return lines.join("\n");
  }
  
  generateTypeScript(actions: RecordedAction[]): string {
    const lines: string[] = [
      "import { ScriptContext } from '@formula/scripting';",
      "",
      "export default async function main(ctx: ScriptContext) {",
      "  const sheet = ctx.activeSheet;",
      ""
    ];
    
    for (const action of actions) {
      lines.push(`  ${this.actionToTypeScript(action)}`);
    }
    
    lines.push("}");
    
    return lines.join("\n");
  }
  
  private actionToPython(action: RecordedAction): string {
    switch (action.type) {
      case "setCellValue":
        if (action.params.formula) {
          return `sheet["${action.params.cell}"].formula = "${action.params.formula}"`;
        }
        return `sheet["${action.params.cell}"] = ${JSON.stringify(action.params.value)}`;
        
      case "setRangeValues":
        return `sheet["${action.params.range}"].value = ${JSON.stringify(action.params.values)}`;
        
      case "applyFormat":
        return `sheet["${action.params.range}"].format(${JSON.stringify(action.params.format)})`;
        
      default:
        return `# Unknown action: ${action.type}`;
    }
  }
}
```

---

## AI-Assisted Migration

### VBA to Python Conversion

```typescript
class VBAMigrator {
  async migrateModule(vbaCode: string): Promise<MigrationResult> {
    // Parse VBA
    const vbaAST = this.parseVBA(vbaCode);
    
    // Analyze dependencies
    const deps = this.analyzeDependencies(vbaAST);
    
    // Convert with AI assistance
    const pythonCode = await this.convertToPython(vbaAST, deps);
    
    // Validate conversion
    const validation = await this.validateConversion(vbaCode, pythonCode);
    
    return {
      originalCode: vbaCode,
      convertedCode: pythonCode,
      warnings: validation.warnings,
      unsupportedFeatures: validation.unsupported,
      manualReviewRequired: validation.needsReview
    };
  }
  
  private async convertToPython(ast: VBAProgram, deps: Dependencies): Promise<string> {
    const prompt = `
Convert this VBA code to Python using the Formula spreadsheet API.

VBA Code:
${this.astToCode(ast)}

Available Python API:
- sheet = formula.active_sheet
- sheet["A1"] or sheet.range("A1:B10")
- range.value, range.formula
- range.to_dataframe(), range.from_dataframe(df)
- formula.get_sheet(name), formula.create_sheet(name, index=None)
- @formula.custom_function decorator for UDFs

Convert the code, preserving the logic and making it idiomatic Python.
Note any VBA features that cannot be directly converted.
`;
    
    const response = await this.llm.complete(prompt);
    return this.extractCode(response);
  }
}
```

### Migration Report

```typescript
interface MigrationReport {
  summary: {
    totalModules: number;
    successfullyConverted: number;
    partiallyConverted: number;
    failed: number;
  };
  
  modules: ModuleMigrationResult[];
  
  recommendations: Recommendation[];
}

interface ModuleMigrationResult {
  moduleName: string;
  originalLines: number;
  convertedLines: number;
  status: "success" | "partial" | "failed";
  
  unsupportedFeatures: UnsupportedFeature[];
  warnings: Warning[];
  
  originalCode: string;
  convertedCode: string;
}

interface UnsupportedFeature {
  feature: string;
  location: { line: number; col: number };
  reason: string;
  workaround?: string;
}
```

---

## Security

### Trust Center (desktop)

The desktop app implements a minimal "Trust Center" policy layer for VBA macro execution:

- `blocked`: macros never run.
- `trusted_once` / `trusted_always`: macros run based on a workbook fingerprint allow-list.
- `trusted_signed_only`: macros run **only** when the workbook's VBA project signature is
  **cryptographically verified and bound** to the VBA project contents (MS-OVBA §2.4.2 + MS-OSHARED §4.3).

Important notes:

- Signature verification is performed against the embedded PKCS#7/CMS structure inside the
  `\x05DigitalSignature*` OLE stream(s). This is intended to prevent treating "has a signature blob"
  as "is signed".
- Signature verification includes *binding* the signature to the VBA project contents by extracting
  the signed digest (`SpcIndirectDataContent` → `DigestInfo.digest`) and comparing it to a freshly
  computed MS-OVBA-style digest:
  - `Content Hash` (MS-OVBA §2.4.2.3), computed over `ContentNormalizedData` (normalized module
    source and select metadata), and
  - `Agile Content Hash` (MS-OVBA §2.4.2.4), which extends the hash transcript with
    `FormsNormalizedData` (designer/UserForm storages).
  - For the `DigitalSignatureExt` stream variant, Office uses the MS-OVBA §2.4.2 v3 transcript and
    computes the **V3 Content Hash** (MS-OVBA §2.4.2.7):
    `V3ContentHash = SHA-256(ProjectNormalizedData)`, where
    `ProjectNormalizedData = V3ContentNormalizedData || FormsNormalizedData`.
  
  Per MS-OSHARED §4.3, the digest bytes embedded in **legacy** VBA signature streams
  (`DigitalSignature` / `DigitalSignatureEx`) are always **MD5 (16 bytes)** even when the PKCS#7/CMS
  signature uses SHA-256 and even when `DigestInfo.digestAlgorithm.algorithm` indicates SHA-256 (the
  OID is informational for v1/v2 VBA binding). For the newest `DigitalSignatureExt` stream, Office
  uses the MS-OVBA v3 digest over v3 `ProjectNormalizedData` and the digest algorithm OID is
  meaningful (typically SHA-256).
  
  This binding check is exposed via `formula-vba` as `VbaDigitalSignature::binding`. The desktop
  Trust Center treats a VBA project as "signed" only when the PKCS#7/CMS signature verifies **and**
  the digest binding check reports `Bound`.
  - For callers that need more detail (hash algorithm OID/name, signed digest bytes, computed digest
    bytes), use `formula_vba::verify_vba_digital_signature_bound`.
- Unknown/unparseable digest structures are treated conservatively as unverified (`binding == Unknown`).
- Binding verification is deterministic best-effort and may not match Excel/MS-OVBA for all
  real-world files yet; this policy can produce **false negatives** (a legitimately signed workbook
  treated as untrusted). We prefer false negatives over false positives for `trusted_signed_only`.
- Certificate chain trust ("trusted publisher") is not enforced by default (OpenSSL `NOVERIFY`), but
  can be evaluated as an **opt-in** step by calling `formula-vba`'s
  `verify_vba_digital_signature_with_trust` with an explicit root certificate set.
- Timestamp validation may be incomplete.

#### VBA digital signatures: stream location, payload variants, and digest binding

Excel stores VBA macro signatures **inside** `xl/vbaProject.bin` (an OLE compound document), in one
of the control-character-prefixed streams:

- `\x05DigitalSignature`
- `\x05DigitalSignatureEx`
- `\x05DigitalSignatureExt`

If more than one signature stream exists, Excel prefers the newest stream:
`DigitalSignatureExt` → `DigitalSignatureEx` → `DigitalSignature`.
(This stream-name precedence is not normatively specified in MS-OVBA.)

Edge case: some producers store a **storage** named `\x05DigitalSignature*` containing a nested
stream (e.g. `\x05DigitalSignature/sig`). When searching for signatures, match on any path
component, not just root-level streams.

Common signature stream payload shapes we must handle:
 
1. **Raw PKCS#7/CMS DER** (`ContentInfo`, usually begins with ASN.1 `SEQUENCE (0x30)`).
2. **Office DigSig wrapper**, which contains (or points to) the DER-encoded PKCS#7/CMS payload:
   - MS-OSHARED `DigSigBlob` / `DigSigInfoSerialized` (offset-based), or
   - a length-prefixed DigSigInfoSerialized-like header commonly seen in the wild.
3. **Detached `content || pkcs7`**, where the stream is `signed_content_bytes` followed by a
   detached PKCS#7 signature over those bytes.

How to obtain the signed digest for MS-OVBA signature binding:

- Parse the PKCS#7/CMS `ContentInfo`. The `contentType` is typically `signedData`
  (`1.2.840.113549.1.7.2`).
- From `SignedData.encapContentInfo`, read the `eContent` bytes (if embedded). For Authenticode-like
  signatures, `eContentType` is typically `SpcIndirectDataContent` (`1.3.6.1.4.1.311.2.1.4`).
  - In the detached `content || pkcs7` variant, the detached `content` prefix plays the same role
    as `eContent`.
- Decode the `SpcIndirectDataContent` / `SpcIndirectDataContentV2` structure and extract its
  VBA signature binding digest bytes:
  - `SpcIndirectDataContent`: `messageDigest: DigestInfo.digest`
  - `SpcIndirectDataContentV2`: `SigDataV1Serialized.sourceHash`
  - digest bytes are expected to be:
    - 16-byte MD5 for legacy v1/v2 signature streams (`DigitalSignature` / `DigitalSignatureEx`)
      per MS-OSHARED §4.3 (even when the algorithm OID indicates SHA-256), or
    - 32-byte SHA-256 for the v3 `DigitalSignatureExt` stream (MS-OVBA §2.4.2).

Binding (best-effort; see MS-OVBA):

1. Determine which signature stream variant is being verified:
   - v1/v2: `DigitalSignature` / `DigitalSignatureEx`
   - v3: `DigitalSignatureExt`
2. For v1/v2 streams, compute the MS-OVBA `Content Hash` = `MD5(ContentNormalizedData)` (MS-OVBA §2.4.2.3)
   and the MS-OVBA `Agile Content Hash` = `MD5(ContentNormalizedData || FormsNormalizedData)` (MS-OVBA §2.4.2.4),
   then compare the signed digest bytes to either digest.
3. For the v3 `DigitalSignatureExt` stream, compute the MS-OVBA **V3 Content Hash**
   `V3ContentHash = SHA-256(ProjectNormalizedData)` (MS-OVBA §2.4.2.7) and compare it to the signed
   digest bytes.
4. `trusted_signed_only` is treated as satisfied only when:
    - the PKCS#7/CMS signature verifies (`SignedVerified`), **and**
    - the digest comparison matches (`VbaSignatureBinding::Bound`).

If the PKCS#7/CMS signature verifies but the digest comparison fails, the signature is treated as
present-but-invalid for Trust Center purposes. If binding cannot be verified (`Unknown`), it is
treated conservatively as unverified.

For more detail, see [`vba-digital-signatures.md`](./vba-digital-signatures.md).

Relevant specs:

- MS-OVBA (VBA project storage + Contents Hash / Agile Content Hash): https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/
- MS-OSHARED (VBA digital signature storage / DigSigBlob + DigSigInfoSerialized + MD5 VBA project hash rule): https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/

### Script Sandboxing

```typescript
interface SandboxPermissions {
  fileSystem: "none" | "read" | "readwrite";
  network: "none" | "allowlist" | "full";
  networkAllowlist?: string[];
  clipboard: boolean;
  notifications: boolean;
  maxExecutionTime: number;  // ms
  maxMemory: number;  // bytes
}

class ScriptSandbox {
  private permissions: SandboxPermissions;
  
  async execute(code: string, context: ScriptContext): Promise<ScriptResult> {
    const worker = new Worker("script-sandbox-worker.js");
    
    // Set up permission checks
    const api = this.createRestrictedAPI(context);
    
    // Set timeout
    const timeoutId = setTimeout(() => {
      worker.terminate();
      throw new Error("Script execution timed out");
    }, this.permissions.maxExecutionTime);
    
    try {
      const result = await new Promise((resolve, reject) => {
        worker.postMessage({ code, api: this.serializeAPI(api) });
        worker.onmessage = (e) => resolve(e.data);
        worker.onerror = (e) => reject(e);
      });
      
      return result;
    } finally {
      clearTimeout(timeoutId);
      worker.terminate();
    }
  }
  
  private createRestrictedAPI(context: ScriptContext): RestrictedAPI {
    return {
      getCell: (row: number, col: number) => {
        return context.api.getCell(row, col);
      },
      setCell: (row: number, col: number, value: any) => {
        return context.api.setCell(row, col, value);
      },
      fetch: async (url: string, options?: RequestInit) => {
        if (this.permissions.network === "none") {
          throw new Error("Network access not permitted");
        }
        if (this.permissions.network === "allowlist") {
          const urlObj = new URL(url);
          if (!this.permissions.networkAllowlist?.includes(urlObj.hostname)) {
            throw new Error(`Network access to ${urlObj.hostname} not permitted`);
          }
        }
        return fetch(url, options);
      }
      // ... more restricted methods
    };
  }
}
```

---

## Testing Strategy

### VBA Compatibility Tests

```typescript
describe("VBA Preservation", () => {
  it("preserves vbaProject.bin on round-trip", async () => {
    const original = await readFile("test-with-macro.xlsm");
    const workbook = await loadWorkbook(original);
    const saved = await saveWorkbook(workbook);
    const reloaded = await loadWorkbook(saved);
    
    // Extract vbaProject.bin from both
    const originalVBA = extractVBAProject(original);
    const savedVBA = extractVBAProject(saved);
    
    // Should be byte-for-byte identical
    expect(savedVBA).toEqual(originalVBA);
  });
});

describe("VBA Parser", () => {
  it("parses simple Sub", () => {
    const code = `
Sub HelloWorld()
    MsgBox "Hello, World!"
End Sub
`;
    const result = parser.parse(code);
    
    expect(result.modules[0].procedures[0].name).toBe("HelloWorld");
  });
});

describe("Script Execution", () => {
  it("executes Python script correctly", async () => {
    const script = `
sheet = formula.active_sheet
sheet["A1"] = 42
sheet["A2"] = "=A1*2"
`;
    
    await runtime.execute(script);
    
    expect(sheet.getCell(0, 0).value).toBe(42);
    expect(sheet.getCell(1, 0).value).toBe(84);
  });
});
```
