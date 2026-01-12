import { BrowserExtensionHost } from "@formula/extension-host/browser";
import { MarketplaceClient, WebExtensionManager } from "@formula/extension-marketplace";

type CellValue = string | number | boolean | null;

class TestSpreadsheetApi {
  private readonly _cells = new Map<string, CellValue>();
  private _selection = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };

  private readonly _selectionListeners = new Set<(e: { selection: ReturnType<TestSpreadsheetApi["getSelection"]> }) => void>();
  private readonly _cellListeners = new Set<(e: { row: number; col: number; value: CellValue }) => void>();

  private _key(row: number, col: number): string {
    return `${row},${col}`;
  }

  getCell(row: number, col: number): CellValue {
    const key = this._key(row, col);
    return this._cells.has(key) ? (this._cells.get(key) as CellValue) : null;
  }

  async setCell(row: number, col: number, value: CellValue): Promise<void> {
    this._cells.set(this._key(row, col), value);
    const evt = { row, col, value };
    for (const listener of [...this._cellListeners]) {
      try {
        listener(evt);
      } catch {
        // ignore
      }
    }
  }

  setSelection(range: { startRow: number; startCol: number; endRow: number; endCol: number }): void {
    this._selection = { ...range };
    const evt = { selection: this.getSelection() };
    for (const listener of [...this._selectionListeners]) {
      try {
        listener(evt);
      } catch {
        // ignore
      }
    }
  }

  getSelection() {
    const { startRow, startCol, endRow, endCol } = this._selection;
    const values: CellValue[][] = [];
    for (let r = startRow; r <= endRow; r++) {
      const row: CellValue[] = [];
      for (let c = startCol; c <= endCol; c++) {
        row.push(this.getCell(r, c));
      }
      values.push(row);
    }
    return { startRow, startCol, endRow, endCol, values };
  }

  onSelectionChanged(callback: (e: { selection: ReturnType<TestSpreadsheetApi["getSelection"]> }) => void) {
    this._selectionListeners.add(callback);
    return { dispose: () => this._selectionListeners.delete(callback) };
  }

  onCellChanged(callback: (e: { row: number; col: number; value: CellValue }) => void) {
    this._cellListeners.add(callback);
    return { dispose: () => this._cellListeners.delete(callback) };
  }
}

export async function setupExtensionTestHarness(): Promise<void> {
  if (typeof window === "undefined") return;
  const params = new URLSearchParams(window.location.search);
  if (!params.has("extTest")) return;

  const spreadsheet = new TestSpreadsheetApi();
  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: spreadsheet,
    permissionPrompt: async () => true
  });

  const marketplaceClient = new MarketplaceClient({ baseUrl: "/api" });
  const manager = new WebExtensionManager({ marketplaceClient, host });

  const api = {
    async installSampleHello() {
      const id = "formula.sample-hello";
      await manager.install(id);
      await manager.loadInstalled(id);
      return id;
    },
    async executeCommand(commandId: string, ...args: any[]) {
      return host.executeCommand(commandId, ...args);
    },
    setCell(row: number, col: number, value: CellValue) {
      return spreadsheet.setCell(row, col, value);
    },
    getCell(row: number, col: number) {
      return spreadsheet.getCell(row, col);
    },
    setSelection(range: { startRow: number; startCol: number; endRow: number; endCol: number }) {
      spreadsheet.setSelection(range);
    },
    getMessages() {
      return host.getMessages();
    },
    async dispose() {
      await manager.dispose();
      await host.dispose();
    }
  };

  (window as any).__formulaExtensionTest = api;
}
