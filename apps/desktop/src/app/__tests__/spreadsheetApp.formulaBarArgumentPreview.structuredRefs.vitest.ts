/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";

let priorGridMode: string | undefined;

function createInMemoryLocalStorage(): Storage {
  const store = new Map<string, string>();
  return {
    getItem: (key: string) => (store.has(key) ? store.get(key)! : null),
    setItem: (key: string, value: string) => {
      store.set(String(key), String(value));
    },
    removeItem: (key: string) => {
      store.delete(String(key));
    },
    clear: () => {
      store.clear();
    },
    key: (index: number) => Array.from(store.keys())[index] ?? null,
    get length() {
      return store.size;
    },
  } as Storage;
}

function createMockCanvasContext(): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;
  const context = new Proxy(
    {
      canvas: document.createElement("canvas"),
      measureText: (text: string) => ({ width: text.length * 8 }),
      createLinearGradient: () => gradient,
      createPattern: () => null,
      getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
      putImageData: noop,
    },
    {
      get(target, prop) {
        if (prop in target) return (target as any)[prop];
        return noop;
      },
      set(target, prop, value) {
        (target as any)[prop] = value;
        return true;
      },
    },
  );
  return context as any;
}

function createRoot(): HTMLElement {
  const root = document.createElement("div");
  root.tabIndex = 0;
  root.getBoundingClientRect = () =>
    ({
      width: 800,
      height: 600,
      left: 0,
      top: 0,
      right: 800,
      bottom: 600,
      x: 0,
      y: 0,
      toJSON: () => {},
    }) as any;
  document.body.appendChild(root);
  return root;
}

describe("SpreadsheetApp formula-bar argument preview evaluation (structured references)", () => {
  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    document.body.innerHTML = "";

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("rewrites and evaluates structured table specifiers (#All/#Headers/#Totals/#Data)", () => {
    const root = createRoot();
    const formulaBarHost = document.createElement("div");
    document.body.appendChild(formulaBarHost);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar: formulaBarHost });

    const doc = app.getDocument();
    // Table range includes header row at A1:B1 and data rows at A2:B4.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "Amount");
    doc.setCellValue("Sheet1", { row: 0, col: 1 }, "Other");
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 1);
    doc.setCellValue("Sheet1", { row: 1, col: 1 }, 10);
    doc.setCellValue("Sheet1", { row: 2, col: 0 }, 2);
    doc.setCellValue("Sheet1", { row: 2, col: 1 }, 20);
    doc.setCellValue("Sheet1", { row: 3, col: 0 }, 3);
    doc.setCellValue("Sheet1", { row: 3, col: 1 }, 30);

    app.getSearchWorkbook().addTable({
      name: "Table1",
      sheetName: "Sheet1",
      startRow: 0,
      startCol: 0,
      endRow: 3,
      endCol: 1,
      columns: ["Amount", "Other"],
    });

    const evalPreview = (expr: string) => (app as any).evaluateFormulaBarArgumentPreview(expr);

    expect(evalPreview("SUM(Table1[#All])")).toBe(66);
    expect(evalPreview("SUM(Table1[#Headers])")).toBe(0);
    expect(evalPreview("SUM(Table1[#Data])")).toBe(66);
    expect(evalPreview("SUM(Table1[#Totals])")).toBe(33);

    app.destroy();
    root.remove();
    formulaBarHost.remove();
  });

  it("rewrites and evaluates structured table column references with selectors (Table1[[#Headers],[Col]])", () => {
    const root = createRoot();
    const formulaBarHost = document.createElement("div");
    document.body.appendChild(formulaBarHost);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar: formulaBarHost });

    const doc = app.getDocument();
    // Table range includes header row at A1:B1 and data rows at A2:B4.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "Amount");
    doc.setCellValue("Sheet1", { row: 0, col: 1 }, "Other");
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 1);
    doc.setCellValue("Sheet1", { row: 1, col: 1 }, 10);
    doc.setCellValue("Sheet1", { row: 2, col: 0 }, 2);
    doc.setCellValue("Sheet1", { row: 2, col: 1 }, 20);
    doc.setCellValue("Sheet1", { row: 3, col: 0 }, 3);
    doc.setCellValue("Sheet1", { row: 3, col: 1 }, 30);

    app.getSearchWorkbook().addTable({
      name: "Table1",
      sheetName: "Sheet1",
      startRow: 0,
      startCol: 0,
      endRow: 3,
      endCol: 1,
      columns: ["Amount", "Other"],
    });

    const evalPreview = (expr: string) => (app as any).evaluateFormulaBarArgumentPreview(expr);

    expect(evalPreview("SUM(Table1[Amount])")).toBe(6);
    expect(evalPreview("SUM(Table1[[#All],[Amount]])")).toBe(6);
    expect(evalPreview("SUM(Table1[[#Headers],[Amount]])")).toBe(0);
    expect(evalPreview("SUM(Table1[[#Totals],[Amount]])")).toBe(3);

    app.destroy();
    root.remove();
    formulaBarHost.remove();
  });

  it("supports escaped closing brackets inside column names (Table2[Amount]]])", () => {
    const root = createRoot();
    const formulaBarHost = document.createElement("div");
    document.body.appendChild(formulaBarHost);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar: formulaBarHost });

    const doc = app.getDocument();
    // Header row at A1, data rows at A2:A4.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "Amount]");
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 1);
    doc.setCellValue("Sheet1", { row: 2, col: 0 }, 2);
    doc.setCellValue("Sheet1", { row: 3, col: 0 }, 3);

    app.getSearchWorkbook().addTable({
      name: "Table2",
      sheetName: "Sheet1",
      startRow: 0,
      startCol: 0,
      endRow: 3,
      endCol: 0,
      columns: ["Amount]"],
    });

    const evalPreview = (expr: string) => (app as any).evaluateFormulaBarArgumentPreview(expr);

    // `]` inside column names is escaped via doubling: `Amount]` -> `Amount]]`.
    expect(evalPreview("SUM(Table2[Amount]]])")).toBe(6);

    app.destroy();
    root.remove();
    formulaBarHost.remove();
  });

  it("supports commas in structured reference column names when `]` is escaped (Table3[Total]],USD])", () => {
    const root = createRoot();
    const formulaBarHost = document.createElement("div");
    document.body.appendChild(formulaBarHost);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar: formulaBarHost });

    const doc = app.getDocument();
    // Header row at A1, data rows at A2:A4.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "Total],USD");
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 1);
    doc.setCellValue("Sheet1", { row: 2, col: 0 }, 2);
    doc.setCellValue("Sheet1", { row: 3, col: 0 }, 3);

    app.getSearchWorkbook().addTable({
      name: "Table3",
      sheetName: "Sheet1",
      startRow: 0,
      startCol: 0,
      endRow: 3,
      endCol: 0,
      columns: ["Total],USD"],
    });

    const evalPreview = (expr: string) => (app as any).evaluateFormulaBarArgumentPreview(expr);

    // `]` inside column names is escaped via doubling: `Total],USD` -> `Total]],USD`.
    expect(evalPreview("SUM(Table3[Total]],USD])")).toBe(6);

    app.destroy();
    root.remove();
    formulaBarHost.remove();
  });

  it("supports #This Row and @Column structured references when editing inside the table", () => {
    const root = createRoot();
    const formulaBarHost = document.createElement("div");
    document.body.appendChild(formulaBarHost);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar: formulaBarHost });

    const doc = app.getDocument();
    // Table range includes header row at A1 and data rows at A2:A4.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "Amount");
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 10);
    doc.setCellValue("Sheet1", { row: 2, col: 0 }, 20);
    doc.setCellValue("Sheet1", { row: 3, col: 0 }, 30);

    app.getSearchWorkbook().addTable({
      name: "TableThisRow",
      sheetName: "Sheet1",
      startRow: 0,
      startCol: 0,
      endRow: 3,
      endCol: 0,
      columns: ["Amount"],
    });

    // Pretend we're editing the formula in row 3 (0-based row 2), inside the table.
    (app as any).formulaEditCell = { sheetId: "Sheet1", cell: { row: 2, col: 0 } };

    const evalPreview = (expr: string) => (app as any).evaluateFormulaBarArgumentPreview(expr);

    expect(evalPreview("TableThisRow[[#This Row],[Amount]]")).toBe(20);
    expect(evalPreview("TableThisRow[@Amount]")).toBe(20);
    // Implicit "this row" shorthand used in calculated column formulas.
    expect(evalPreview("[@Amount]")).toBe(20);
    expect(evalPreview("SUM(TableThisRow[[#This Row],[Amount]], 5)")).toBe(25);
    expect(evalPreview("SUM([@Amount], 5)")).toBe(25);

    app.destroy();
    root.remove();
    formulaBarHost.remove();
  });
});
