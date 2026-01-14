/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";

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

describe("SpreadsheetApp formula-bar range preview tooltip", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
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

  it("renders a 2x2 preview grid for A1:B2", () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const doc = app.getDocument();
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, 1);
    doc.setCellValue("Sheet1", { row: 0, col: 1 }, 2);
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 3);
    doc.setCellValue("Sheet1", { row: 1, col: 1 }, 4);

    const bar = (app as any).formulaBar;
    bar.setActiveCell({ address: "C1", input: "=SUM(A1:B2)", value: null });

    const highlight = formulaBar.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpan = highlight?.querySelector<HTMLElement>('span[data-kind="reference"]');
    expect(refSpan?.textContent).toBe("A1:B2");
    refSpan?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    const tooltip = formulaBar.querySelector<HTMLElement>('[data-testid="formula-range-preview-tooltip"]');
    expect(tooltip).not.toBeNull();
    expect(tooltip?.hidden).toBe(false);

    const cells = Array.from(tooltip!.querySelectorAll("td")).map((td) => td.textContent);
    expect(cells).toEqual(["1", "2", "3", "4"]);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("renders a 3x3 sample + '(range too large: N cells)' summary for large ranges", () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const doc = app.getDocument();
    // Seed a 3x3 sample in the top-left corner of a large range.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, 1);
    doc.setCellValue("Sheet1", { row: 0, col: 1 }, 2);
    doc.setCellValue("Sheet1", { row: 0, col: 2 }, 3);
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 4);
    doc.setCellValue("Sheet1", { row: 1, col: 1 }, 5);
    doc.setCellValue("Sheet1", { row: 1, col: 2 }, 6);
    doc.setCellValue("Sheet1", { row: 2, col: 0 }, 7);
    doc.setCellValue("Sheet1", { row: 2, col: 1 }, 8);
    doc.setCellValue("Sheet1", { row: 2, col: 2 }, 9);

    // 11 rows x 10 cols = 110 cells (> 100 cap).
    const bar = (app as any).formulaBar;
    bar.setActiveCell({ address: "C1", input: "=SUM(A1:J11)", value: null });

    const highlight = formulaBar.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpan = highlight?.querySelector<HTMLElement>('span[data-kind="reference"]');
    expect(refSpan?.textContent).toBe("A1:J11");
    refSpan?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    const tooltip = formulaBar.querySelector<HTMLElement>('[data-testid="formula-range-preview-tooltip"]');
    expect(tooltip).not.toBeNull();
    expect(tooltip?.hidden).toBe(false);

    const summary = tooltip!.querySelector<HTMLElement>(".formula-range-preview-tooltip__summary");
    expect(summary?.textContent).toContain("range too large");
    expect(summary?.textContent).toContain("cells");
    expect(summary?.textContent).toMatch(/110/);

    const cells = Array.from(tooltip!.querySelectorAll("td")).map((td) => td.textContent);
    expect(cells).toEqual(["1", "2", "3", "4", "5", "6", "7", "8", "9"]);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("refreshes the tooltip sample when referenced cell values change", () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const doc = app.getDocument();
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, 1);

    const bar = (app as any).formulaBar;
    bar.setActiveCell({ address: "C1", input: "=A1", value: null });

    const highlight = formulaBar.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpan = highlight?.querySelector<HTMLElement>('span[data-kind="reference"]');
    expect(refSpan?.textContent).toBe("A1");
    refSpan?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    const tooltip = formulaBar.querySelector<HTMLElement>('[data-testid="formula-range-preview-tooltip"]');
    expect(tooltip).not.toBeNull();
    expect(tooltip?.hidden).toBe(false);

    let cells = Array.from(tooltip!.querySelectorAll("td")).map((td) => td.textContent);
    expect(cells).toEqual(["1"]);

    // Update referenced value; the tooltip should re-render on the next hover update.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, 99);
    refSpan?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    cells = Array.from(tooltip!.querySelectorAll("td")).map((td) => td.textContent);
    expect(cells).toEqual(["99"]);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("clears the tooltip when the active cell changes while hovering a reference in view mode", () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const doc = app.getDocument();
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, 1);
    doc.setCellValue("Sheet1", { row: 0, col: 1 }, 2);
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 3);
    doc.setCellValue("Sheet1", { row: 1, col: 1 }, 4);
    doc.setCellFormula("Sheet1", { row: 0, col: 2 }, "=SUM(A1:B2)");

    // Select the formula cell so view-mode highlight shows the reference span.
    app.activateCell({ row: 0, col: 2 });

    const highlight = formulaBar.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpan = highlight?.querySelector<HTMLElement>('span[data-kind="reference"]');
    expect(refSpan?.textContent).toBe("A1:B2");
    refSpan?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    const tooltip = formulaBar.querySelector<HTMLElement>('[data-testid="formula-range-preview-tooltip"]');
    expect(tooltip).not.toBeNull();
    expect(tooltip?.hidden).toBe(false);
    expect((app as any).referencePreview).not.toBeNull();

    // Changing the active cell can swap out the highlighted <pre> without firing `mouseleave`.
    // Ensure the tooltip/hover outline are explicitly cleared.
    app.activateCell({ row: 0, col: 3 });
    expect(tooltip?.hidden).toBe(true);
    expect((app as any).referencePreview).toBeNull();

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("skips sheet-qualified previews when the referenced sheet is not active", () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });

    // Ensure Sheet2 exists so the resolver can recognize it.
    const doc = app.getDocument();
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, 1);

    const bar = (app as any).formulaBar;
    bar.setActiveCell({ address: "C1", input: "=SUM(Sheet2!A1)", value: null });

    const highlight = formulaBar.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpan = highlight?.querySelector<HTMLElement>('span[data-kind="reference"]');
    expect(refSpan?.textContent).toBe("Sheet2!A1");
    refSpan?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    const tooltip = formulaBar.querySelector<HTMLElement>('[data-testid="formula-range-preview-tooltip"]');
    expect(tooltip).not.toBeNull();
    expect(tooltip?.hidden).toBe(true);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("hides and re-syncs the tooltip when switching sheets during formula editing", () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const doc = app.getDocument();
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, 123);

    const textarea = formulaBar.querySelector('[data-testid="formula-input"]') as HTMLTextAreaElement | null;
    expect(textarea).not.toBeNull();

    // Begin editing on Sheet1, then navigate to Sheet2 and enter a sheet-qualified reference.
    textarea!.dispatchEvent(new Event("focus"));
    app.activateSheet("Sheet2");

    textarea!.value = "=Sheet2!A1";
    textarea!.setSelectionRange(textarea!.value.length, textarea!.value.length);
    textarea!.dispatchEvent(new Event("input"));

    const tooltip = formulaBar.querySelector<HTMLElement>('[data-testid="formula-range-preview-tooltip"]');
    expect(tooltip).not.toBeNull();
    expect(tooltip?.hidden).toBe(false);

    // Switching away from the referenced sheet should immediately hide the tooltip (no new hover event).
    app.activateSheet("Sheet1");
    expect(tooltip?.hidden).toBe(true);

    // Switching back should re-sync from the current formula bar hover/caret state.
    app.activateSheet("Sheet2");
    expect(tooltip?.hidden).toBe(false);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("skips unqualified previews when editing a cell on another sheet", () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });

    // Ensure both sheets exist and have different values so a wrong-sheet preview would be obvious.
    const doc = app.getDocument();
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, 111);
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, 222);

    // Begin editing on Sheet1, then switch to Sheet2 without inserting a sheet-qualified reference.
    const textarea = formulaBar.querySelector('[data-testid="formula-input"]') as HTMLTextAreaElement | null;
    expect(textarea).not.toBeNull();
    textarea!.dispatchEvent(new Event("focus"));
    app.activateSheet("Sheet2");

    textarea!.value = "=A1";
    textarea!.setSelectionRange(textarea!.value.length, textarea!.value.length);
    textarea!.dispatchEvent(new Event("input"));

    const tooltip = formulaBar.querySelector<HTMLElement>('[data-testid="formula-range-preview-tooltip"]');
    expect(tooltip).not.toBeNull();
    expect(tooltip?.hidden).toBe(true);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("renders a preview grid for named ranges when hovering an identifier in view mode", () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const doc = app.getDocument();
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, 1);
    doc.setCellValue("Sheet1", { row: 0, col: 1 }, 2);
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 3);
    doc.setCellValue("Sheet1", { row: 1, col: 1 }, 4);
    app.getSearchWorkbook().defineName("SalesData", {
      sheetName: "Sheet1",
      range: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
    });

    const bar = (app as any).formulaBar;
    bar.setActiveCell({ address: "C1", input: "=SUM(SalesData)", value: null });

    const highlight = formulaBar.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const nameSpan = highlight?.querySelector<HTMLElement>('span[data-kind="identifier"]');
    expect(nameSpan?.textContent).toBe("SalesData");
    nameSpan?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    const tooltip = formulaBar.querySelector<HTMLElement>('[data-testid="formula-range-preview-tooltip"]');
    expect(tooltip).not.toBeNull();
    expect(tooltip?.hidden).toBe(false);

    const header = tooltip!.querySelector<HTMLElement>(".formula-range-preview-tooltip__header");
    expect(header?.textContent).toBe("SalesData (A1:B2)");

    const cells = Array.from(tooltip!.querySelectorAll("td")).map((td) => td.textContent);
    expect(cells).toEqual(["1", "2", "3", "4"]);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("skips named-range previews when the named range points at a non-active sheet", () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const doc = app.getDocument();
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, 123);
    app.getSearchWorkbook().defineName("OtherSheetName", {
      sheetName: "Sheet2",
      range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
    });

    const bar = (app as any).formulaBar;
    bar.setActiveCell({ address: "C1", input: "=SUM(OtherSheetName)", value: null });

    const highlight = formulaBar.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const nameSpan = highlight?.querySelector<HTMLElement>('span[data-kind="identifier"]');
    expect(nameSpan?.textContent).toBe("OtherSheetName");
    nameSpan?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    const tooltip = formulaBar.querySelector<HTMLElement>('[data-testid="formula-range-preview-tooltip"]');
    expect(tooltip).not.toBeNull();
    expect(tooltip?.hidden).toBe(true);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("renders a preview for structured table references (Table1[Amount])", () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const doc = app.getDocument();
    // Table range includes header row at A1 and data rows at A2:A4.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "Amount");
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 1);
    doc.setCellValue("Sheet1", { row: 2, col: 0 }, 2);
    doc.setCellValue("Sheet1", { row: 3, col: 0 }, 3);

    app.getSearchWorkbook().addTable({
      name: "Table1",
      sheetName: "Sheet1",
      startRow: 0,
      startCol: 0,
      endRow: 3,
      endCol: 0,
      columns: ["Amount"],
    });

    const bar = (app as any).formulaBar;
    bar.setActiveCell({ address: "C1", input: "=SUM(Table1[Amount])", value: null });

    const highlight = formulaBar.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpans = Array.from(highlight?.querySelectorAll<HTMLElement>('span[data-kind="reference"]') ?? []);
    const refSpan = refSpans.find((s) => s.textContent === "Table1[Amount]") ?? null;
    expect(refSpan?.textContent).toBe("Table1[Amount]");
    refSpan?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    const tooltip = formulaBar.querySelector<HTMLElement>('[data-testid="formula-range-preview-tooltip"]');
    expect(tooltip).not.toBeNull();
    expect(tooltip?.hidden).toBe(false);

    const header = tooltip!.querySelector<HTMLElement>(".formula-range-preview-tooltip__header");
    expect(header?.textContent).toBe("Table1[Amount] (A2:A4)");

    const cells = Array.from(tooltip!.querySelectorAll("td")).map((td) => td.textContent);
    expect(cells).toEqual(["1", "2", "3"]);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("renders a preview for structured table reference specifiers (Table1[#All])", () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });

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

    const bar = (app as any).formulaBar;
    bar.setActiveCell({ address: "C1", input: "=SUM(Table1[#All])", value: null });

    const highlight = formulaBar.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpans = Array.from(highlight?.querySelectorAll<HTMLElement>('span[data-kind="reference"]') ?? []);
    const refSpan = refSpans.find((s) => s.textContent === "Table1[#All]") ?? null;
    expect(refSpan?.textContent).toBe("Table1[#All]");
    refSpan?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    const tooltip = formulaBar.querySelector<HTMLElement>('[data-testid="formula-range-preview-tooltip"]');
    expect(tooltip).not.toBeNull();
    expect(tooltip?.hidden).toBe(false);

    const header = tooltip!.querySelector<HTMLElement>(".formula-range-preview-tooltip__header");
    expect(header?.textContent).toBe("Table1[#All] (A1:B4)");

    const cells = Array.from(tooltip!.querySelectorAll("td")).map((td) => td.textContent);
    // Sample grid is capped to 3 rows; includes header row for #All.
    expect(cells).toEqual(["Amount", "Other", "1", "10", "2", "20"]);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("renders a preview for structured table reference specifiers (Table1[#Headers])", () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const doc = app.getDocument();
    // Table range includes header row at A1:B1 and data rows at A2:B4.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "Amount");
    doc.setCellValue("Sheet1", { row: 0, col: 1 }, "Other");
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 1);
    doc.setCellValue("Sheet1", { row: 1, col: 1 }, 10);

    app.getSearchWorkbook().addTable({
      name: "Table1",
      sheetName: "Sheet1",
      startRow: 0,
      startCol: 0,
      endRow: 3,
      endCol: 1,
      columns: ["Amount", "Other"],
    });

    const bar = (app as any).formulaBar;
    bar.setActiveCell({ address: "C1", input: "=SUM(Table1[#Headers])", value: null });

    const highlight = formulaBar.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpans = Array.from(highlight?.querySelectorAll<HTMLElement>('span[data-kind="reference"]') ?? []);
    const refSpan = refSpans.find((s) => s.textContent === "Table1[#Headers]") ?? null;
    expect(refSpan?.textContent).toBe("Table1[#Headers]");
    refSpan?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    const tooltip = formulaBar.querySelector<HTMLElement>('[data-testid="formula-range-preview-tooltip"]');
    expect(tooltip).not.toBeNull();
    expect(tooltip?.hidden).toBe(false);

    const header = tooltip!.querySelector<HTMLElement>(".formula-range-preview-tooltip__header");
    expect(header?.textContent).toBe("Table1[#Headers] (A1:B1)");

    const cells = Array.from(tooltip!.querySelectorAll("td")).map((td) => td.textContent);
    expect(cells).toEqual(["Amount", "Other"]);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("renders a preview for structured table reference specifiers (Table1[#Data])", () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });

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

    const bar = (app as any).formulaBar;
    bar.setActiveCell({ address: "C1", input: "=SUM(Table1[#Data])", value: null });

    const highlight = formulaBar.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpans = Array.from(highlight?.querySelectorAll<HTMLElement>('span[data-kind="reference"]') ?? []);
    const refSpan = refSpans.find((s) => s.textContent === "Table1[#Data]") ?? null;
    expect(refSpan?.textContent).toBe("Table1[#Data]");
    refSpan?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    const tooltip = formulaBar.querySelector<HTMLElement>('[data-testid="formula-range-preview-tooltip"]');
    expect(tooltip).not.toBeNull();
    expect(tooltip?.hidden).toBe(false);

    const header = tooltip!.querySelector<HTMLElement>(".formula-range-preview-tooltip__header");
    expect(header?.textContent).toBe("Table1[#Data] (A2:B4)");

    const cells = Array.from(tooltip!.querySelectorAll("td")).map((td) => td.textContent);
    // Sample grid is capped to 3 rows; #Data excludes the header row.
    expect(cells).toEqual(["1", "10", "2", "20", "3", "30"]);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("renders a preview for structured table reference specifiers (Table1[#Totals])", () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const doc = app.getDocument();
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "Amount");
    doc.setCellValue("Sheet1", { row: 0, col: 1 }, "Other");
    // Treat the last table row as the totals row.
    doc.setCellValue("Sheet1", { row: 3, col: 0 }, 123);
    doc.setCellValue("Sheet1", { row: 3, col: 1 }, 456);

    app.getSearchWorkbook().addTable({
      name: "Table1",
      sheetName: "Sheet1",
      startRow: 0,
      startCol: 0,
      endRow: 3,
      endCol: 1,
      columns: ["Amount", "Other"],
    });

    const bar = (app as any).formulaBar;
    bar.setActiveCell({ address: "C1", input: "=SUM(Table1[#Totals])", value: null });

    const highlight = formulaBar.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpans = Array.from(highlight?.querySelectorAll<HTMLElement>('span[data-kind="reference"]') ?? []);
    const refSpan = refSpans.find((s) => s.textContent === "Table1[#Totals]") ?? null;
    expect(refSpan?.textContent).toBe("Table1[#Totals]");
    refSpan?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    const tooltip = formulaBar.querySelector<HTMLElement>('[data-testid="formula-range-preview-tooltip"]');
    expect(tooltip).not.toBeNull();
    expect(tooltip?.hidden).toBe(false);

    const header = tooltip!.querySelector<HTMLElement>(".formula-range-preview-tooltip__header");
    expect(header?.textContent).toBe("Table1[#Totals] (A4:B4)");

    const cells = Array.from(tooltip!.querySelectorAll("td")).map((td) => td.textContent);
    expect(cells).toEqual(["123", "456"]);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("renders a preview for structured table references with #All (Table1[[#All],[Amount]])", () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const doc = app.getDocument();
    // Table range includes header row at A1 and data rows at A2:A4.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "Amount");
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 1);
    doc.setCellValue("Sheet1", { row: 2, col: 0 }, 2);
    doc.setCellValue("Sheet1", { row: 3, col: 0 }, 3);

    app.getSearchWorkbook().addTable({
      name: "Table1",
      sheetName: "Sheet1",
      startRow: 0,
      startCol: 0,
      endRow: 3,
      endCol: 0,
      columns: ["Amount"],
    });

    const bar = (app as any).formulaBar;
    bar.setActiveCell({ address: "C1", input: "=SUM(Table1[[#All],[Amount]])", value: null });

    const highlight = formulaBar.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpans = Array.from(highlight?.querySelectorAll<HTMLElement>('span[data-kind="reference"]') ?? []);
    const refSpan = refSpans.find((s) => s.textContent === "Table1[[#All],[Amount]]") ?? null;
    expect(refSpan?.textContent).toBe("Table1[[#All],[Amount]]");
    refSpan?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    const tooltip = formulaBar.querySelector<HTMLElement>('[data-testid="formula-range-preview-tooltip"]');
    expect(tooltip).not.toBeNull();
    expect(tooltip?.hidden).toBe(false);

    const header = tooltip!.querySelector<HTMLElement>(".formula-range-preview-tooltip__header");
    expect(header?.textContent).toBe("Table1[[#All],[Amount]] (A1:A4)");

    // Sample grid is capped to 3 rows; #All includes the header row.
    const cells = Array.from(tooltip!.querySelectorAll("td")).map((td) => td.textContent);
    expect(cells).toEqual(["Amount", "1", "2"]);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("skips structured table previews when the table belongs to a non-active sheet", () => {
    const root = createRoot();
    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { formulaBar });

    const doc = app.getDocument();
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "Amount");
    doc.setCellValue("Sheet2", { row: 1, col: 0 }, 99);
    doc.setCellValue("Sheet2", { row: 2, col: 0 }, 100);

    app.getSearchWorkbook().addTable({
      name: "OtherTable",
      sheetName: "Sheet2",
      startRow: 0,
      startCol: 0,
      endRow: 2,
      endCol: 0,
      columns: ["Amount"],
    });

    const bar = (app as any).formulaBar;
    bar.setActiveCell({ address: "C1", input: "=SUM(OtherTable[Amount])", value: null });

    const highlight = formulaBar.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpans = Array.from(highlight?.querySelectorAll<HTMLElement>('span[data-kind="reference"]') ?? []);
    const refSpan = refSpans.find((s) => s.textContent === "OtherTable[Amount]") ?? null;
    expect(refSpan?.textContent).toBe("OtherTable[Amount]");
    refSpan?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    const tooltip = formulaBar.querySelector<HTMLElement>('[data-testid="formula-range-preview-tooltip"]');
    expect(tooltip).not.toBeNull();
    expect(tooltip?.hidden).toBe(true);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });
});
