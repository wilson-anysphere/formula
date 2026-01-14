import { afterEach, describe, expect, it, vi } from "vitest";
import { ToolExecutor } from "../src/executor/tool-executor.js";
import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";
import { parseA1Cell, parseA1Range } from "../src/spreadsheet/a1.js";
import type { SpreadsheetApi } from "../src/spreadsheet/api.js";

describe("ToolExecutor", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("write_cell writes a scalar value", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const result = await executor.execute({
      name: "write_cell",
      parameters: { cell: "Sheet1!A1", value: 42 }
    });

    expect(result.ok).toBe(true);
    expect(workbook.getCell(parseA1Cell("Sheet1!A1")).value).toBe(42);
  });

  it("resolves display sheet names to stable sheet ids when sheet_name_resolver is provided", async () => {
    const workbook = new InMemoryWorkbook(["Sheet2"]);
    const sheetNameResolver = {
      getSheetIdByName(name: string) {
        return name.toLowerCase() === "budget" ? "Sheet2" : null;
      },
      getSheetNameById(id: string) {
        return id === "Sheet2" ? "Budget" : null;
      }
    };

    const executor = new ToolExecutor(workbook, { default_sheet: "Sheet2", sheet_name_resolver: sheetNameResolver });
    const result = await executor.execute({
      name: "write_cell",
      parameters: { cell: "Budget!A1", value: 99 }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("write_cell");
    if (!result.ok || result.tool !== "write_cell") throw new Error("Unexpected tool result");
    expect(result.data?.cell).toBe("Budget!A1");
    expect(workbook.getCell(parseA1Cell("Sheet2!A1")).value).toBe(99);
    expect(workbook.listSheets()).toEqual(["Sheet2"]);
  });

  it("canonicalizes default_sheet when it is provided as a display sheet name", async () => {
    const workbook = new InMemoryWorkbook(["Sheet2"]);
    const sheetNameResolver = {
      getSheetIdByName(name: string) {
        return name.toLowerCase() === "budget" ? "Sheet2" : null;
      },
      getSheetNameById(id: string) {
        return id === "Sheet2" ? "Budget" : null;
      }
    };

    const executor = new ToolExecutor(workbook, { default_sheet: "Budget", sheet_name_resolver: sheetNameResolver });
    const result = await executor.execute({
      name: "write_cell",
      parameters: { cell: "A1", value: 123 }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("write_cell");
    if (!result.ok || result.tool !== "write_cell") throw new Error("Unexpected tool result");
    // Returned cell ref is user-facing (display name).
    expect(result.data?.cell).toBe("Budget!A1");

    // Underlying workbook still uses stable ids and does not create a "Budget" sheet.
    expect(workbook.getCell(parseA1Cell("Sheet2!A1")).value).toBe(123);
    expect(workbook.listSheets()).toEqual(["Sheet2"]);
  });

  it("refreshes pivots when subsequent edits use a display sheet name", async () => {
    const workbook = new InMemoryWorkbook(["Sheet2"]);
    const sheetNameResolver = {
      getSheetIdByName(name: string) {
        return name.toLowerCase() === "budget" ? "Sheet2" : null;
      },
      getSheetNameById(id: string) {
        return id === "Sheet2" ? "Budget" : null;
      }
    };

    const executor = new ToolExecutor(workbook, { default_sheet: "Sheet2", sheet_name_resolver: sheetNameResolver });

    // Source data.
    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Budget!A1:B3",
        values: [
          ["Category", "Value"],
          ["A", 10],
          ["B", 20]
        ]
      }
    });

    // Create a simple pivot at D1.
    const pivot = await executor.execute({
      name: "create_pivot_table",
      parameters: {
        source_range: "Budget!A1:B3",
        destination: "Budget!D1",
        rows: ["Category"],
        columns: [],
        values: [{ field: "Value", aggregation: "sum" }]
      }
    });

    expect(pivot.ok).toBe(true);
    expect(pivot.tool).toBe("create_pivot_table");
    if (!pivot.ok || pivot.tool !== "create_pivot_table") throw new Error("Unexpected tool result");
    expect(pivot.data?.destination_range).toBe("Budget!D1:E4");

    // Validate initial pivot output written to the stable-id sheet.
    expect(workbook.getCell(parseA1Cell("Sheet2!D2")).value).toBe("A");
    expect(workbook.getCell(parseA1Cell("Sheet2!E2")).value).toBe(10);
    expect(workbook.getCell(parseA1Cell("Sheet2!E4")).value).toBe(30);

    // Edit the source using the display sheet name and ensure pivot refreshes.
    await executor.execute({
      name: "write_cell",
      parameters: { cell: "Budget!B2", value: 15 }
    });

    expect(workbook.getCell(parseA1Cell("Sheet2!E2")).value).toBe(15);
    expect(workbook.getCell(parseA1Cell("Sheet2!E4")).value).toBe(35);
    // Ensure we did not create a phantom "Budget" sheet.
    expect(workbook.listSheets()).toEqual(["Sheet2"]);
  });

  it("create_chart passes stable sheet ids to the host but returns display names to the caller", async () => {
    const workbook = new InMemoryWorkbook(["Sheet2"]);
    const sheetNameResolver = {
      getSheetIdByName(name: string) {
        return name.toLowerCase() === "budget" ? "Sheet2" : null;
      },
      getSheetNameById(id: string) {
        return id === "Sheet2" ? "Budget" : null;
      }
    };

    const executor = new ToolExecutor(workbook, { default_sheet: "Sheet2", sheet_name_resolver: sheetNameResolver });
    const result = await executor.execute({
      name: "create_chart",
      parameters: {
        chart_type: "bar",
        data_range: "Budget!A1:B3",
        title: "Demo"
      }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("create_chart");
    if (!result.ok || result.tool !== "create_chart") throw new Error("Unexpected tool result");

    // User-facing output uses display sheet name.
    expect(result.data?.data_range).toBe("Budget!A1:B3");

    // Host receives stable sheet id.
    const charts = workbook.listCharts();
    expect(charts).toHaveLength(1);
    expect(charts[0]?.spec.data_range).toBe("Sheet2!A1:B3");
  });

  it("supports quoted display sheet names (spaces/quotes) with sheet_name_resolver", async () => {
    const workbook = new InMemoryWorkbook(["Sheet2"]);
    const sheetNameResolver = {
      getSheetIdByName(name: string) {
        const normalized = name.trim().toLowerCase();
        if (normalized === "q1 budget") return "Sheet2";
        if (normalized === "bob's budget") return "Sheet2";
        return null;
      },
      getSheetNameById(id: string) {
        return id === "Sheet2" ? "Bob's Budget" : null;
      }
    };

    const executor = new ToolExecutor(workbook, { default_sheet: "Sheet2", sheet_name_resolver: sheetNameResolver });

    const result = await executor.execute({
      name: "write_cell",
      parameters: { cell: "'Bob''s Budget'!A1", value: 7 }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("write_cell");
    if (!result.ok || result.tool !== "write_cell") throw new Error("Unexpected tool result");
    expect(result.data?.cell).toBe("'Bob''s Budget'!A1");

    expect(workbook.getCell(parseA1Cell("Sheet2!A1")).value).toBe(7);
    expect(workbook.listSheets()).toEqual(["Sheet2"]);
  });

  it("write_cell writes a formula when value starts with '='", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "write_cell",
      parameters: { cell: "Sheet1!B2", value: "=SUM(A1:A10)" }
    });

    const cell = workbook.getCell(parseA1Cell("Sheet1!B2"));
    expect(cell.value).toBeNull();
    expect(cell.formula).toBe("=SUM(A1:A10)");
  });

  it("write_cell normalizes formula whitespace (canonical display semantics)", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "write_cell",
      parameters: { cell: "Sheet1!B2", value: "  =  SUM(A1:A10)  " }
    });

    const cell = workbook.getCell(parseA1Cell("Sheet1!B2"));
    expect(cell.value).toBeNull();
    expect(cell.formula).toBe("=SUM(A1:A10)");
  });

  it("write_cell treats bare '=' as an empty formula (clears the cell)", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "write_cell",
      parameters: { cell: "Sheet1!A1", value: 5 }
    });

    const result = await executor.execute({
      name: "write_cell",
      parameters: { cell: "Sheet1!A1", value: "=" }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("write_cell");
    if (!result.ok || result.tool !== "write_cell") throw new Error("Unexpected tool result");
    expect(result.data?.changed).toBe(true);

    const cell = workbook.getCell(parseA1Cell("Sheet1!A1"));
    expect(cell.value).toBeNull();
    expect(cell.formula).toBeUndefined();
  });

  it("set_range normalizes formula inputs and clears empty formulas", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:B1",
        values: [["  = 1+1  ", "="]]
      }
    });

    const a1 = workbook.getCell(parseA1Cell("Sheet1!A1"));
    expect(a1.value).toBeNull();
    expect(a1.formula).toBe("=1+1");

    const b1 = workbook.getCell(parseA1Cell("Sheet1!B1"));
    expect(b1.value).toBeNull();
    expect(b1.formula).toBeUndefined();
  });

  it("set_range updates a rectangular range", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const result = await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:B2",
        values: [
          [1, 2],
          [3, 4]
        ]
      }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("set_range");
    if (!result.ok || result.tool !== "set_range") throw new Error("Unexpected tool result");
    expect(result.data?.updated_cells).toBe(4);
    const range = parseA1Range("Sheet1!A1:B2");
    const values = workbook.readRange(range).map((row) => row.map((cell) => cell.value));
    expect(values).toEqual([
      [1, 2],
      [3, 4]
    ]);
  });

  it("set_range expands from a start cell when given a single-cell range", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const result = await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!C3",
        values: [
          [1, 2, 3],
          [4, 5, 6]
        ]
      }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("set_range");
    if (!result.ok || result.tool !== "set_range") throw new Error("Unexpected tool result");
    expect(result.data?.updated_cells).toBe(6);
    expect(result.data?.range).toBe("Sheet1!C3:E4");

    const values = workbook
      .readRange(parseA1Range("Sheet1!C3:E4"))
      .map((row) => row.map((cell) => cell.value));
    expect(values).toEqual([
      [1, 2, 3],
      [4, 5, 6]
    ]);
  });

  it("apply_formula_column fills formulas down to the last used row when end_row = -1", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:A3",
        values: [["Header"], [10], [20]]
      }
    });

    const result = await executor.execute({
      name: "apply_formula_column",
      parameters: { column: "C", formula_template: "=A{row}*2", start_row: 2, end_row: -1 }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("apply_formula_column");
    if (!result.ok || result.tool !== "apply_formula_column") throw new Error("Unexpected tool result");
    expect(result.data?.updated_cells).toBe(2);
    expect(workbook.getCell(parseA1Cell("Sheet1!C2")).formula).toBe("=A2*2");
    expect(workbook.getCell(parseA1Cell("Sheet1!C3")).formula).toBe("=A3*2");
  });

  it("accepts camelCase parameter aliases from docs examples", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "write_cell",
      parameters: { cell: "Sheet1!A1", value: 5 }
    });

    const result = await executor.execute({
      name: "apply_formula_column",
      parameters: { column: "B", formulaTemplate: "=A{row}*10", startRow: 1, endRow: 1 }
    });

    expect(result.ok).toBe(true);
    expect(workbook.getCell(parseA1Cell("Sheet1!B1")).formula).toBe("=A1*10");
  });

  it("apply_formatting accepts camelCase format field aliases from docs examples", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { default_sheet: "Sheet1" });

    const result = await executor.execute({
      name: "apply_formatting",
      parameters: {
        range: "A1",
        format: {
          bold: true,
          italic: true,
          fontSize: 16,
          fontColor: "#FF00FF00",
          backgroundColor: "#FFFFFF00",
          numberFormat: "$#,##0.00",
          horizontalAlign: "center",
        },
      },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("apply_formatting");
    if (!result.ok || result.tool !== "apply_formatting") throw new Error("Unexpected tool result");
    expect(result.data?.range).toBe("Sheet1!A1");
    expect(result.data?.formatted_cells).toBe(1);

    expect(workbook.getCell(parseA1Cell("Sheet1!A1")).format).toEqual({
      bold: true,
      italic: true,
      font_size: 16,
      font_color: "#FF00FF00",
      background_color: "#FFFFFF00",
      number_format: "$#,##0.00",
      horizontal_align: "center",
    });
  });

  it("returns validation_error for invalid A1 references", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const result = await executor.execute({
      name: "write_cell",
      parameters: { cell: "NotACell", value: 1 }
    });

    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("validation_error");
  });

  it("accepts $-absolute A1 references (e.g. $A$1) in ranges", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:B2",
        values: [
          [1, 2],
          [3, 4],
        ],
      },
    });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "$A$1:$B$2" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");
    expect(result.data?.range).toBe("Sheet1!A1:B2");
    expect(result.data?.values).toEqual([
      [1, 2],
      [3, 4],
    ]);
  });

  it("read_range enforces max_read_range_cells to prevent huge matrices", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    // 100x100 = 10,000 cells (default limit is 5,000).
    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:CV100" },
    });

    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("permission_denied");
    expect(result.error?.message).toMatch(/max_read_range_cells/i);
    expect(result.error?.message).toMatch(/10000/);
  });

  it("read_range normalizes object cell values to JSON-safe scalars", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: { foo: "bar" } as any });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A1" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    expect(result.data?.values).toEqual([['{"foo":"bar"}']]);
    expect(typeof result.data?.values?.[0]?.[0]).toBe("string");
    expect(() => JSON.stringify(result)).not.toThrow();
  });

  it("read_range formats rich text objects as plain text", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    workbook.setCell(parseA1Cell("Sheet1!A1"), {
      value: {
        text: "Hello world",
        runs: [{ text: "Hello" }, { text: " world" }],
      } as any,
    });
    // Some backends attach extra metadata fields; still prefer the plain text.
    workbook.setCell(parseA1Cell("Sheet1!A2"), {
      value: {
        text: "Hello with metadata",
        runs: [],
        meta: { foo: "bar" },
      } as any,
    });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A2" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    expect(result.data?.values).toEqual([["Hello world"], ["Hello with metadata"]]);
    expect(() => JSON.stringify(result)).not.toThrow();
  });

  it("read_range formats in-cell image values as altText / [Image]", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    // Envelope shape: `{ type: "image", value: { imageId, altText? } }`
    workbook.setCell(parseA1Cell("Sheet1!A1"), {
      value: { type: "image", value: { imageId: "img_1", altText: "Product photo" } } as any,
    });
    workbook.setCell(parseA1Cell("Sheet1!A2"), {
      value: { type: "image", value: { imageId: "img_2" } } as any,
    });

    // Direct payload shape: `{ imageId, altText?, width?, height? }`
    workbook.setCell(parseA1Cell("Sheet1!A3"), {
      value: { imageId: "img_3", altText: "Logo", width: 120, height: 60 } as any,
    });
    workbook.setCell(parseA1Cell("Sheet1!A4"), {
      value: { imageId: "img_4", altText: "   " } as any,
    });
    workbook.setCell(parseA1Cell("Sheet1!A5"), {
      value: { type: "image", value: { image_id: "img_5", alt_text: "Alt (snake_case)" } } as any,
    });
    workbook.setCell(parseA1Cell("Sheet1!A6"), {
      value: { type: "image", value: { id: "img_6", altText: "Alt (id)" } } as any,
    });
    workbook.setCell(parseA1Cell("Sheet1!A7"), {
      // Some adapters may use `id` for the direct payload shape.
      value: { id: "img_7", altText: "Alt (direct id)" } as any,
    });
    workbook.setCell(parseA1Cell("Sheet1!A8"), {
      // Ensure we do not misclassify generic objects with `id` as images.
      value: { id: "not-image", foo: "bar" } as any,
    });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A8" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    expect(result.data?.values).toEqual([
      ["Product photo"],
      ["[Image]"],
      ["Logo"],
      ["[Image]"],
      ["Alt (snake_case)"],
      ["Alt (id)"],
      ["Alt (direct id)"],
      ['{"id":"not-image","foo":"bar"}'],
    ]);
    expect(() => JSON.stringify(result)).not.toThrow();
  });

  it("read_range summarizes ArrayBuffer views (typed arrays) to avoid huge JSON payloads", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const bytes = new Uint8Array([1, 2, 3]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: bytes as any });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A1" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    expect(result.data?.values).toEqual([
      [JSON.stringify({ __type: "Uint8Array", length: 3, byteLength: 3, sample: [1, 2, 3] })],
    ]);
    expect(() => JSON.stringify(result)).not.toThrow();
  });

  it("read_range summarizes huge arrays to avoid huge JSON payloads", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const arr = Array.from({ length: 1000 }, (_, idx) => idx);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: arr as any });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A1" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    expect(result.data?.values).toEqual([
      [JSON.stringify({ __type: "Array", length: 1000, sample: arr.slice(0, 32), truncated: true })],
    ]);
    expect(() => JSON.stringify(result)).not.toThrow();
  });

  it("read_range summarizes huge Maps to avoid huge JSON payloads", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const map = new Map<number, number>();
    for (let i = 0; i < 1000; i++) map.set(i, i);

    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: map as any });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A1" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    expect(result.data?.values).toEqual([
      [
        JSON.stringify({
          __type: "Map",
          size: 1000,
          sample: Array.from(map.entries()).slice(0, 32),
          truncated: true,
        }),
      ],
    ]);
    expect(() => JSON.stringify(result)).not.toThrow();
  });

  it("read_range summarizes huge objects to avoid huge JSON payloads", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const obj: Record<string, number> = {};
    for (let i = 0; i < 1000; i++) obj[`k${i}`] = i;

    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: obj as any });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A1" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    const keys = Object.keys(obj).slice(0, 32);
    const sample: Record<string, number> = {};
    for (const key of keys) sample[key] = obj[key]!;
    expect(result.data?.values).toEqual([[JSON.stringify({ __type: "Object", keys, sample, truncated: true })]]);
    expect(() => JSON.stringify(result)).not.toThrow();
  });

  it("read_range handles circular rich values without crashing", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const arr: any[] = Array.from({ length: 300 }, (_, idx) => idx);
    arr[0] = arr;
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: arr as any });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A1" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    const value = result.data?.values?.[0]?.[0];
    expect(typeof value).toBe("string");
    expect(value).toContain("[Circular]");
    expect(() => JSON.stringify(result)).not.toThrow();
  });

  it("read_range tolerates missing/invalid CellData entries from SpreadsheetApi.readRange", async () => {
    const spreadsheet: any = {
      listSheets: () => ["Sheet1"],
      listNonEmptyCells: () => [],
      getCell: () => ({ value: null }),
      setCell: () => {},
      readRange: () => [[undefined]],
      writeRange: () => {},
      applyFormatting: () => 0,
      getLastUsedRow: () => 0,
      clone() {
        return this;
      },
    };
    const executor = new ToolExecutor(spreadsheet);

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A1" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");
    expect(result.data?.values).toEqual([[null]]);
    expect(() => JSON.stringify(result)).not.toThrow();
  });

  it("read_range truncates huge string cell payloads (per-cell) and stays JSON-serializable", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const raw = "x".repeat(210_000);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: raw });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A1" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    const value = result.data?.values?.[0]?.[0];
    expect(typeof value).toBe("string");
    expect(value).toBe(`${raw.slice(0, 10_000)}…[truncated 200000 chars]`);
    expect(() => JSON.stringify(result)).not.toThrow();
  });

  it("read_range truncates huge object cell payloads (stringified) and stays JSON-serializable", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const obj = { foo: "y".repeat(50_000) };
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: obj as any });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A1" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    const value = result.data?.values?.[0]?.[0];
    expect(typeof value).toBe("string");
    if (typeof value !== "string") {
      throw new Error("Expected read_range cell value to be a string");
    }
    // Exact serialization may vary depending on internal rich-value bounding heuristics, but it must be bounded.
    if (typeof value !== "string") throw new Error("Expected string cell payload");
    expect(value.length).toBeLessThanOrEqual(10_100);
    expect(value).toContain("truncated");
    expect(() => JSON.stringify(result)).not.toThrow();
  });

  it("read_range never throws when a cell value cannot be stringified or coerced", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const bad = {
      toJSON() {
        throw new Error("nope");
      },
      toString() {
        throw new Error("also nope");
      },
    };
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: bad as any });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A1" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    expect(result.data?.values).toEqual([["[Unserializable cell value]"]]);
    expect(() => JSON.stringify(result)).not.toThrow();
  });

  it("read_range normalizes non-string formulas to JSON-safe scalars", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: null, formula: { op: "SUM", args: ["A1:A3"] } as any });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A1", include_formulas: true },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    expect(result.data?.values).toEqual([[null]]);
    expect(result.data?.formulas).toEqual([[JSON.stringify({ op: "SUM", args: ["A1:A3"] })]]);
    expect(() => JSON.stringify(result)).not.toThrow();
  });

  it("read_range enforces max_read_range_chars based on normalized output size", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    // 21 * (10k chars) ~= 210k chars, exceeding default max_read_range_chars (200k).
    const long = "x".repeat(10_000);
    const set = await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:U1",
        values: [Array.from({ length: 21 }, () => long)],
      },
    });
    expect(set.ok).toBe(true);

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:U1" },
    });

    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("permission_denied");
    expect(result.error?.message).toMatch(/max_read_range_chars/i);
  });

  it("read_range enforces max_read_range_chars using JSON-escaped string length (quotes/backslashes)", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    // Keep the limit small so we can exercise escaping overhead.
    const executor = new ToolExecutor(workbook, { max_read_range_chars: 30 });

    // Raw length = 20, but JSON-escaped length = 42 (each quote/backslash becomes 2 chars, plus outer quotes).
    // This used to bypass the limit when we estimated `value.length + 2`.
    const escapey = '"\\'.repeat(10);
    const set = await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:A1",
        values: [[escapey]],
      },
    });
    expect(set.ok).toBe(true);

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:A1" },
    });

    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("permission_denied");
    expect(result.error?.message).toMatch(/max_read_range_chars/i);
  });

  it("apply_formatting returns runtime_error when SpreadsheetApi.applyFormatting throws", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]) as any;
    workbook.applyFormatting = () => {
      throw new Error("Formatting could not be applied to Sheet1!A1. Try selecting fewer cells/rows.");
    };
    const executor = new ToolExecutor(workbook as SpreadsheetApi, { default_sheet: "Sheet1" });

    const result = await executor.execute({
      name: "apply_formatting",
      parameters: { range: "A1", format: { bold: true } },
    });

    expect(result.ok).toBe(false);
    expect(result.tool).toBe("apply_formatting");
    expect(result.error?.code).toBe("runtime_error");
    expect(result.error?.message ?? "").toMatch(/Formatting could not be applied/i);
  });

  it("filter_range caps matching_rows list and preserves total match count", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    // 1,201 matches > default cap (1,000)
    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:A1201",
        values: Array.from({ length: 1201 }, () => [1]),
      },
    });

    const result = await executor.execute({
      name: "filter_range",
      parameters: {
        range: "Sheet1!A1:A1201",
        criteria: [{ column: "A", operator: "equals", value: 1 }],
      },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("filter_range");
    if (!result.ok || result.tool !== "filter_range") throw new Error("Unexpected tool result");

    expect(result.data?.count).toBe(1201);
    expect(result.data?.matching_rows).toHaveLength(1000);
    expect(result.data?.truncated).toBe(true);
  });

  it("detect_anomalies caps large anomaly lists and reports total_anomalies", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:A1500",
        values: Array.from({ length: 1500 }, (_, idx) => [idx + 1]),
      },
    });

    // Use a tiny positive threshold (schema requires > 0) so every cell is treated as an anomaly.
    const result = await executor.execute({
      name: "detect_anomalies",
      parameters: { range: "Sheet1!A1:A1500", method: "zscore", threshold: 0.0001 },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("detect_anomalies");
    if (!result.ok || result.tool !== "detect_anomalies") throw new Error("Unexpected tool result");
    if (!result.data || result.data.method !== "zscore") throw new Error("Unexpected anomaly result");

    expect(result.data.anomalies).toHaveLength(1000);
    expect(result.data.truncated).toBe(true);
    expect(result.data.total_anomalies).toBe(1500);
  });

  it("quotes sheet names with spaces when formatting results", async () => {
    const workbook = new InMemoryWorkbook(["My Sheet"]);
    const executor = new ToolExecutor(workbook, { default_sheet: "My Sheet" });

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "A1:B1",
        values: [[1, 2]],
      },
    });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "A1:B1" },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");
    expect(result.data?.range).toBe("'My Sheet'!A1:B1");
  });

  it("create_pivot_table writes a pivot output table", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:C5",
        values: [
          ["Region", "Product", "Sales"],
          ["East", "A", 100],
          ["East", "B", 150],
          ["West", "A", 200],
          ["West", "B", 250]
        ]
      }
    });

    const result = await executor.execute({
      name: "create_pivot_table",
      parameters: {
        source_range: "Sheet1!A1:C5",
        rows: ["Region"],
        columns: ["Product"],
        values: [{ field: "Sales", aggregation: "sum" }],
        destination: "Sheet1!E1"
      }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("create_pivot_table");

    const out = workbook
      .readRange(parseA1Range("Sheet1!E1:H4"))
      .map((row) => row.map((cell) => cell.value));

    expect(out).toEqual([
      ["Region", "A - Sum of Sales", "B - Sum of Sales", "Grand Total - Sum of Sales"],
      ["East", 100, 150, 250],
      ["West", 200, 250, 450],
      ["Grand Total", 300, 400, 700]
    ]);

    // Updating the source range should refresh the pivot output automatically.
    await executor.execute({
      name: "write_cell",
      parameters: { cell: "Sheet1!C2", value: 110 }
    });

    const refreshed = workbook
      .readRange(parseA1Range("Sheet1!E1:H4"))
      .map((row) => row.map((cell) => cell.value));

    expect(refreshed).toEqual([
      ["Region", "A - Sum of Sales", "B - Sum of Sales", "Grand Total - Sum of Sales"],
      ["East", 110, 150, 260],
      ["West", 200, 250, 450],
      ["Grand Total", 310, 400, 710]
    ]);
  });

  it("create_pivot_table supports variance/stddev aggregations", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    await executor.execute({
      name: "set_range",
      parameters: {
        range: "Sheet1!A1:C5",
        values: [
          ["Region", "Product", "Sales"],
          ["East", "A", 100],
          ["East", "B", 150],
          ["West", "A", 200],
          ["West", "B", 250]
        ]
      }
    });

    const result = await executor.execute({
      name: "create_pivot_table",
      parameters: {
        source_range: "Sheet1!A1:C5",
        rows: ["Region"],
        values: [
          { field: "Sales", aggregation: "varp" },
          { field: "Sales", aggregation: "stddevp" }
        ],
        destination: "Sheet1!E1"
      }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("create_pivot_table");

    const out = workbook
      .readRange(parseA1Range("Sheet1!E1:G4"))
      .map((row) => row.map((cell) => cell.value));

    expect(out[0]).toEqual(["Region", "VarP of Sales", "StdDevP of Sales"]);
    expect(out[1]).toEqual(["East", 625, 25]);
    expect(out[2]).toEqual(["West", 625, 25]);

    // Grand total is based on all records; check it roughly matches expected values.
    expect(out[3]?.[0]).toBe("Grand Total");
    expect(out[3]?.[1]).toBeCloseTo(3125, 10);
    expect(out[3]?.[2]).toBeCloseTo(Math.sqrt(3125), 10);
  });

  it("create_chart delegates to SpreadsheetApi.createChart", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { default_sheet: "Sheet1" });

    const result = await executor.execute({
      name: "create_chart",
      parameters: {
        chart_type: "bar",
        data_range: "A1:B3",
        title: "Sales"
      }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("create_chart");
    if (!result.ok || result.tool !== "create_chart") throw new Error("Unexpected tool result");
    expect(result.data?.status).toBe("ok");
    expect(result.data?.chart_id).toBe("chart_1");
    expect(workbook.listCharts()).toHaveLength(1);
    expect(workbook.listCharts()[0]?.spec).toMatchObject({
      chart_type: "bar",
      data_range: "Sheet1!A1:B3",
      title: "Sales"
    });
  });

  it("create_chart trims chart_id/title returned by the host (canonical identifiers)", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]) as any;
    const originalCreateChart = workbook.createChart.bind(workbook);
    workbook.createChart = (spec: any) => {
      const result = originalCreateChart(spec);
      return { ...result, chart_id: ` ${result.chart_id} ` };
    };

    const executor = new ToolExecutor(workbook as SpreadsheetApi, { default_sheet: "Sheet1" });
    const result = await executor.execute({
      name: "create_chart",
      parameters: {
        chart_type: "bar",
        data_range: "A1:B3",
        title: " Sales "
      }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("create_chart");
    if (!result.ok || result.tool !== "create_chart") throw new Error("Unexpected tool result");
    expect(result.data?.status).toBe("ok");
    expect(result.data?.chart_id).toBe("chart_1");
    expect(result.data?.title).toBe("Sales");
    expect(workbook.listCharts()[0]?.spec.title).toBe("Sales");
  });

  it("create_chart returns not_implemented when SpreadsheetApi lacks chart support", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]) as any;
    workbook.createChart = undefined;
    const executor = new ToolExecutor(workbook as SpreadsheetApi, { default_sheet: "Sheet1" });

    const result = await executor.execute({
      name: "create_chart",
      parameters: {
        chart_type: "bar",
        data_range: "A1:B3"
      }
    });

    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("not_implemented");
  });

  it("create_chart validates position as an A1 reference", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { default_sheet: "Sheet1" });

    const result = await executor.execute({
      name: "create_chart",
      parameters: {
        chart_type: "bar",
        data_range: "A1:B3",
        position: "NotACell"
      }
    });

    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("validation_error");
  });

  it("fetch_external_data json_to_table writes headers + rows and returns provenance metadata", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { allow_external_data: true, allowed_external_hosts: ["api.example.com"] });

    const payload = JSON.stringify([
      { a: 1, b: "two" },
      { a: 3, b: "four" }
    ]);
    const payloadBytes = Buffer.byteLength(payload);

    const fetchMock = vi.fn(async (_url: string, _init?: any) => {
      return new Response(payload, {
        status: 200,
        headers: {
          "content-type": "application/json",
          "content-length": String(payloadBytes)
        }
      });
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://api.example.com/data",
        destination: "Sheet1!A1",
        transform: "json_to_table"
      }
    });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls[0]?.[1]).toMatchObject({
      credentials: "omit",
      cache: "no-store",
      referrerPolicy: "no-referrer",
      redirect: "manual"
    });
    expect(result.ok).toBe(true);
    expect(result.tool).toBe("fetch_external_data");
    if (!result.ok || result.tool !== "fetch_external_data") throw new Error("Unexpected tool result");
    if (!result.data) throw new Error("Expected fetch_external_data to return data");

    expect(result.data).toMatchObject({
      url: "https://api.example.com/data",
      destination: "Sheet1!A1",
      status_code: 200,
      content_type: "application/json",
      content_length_bytes: payloadBytes,
      written_cells: 6,
      shape: { rows: 3, cols: 2 }
    });
    expect(typeof result.data.fetched_at_ms).toBe("number");

    const written = workbook
      .readRange(parseA1Range("Sheet1!A1:B3"))
      .map((row) => row.map((cell) => cell.value));
    expect(written).toEqual([
      ["a", "b"],
      [1, "two"],
      [3, "four"]
    ]);
  });

  it("fetch_external_data enforces max_tool_range_cells before writing large json_to_table results", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { allow_external_data: true, allowed_external_hosts: ["api.example.com"] });

    // 201 * 1000 = 201,000 cells (exceeds the default max_tool_range_cells of 200k).
    const rows = 201;
    const cols = 1000;
    const data = Array.from({ length: rows }, () => Array.from({ length: cols }, () => 0));
    const payload = JSON.stringify(data);
    const payloadBytes = Buffer.byteLength(payload);

    const fetchMock = vi.fn(async (_url: string, _init?: any) => {
      return new Response(payload, {
        status: 200,
        headers: {
          "content-type": "application/json",
          "content-length": String(payloadBytes)
        }
      });
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const writeSpy = vi.spyOn(workbook, "writeRange");

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://api.example.com/data",
        destination: "Sheet1!A1",
        transform: "json_to_table"
      }
    });

    expect(result.ok).toBe(false);
    expect(result.tool).toBe("fetch_external_data");
    expect(result.error?.code).toBe("permission_denied");
    expect(result.error?.message).toContain("max_tool_range_cells");
    expect(writeSpy).not.toHaveBeenCalled();
  });

  it("fetch_external_data raw_text writes a single cell and returns provenance metadata", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { allow_external_data: true, allowed_external_hosts: ["example.com"] });

    const payload = "hello world";
    const payloadBytes = Buffer.byteLength(payload);
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => {
        return new Response(payload, {
          status: 200,
          headers: {
            "content-type": "text/plain",
            "content-length": String(payloadBytes)
          }
        });
      }) as any
    );

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://example.com/raw",
        destination: "Sheet1!C3",
        transform: "raw_text"
      }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("fetch_external_data");
    if (!result.ok || result.tool !== "fetch_external_data") throw new Error("Unexpected tool result");
    if (!result.data) throw new Error("Expected fetch_external_data to return data");

    expect(result.data).toMatchObject({
      url: "https://example.com/raw",
      destination: "Sheet1!C3",
      status_code: 200,
      content_type: "text/plain",
      content_length_bytes: payloadBytes,
      written_cells: 1,
      shape: { rows: 1, cols: 1 }
    });
    expect(typeof result.data.fetched_at_ms).toBe("number");
    expect(workbook.getCell(parseA1Cell("Sheet1!C3")).value).toBe(payload);
  });

  it("fetch_external_data enforces host allowlist", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { allow_external_data: true, allowed_external_hosts: ["allowed.example.com"] });

    const fetchMock = vi.fn(async () => {
      throw new Error("fetch should not be called for denied hosts");
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://denied.example.com/data",
        destination: "Sheet1!A1"
      }
    });

    expect(fetchMock).not.toHaveBeenCalled();
    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("permission_denied");
  });

  it("fetch_external_data allowlist matching is case-insensitive and trims whitespace", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { allow_external_data: true, allowed_external_hosts: ["  API.EXAMPLE.COM  "] });

    const fetchMock = vi.fn(async (_url: string, _init?: any) => {
      return new Response(JSON.stringify([{ ok: true }]), {
        status: 200,
        headers: { "content-type": "application/json" }
      });
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://api.example.com/data",
        destination: "Sheet1!A1"
      }
    });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(result.ok).toBe(true);
  });

  it("fetch_external_data allowlist matches hostname entries even when the URL includes an explicit port", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { allow_external_data: true, allowed_external_hosts: ["api.example.com"] });

    const fetchMock = vi.fn(async (_url: string, _init?: any) => {
      return new Response(JSON.stringify([{ ok: true }]), {
        status: 200,
        headers: { "content-type": "application/json" }
      });
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const defaultPortResult = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://api.example.com:443/data",
        destination: "Sheet1!A1"
      }
    });

    const nonDefaultPortResult = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://api.example.com:8443/data",
        destination: "Sheet1!A3"
      }
    });

    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(defaultPortResult.ok).toBe(true);
    expect(nonDefaultPortResult.ok).toBe(true);
  });

  it("fetch_external_data allowlist entries with ports require an exact host:port match", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { allow_external_data: true, allowed_external_hosts: ["api.example.com:8443"] });

    const fetchMock = vi.fn(async (_url: string, _init?: any) => {
      return new Response(JSON.stringify([{ ok: true }]), {
        status: 200,
        headers: { "content-type": "application/json" }
      });
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const allowed = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://api.example.com:8443/data",
        destination: "Sheet1!A1"
      }
    });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(allowed.ok).toBe(true);

    const denied = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://api.example.com:443/data",
        destination: "Sheet1!A3"
      }
    });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(denied.ok).toBe(false);
    expect(denied.error?.code).toBe("permission_denied");
  });

  it("fetch_external_data allowlist entries with default ports match even when the URL omits the port", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { allow_external_data: true, allowed_external_hosts: ["api.example.com:443"] });

    const fetchMock = vi.fn(async (_url: string, _init?: any) => {
      return new Response(JSON.stringify([{ ok: true }]), {
        status: 200,
        headers: { "content-type": "application/json" }
      });
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const httpsImplicitDefaultPort = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://api.example.com/data",
        destination: "Sheet1!A1"
      }
    });
    expect(httpsImplicitDefaultPort.ok).toBe(true);

    const httpImplicitDefaultPort = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "http://api.example.com/data",
        destination: "Sheet1!A3"
      }
    });
    expect(httpImplicitDefaultPort.ok).toBe(false);
    expect(httpImplicitDefaultPort.error?.code).toBe("permission_denied");

    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("fetch_external_data allowlist supports IPv6 hosts (hostname-only and host:port)", async () => {
    const fetchMock = vi.fn(async (_url: string, _init?: any) => {
      return new Response(JSON.stringify([{ ok: true }]), {
        status: 200,
        headers: { "content-type": "application/json" }
      });
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const hostOnlyWorkbook = new InMemoryWorkbook(["Sheet1"]);
    const hostOnlyExecutor = new ToolExecutor(hostOnlyWorkbook, {
      allow_external_data: true,
      allowed_external_hosts: ["[::1]"]
    });
    const hostOnlyResult = await hostOnlyExecutor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "http://[::1]:8080/data",
        destination: "Sheet1!A1"
      }
    });
    expect(hostOnlyResult.ok).toBe(true);

    const hostPortWorkbook = new InMemoryWorkbook(["Sheet1"]);
    const hostPortExecutor = new ToolExecutor(hostPortWorkbook, {
      allow_external_data: true,
      allowed_external_hosts: ["[::1]:8080"]
    });
    const hostPortAllowed = await hostPortExecutor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "http://[::1]:8080/data",
        destination: "Sheet1!A1"
      }
    });
    expect(hostPortAllowed.ok).toBe(true);

    const hostPortDenied = await hostPortExecutor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "http://[::1]:8081/data",
        destination: "Sheet1!A3"
      }
    });
    expect(hostPortDenied.ok).toBe(false);
    expect(hostPortDenied.error?.code).toBe("permission_denied");

    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it("fetch_external_data is disabled by default (requires allow_external_data)", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook);

    const fetchMock = vi.fn(async () => {
      return new Response("ok", { status: 200 });
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://example.com/data",
        destination: "Sheet1!A1"
      }
    });

    expect(fetchMock).not.toHaveBeenCalled();
    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("permission_denied");
  });

  it("fetch_external_data blocks allowlist bypass via redirects", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { allow_external_data: true, allowed_external_hosts: ["api.example.com"] });

    const fetchMock = vi.fn(async (url: string) => {
      if (url === "https://api.example.com/start") {
        return new Response(null, {
          status: 302,
          headers: { location: "https://evil.example.com/data" }
        });
      }
      return new Response(JSON.stringify([{ should: "not fetch" }]), { status: 200 });
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://api.example.com/start",
        destination: "Sheet1!A1"
      }
    });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("permission_denied");
  });

  it("fetch_external_data follows allowlisted redirects and returns the final URL", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { allow_external_data: true, allowed_external_hosts: ["api.example.com"] });

    const payload = JSON.stringify([{ foo: "bar" }]);
    const payloadBytes = Buffer.byteLength(payload);

    const fetchMock = vi.fn(async (url: string) => {
      if (url === "https://api.example.com/start") {
        return new Response(null, {
          status: 302,
          headers: { location: "/final" }
        });
      }
      if (url === "https://api.example.com/final") {
        return new Response(payload, {
          status: 200,
          headers: {
            "content-type": "application/json",
            "content-length": String(payloadBytes)
          }
        });
      }
      throw new Error(`Unexpected URL: ${url}`);
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://api.example.com/start",
        destination: "Sheet1!A1"
      }
    });

    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(result.ok).toBe(true);
    expect(result.tool).toBe("fetch_external_data");
    if (!result.ok || result.tool !== "fetch_external_data") throw new Error("Unexpected tool result");
    if (!result.data) throw new Error("Expected fetch_external_data to return data");

    expect(result.data.url).toBe("https://api.example.com/final");
    expect(result.data.content_length_bytes).toBe(payloadBytes);

    const written = workbook
      .readRange(parseA1Range("Sheet1!A1:A2"))
      .map((row) => row.map((cell) => cell.value));
    expect(written).toEqual([["foo"], ["bar"]]);
  });

  it("fetch_external_data blocks redirects that downgrade https to http", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { allow_external_data: true, allowed_external_hosts: ["api.example.com"] });

    const fetchMock = vi.fn(async (url: string) => {
      if (url === "https://api.example.com/start") {
        return new Response(null, {
          status: 302,
          headers: { location: "http://api.example.com/insecure" }
        });
      }
      throw new Error("Should not follow https->http redirects");
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://api.example.com/start",
        destination: "Sheet1!A1"
      }
    });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("permission_denied");
  });

  it("fetch_external_data drops user-supplied headers when redirecting to a different host", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, {
      allow_external_data: true,
      allowed_external_hosts: ["api.example.com", "download.example.com"]
    });

    const payload = JSON.stringify([{ foo: "bar" }]);

    const fetchMock = vi.fn(async (url: string, _init?: any) => {
      if (url === "https://api.example.com/start") {
        return new Response(null, {
          status: 302,
          headers: { location: "https://download.example.com/final" }
        });
      }
      if (url === "https://download.example.com/final") {
        return new Response(payload, {
          status: 200,
          headers: {
            "content-type": "application/json"
          }
        });
      }
      throw new Error(`Unexpected URL: ${url}`);
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://api.example.com/start",
        destination: "Sheet1!A1",
        headers: { Authorization: "Bearer SECRET" }
      }
    });

    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(fetchMock.mock.calls[0]?.[1]).toMatchObject({ headers: { Authorization: "Bearer SECRET" } });
    expect((fetchMock.mock.calls[1]?.[1] as any)?.headers).toBeUndefined();

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("fetch_external_data");
    if (!result.ok || result.tool !== "fetch_external_data") throw new Error("Unexpected tool result");
    if (!result.data) throw new Error("Expected fetch_external_data to return data");
    expect(result.data.url).toBe("https://download.example.com/final");
  });

  it("fetch_external_data drops headers when redirects are opaque (browser opaqueredirect)", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { allow_external_data: true, allowed_external_hosts: ["api.example.com"] });

    const payload = JSON.stringify([{ foo: "bar" }]);
    const fetchMock = vi.fn(async (_url: string, init?: any) => {
      if (init?.redirect === "manual") {
        return { type: "opaqueredirect" } as any;
      }
      if (init?.redirect === "follow") {
        return new Response(payload, {
          status: 200,
          headers: {
            "content-type": "application/json"
          }
        });
      }
      throw new Error("Unexpected fetch init");
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://api.example.com/start",
        destination: "Sheet1!A1",
        headers: { Authorization: "Bearer SECRET" }
      }
    });

    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(fetchMock.mock.calls[0]?.[1]).toMatchObject({ redirect: "manual", headers: { Authorization: "Bearer SECRET" } });
    expect((fetchMock.mock.calls[1]?.[1] as any)?.headers).toBeUndefined();

    expect(result.ok).toBe(true);
  });

  it("fetch_external_data enforces max_external_bytes using content-length header", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, {
      allow_external_data: true,
      allowed_external_hosts: ["example.com"],
      max_external_bytes: 5
    });

    const fetchMock = vi.fn(async () => {
      return new Response("hello world", {
        status: 200,
        headers: {
          "content-type": "text/plain",
          "content-length": "100"
        }
      });
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://example.com/large",
        destination: "Sheet1!A1",
        transform: "raw_text"
      }
    });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("permission_denied");
    expect(result.error?.message).toMatch(/too large/i);
  });

  it("fetch_external_data streams response bodies and reports actual byte length when content-length is missing", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, {
      allow_external_data: true,
      allowed_external_hosts: ["example.com"],
      max_external_bytes: 100
    });

    const payload = "hello";
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new TextEncoder().encode(payload));
        controller.close();
      }
    });

    const fetchMock = vi.fn(async () => {
      return new Response(stream, {
        status: 200,
        headers: {
          "content-type": "text/plain"
        }
      });
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://example.com/stream-ok",
        destination: "Sheet1!A1",
        transform: "raw_text"
      }
    });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(result.ok).toBe(true);
    expect(result.tool).toBe("fetch_external_data");
    if (!result.ok || result.tool !== "fetch_external_data") throw new Error("Unexpected tool result");
    if (!result.data) throw new Error("Expected fetch_external_data to return data");

    expect(result.data.content_length_bytes).toBe(Buffer.byteLength(payload));
    expect(workbook.getCell(parseA1Cell("Sheet1!A1")).value).toBe(payload);
  });

  it("fetch_external_data enforces max_external_bytes while streaming even when content-length is underreported", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, {
      allow_external_data: true,
      allowed_external_hosts: ["example.com"],
      max_external_bytes: 5
    });

    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new TextEncoder().encode("hello"));
        controller.enqueue(new TextEncoder().encode("world"));
        controller.close();
      }
    });

    const fetchMock = vi.fn(async () => {
      return new Response(stream, {
        status: 200,
        headers: {
          "content-type": "text/plain",
          // The executor still must enforce max_external_bytes even if the declared size is
          // missing or inaccurate.
          "content-length": "5"
        }
      });
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://example.com/stream",
        destination: "Sheet1!A1",
        transform: "raw_text"
      }
    });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("permission_denied");
    expect(result.error?.message).toMatch(/too large/i);
  });

  it("fetch_external_data blocks non-http(s) URLs", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { allow_external_data: true, allowed_external_hosts: ["example.com"] });

    const fetchMock = vi.fn(async () => {
      throw new Error("fetch should not be called for invalid protocols");
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "file:///etc/passwd",
        destination: "Sheet1!A1"
      }
    });

    expect(fetchMock).not.toHaveBeenCalled();
    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("permission_denied");
  });

  it("fetch_external_data rejects URLs with embedded credentials", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { allow_external_data: true, allowed_external_hosts: ["example.com"] });

    const fetchMock = vi.fn(async () => {
      throw new Error("fetch should not be called for credentialed URLs");
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://user:pass@example.com/data",
        destination: "Sheet1!A1"
      }
    });

    expect(fetchMock).not.toHaveBeenCalled();
    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("permission_denied");
  });

  it("fetch_external_data requires an explicit allowed_external_hosts allowlist", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { allow_external_data: true });

    const fetchMock = vi.fn(async () => {
      throw new Error("fetch should not be called without host allowlist");
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://example.com/data",
        destination: "Sheet1!A1"
      }
    });

    expect(fetchMock).not.toHaveBeenCalled();
    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("permission_denied");
    expect(result.error?.message).toMatch(/allowlist/i);
  });

  it("fetch_external_data redacts sensitive query parameters in returned url metadata", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { allow_external_data: true, allowed_external_hosts: ["api.example.com"] });

    const payload = JSON.stringify([{ foo: "bar" }]);
    const fetchMock = vi.fn(async () => {
      return new Response(payload, {
        status: 200,
        headers: { "content-type": "application/json" }
      });
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const result = await executor.execute({
      name: "fetch_external_data",
      parameters: {
        source_type: "api",
        url: "https://api.example.com/data?api_key=SECRET&city=berlin&ACCESS_TOKEN=SECRET2&client_secret=SECRET3#frag",
        destination: "Sheet1!A1"
      }
    });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(result.ok).toBe(true);
    expect(result.tool).toBe("fetch_external_data");
    if (!result.ok || result.tool !== "fetch_external_data") throw new Error("Unexpected tool result");
    if (!result.data) throw new Error("Expected fetch_external_data to return data");

    expect(result.data.url).toContain("api_key=REDACTED");
    expect(result.data.url).toContain("ACCESS_TOKEN=REDACTED");
    expect(result.data.url).toContain("client_secret=REDACTED");
    expect(result.data.url).toContain("city=berlin");
    expect(result.data.url).not.toContain("SECRET");
    expect(result.data.url).not.toContain("frag");
  });
});
