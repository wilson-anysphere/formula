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

  it("fetch_external_data json_to_table writes headers + rows and returns provenance metadata", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const executor = new ToolExecutor(workbook, { allow_external_data: true, allowed_external_hosts: ["api.example.com"] });

    const payload = JSON.stringify([
      { a: 1, b: "two" },
      { a: 3, b: "four" }
    ]);
    const payloadBytes = Buffer.byteLength(payload);

    const fetchMock = vi.fn(async () => {
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
        url: "https://api.example.com/data?api_key=SECRET&city=berlin",
        destination: "Sheet1!A1"
      }
    });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(result.ok).toBe(true);
    expect(result.tool).toBe("fetch_external_data");
    if (!result.ok || result.tool !== "fetch_external_data") throw new Error("Unexpected tool result");
    if (!result.data) throw new Error("Expected fetch_external_data to return data");

    expect(result.data.url).toContain("api_key=REDACTED");
    expect(result.data.url).toContain("city=berlin");
    expect(result.data.url).not.toContain("SECRET");
  });
});
