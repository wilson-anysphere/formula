import { describe, expect, it } from "vitest";

import os from "node:os";
import path from "node:path";
import { promises as fs } from "node:fs";

import { NativePythonRuntime } from "@formula/python-runtime/native";
import { MockWorkbook } from "@formula/python-runtime/test-utils";

describe("python runtime integration (native)", () => {
  it("executes a script that writes values and formulas via the formula API", async () => {
    const workbook = new MockWorkbook();
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    const script = `
import formula

sheet = formula.active_sheet
sheet["A1"] = 42
sheet["A2"] = "=A1*2"
`;

    await runtime.execute(script, { api: workbook });

    expect(workbook.get_cell_value({ sheet_id: workbook.activeSheetId, row: 0, col: 0 })).toBe(42);
    expect(workbook.get_cell_value({ sheet_id: workbook.activeSheetId, row: 1, col: 0 })).toBe(84);
  });

  it("supports 1D list assignment for single-row and single-column ranges", async () => {
    const workbook = new MockWorkbook();
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    const script = `
import formula

sheet = formula.active_sheet

sheet["A1:B1"] = [1, 2]
sheet["C1:C2"] = [3, 4]
`;

    await runtime.execute(script, { api: workbook });

    expect(workbook.get_cell_value({ sheet_id: workbook.activeSheetId, row: 0, col: 0 })).toBe(1);
    expect(workbook.get_cell_value({ sheet_id: workbook.activeSheetId, row: 0, col: 1 })).toBe(2);
    expect(workbook.get_cell_value({ sheet_id: workbook.activeSheetId, row: 0, col: 2 })).toBe(3);
    expect(workbook.get_cell_value({ sheet_id: workbook.activeSheetId, row: 1, col: 2 })).toBe(4);
  });

  it("blocks filesystem access by default", async () => {
    const workbook = new MockWorkbook();
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    const script = `
import formula

sheet = formula.active_sheet
sheet["A1"] = 1

with open("some_file.txt", "w") as f:
    f.write("nope")
`;

    await expect(runtime.execute(script, { api: workbook })).rejects.toThrow(/Filesystem access is not permitted/);
  });

  it("allows read-only filesystem access when permitted, but blocks writes", async () => {
    const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-python-"));
    const filePath = path.join(tmpDir, "hello.txt");
    await fs.writeFile(filePath, "hello", "utf8");

    const workbook = new MockWorkbook();
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    const script = `
import formula

with open(${JSON.stringify(filePath)}, "r") as f:
    data = f.read()

sheet = formula.active_sheet
sheet["A1"] = len(data)

try:
    with open(${JSON.stringify(filePath)}, "w") as f:
        f.write("nope")
except Exception as e:
    sheet["A2"] = str(e)
`;

    await runtime.execute(script, { api: workbook, permissions: { filesystem: "read", network: "none" } });

    expect(workbook.get_cell_value({ sheet_id: workbook.activeSheetId, row: 0, col: 0 })).toBe(5);
    expect(workbook.get_cell_value({ sheet_id: workbook.activeSheetId, row: 1, col: 0 })).toContain(
      "Filesystem write access is not permitted",
    );
  });

  it("blocks common filesystem escape hatches (io.open and os.remove)", async () => {
    const workbook = new MockWorkbook();
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    const script = `
import io
import os

io.open("some_file.txt", "w").write("nope")
os.remove("some_file.txt")
`;

    await expect(runtime.execute(script, { api: workbook })).rejects.toThrow(/Filesystem access is not permitted/);
  });

  it("blocks network imports by default (socket)", async () => {
    const workbook = new MockWorkbook();
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    const script = `
import socket
socket.socket()
`;

    await expect(runtime.execute(script, { api: workbook })).rejects.toThrow(/Import of 'socket' is not permitted/);
  });

  it("blocks obvious command execution escape hatches (os.system)", async () => {
    const workbook = new MockWorkbook();
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    const script = `
import os
os.system("echo should-not-run")
`;

    await expect(runtime.execute(script, { api: workbook })).rejects.toThrow(/Process execution is not permitted/);
  });

  it("enforces script execution timeouts", async () => {
    const workbook = new MockWorkbook();
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    const script = `
import time
time.sleep(10)
`;

    await expect(runtime.execute(script, { api: workbook, timeoutMs: 50 })).rejects.toThrow(/timed out/i);
  });
});
