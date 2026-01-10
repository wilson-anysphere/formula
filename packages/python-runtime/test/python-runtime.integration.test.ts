import { describe, expect, it } from "vitest";

import { NativePythonRuntime } from "../src/native-python-runtime.js";
import { MockWorkbook } from "../src/mock-workbook.js";

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
