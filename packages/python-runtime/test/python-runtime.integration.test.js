import test from "node:test";
import assert from "node:assert/strict";
import { NativePythonRuntime } from "../src/native-python-runtime.js";
import { MockWorkbook } from "../src/mock-workbook.js";

test("native python runtime can set values and formulas via formula API", async () => {
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

  assert.equal(workbook.get_cell_value({ sheet_id: workbook.activeSheetId, row: 0, col: 0 }), 42);
  assert.equal(workbook.get_cell_value({ sheet_id: workbook.activeSheetId, row: 1, col: 0 }), 84);
});

test("native python sandbox blocks filesystem by default", async () => {
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

  await assert.rejects(() => runtime.execute(script, { api: workbook }), /Filesystem access is not permitted/);
});

