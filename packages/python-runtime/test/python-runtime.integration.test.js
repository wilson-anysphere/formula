import test from "node:test";
import assert from "node:assert/strict";
import { NativePythonRuntime } from "@formula/python-runtime/native";
import { MockWorkbook } from "@formula/python-runtime/test-utils";

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

test("native python runtime returns captured stderr output (print)", async () => {
  const workbook = new MockWorkbook();
  const runtime = new NativePythonRuntime({
    timeoutMs: 10_000,
    maxMemoryBytes: 256 * 1024 * 1024,
    permissions: { filesystem: "none", network: "none" },
  });

  const script = `
print("hello from python")
`;

  const result = await runtime.execute(script, { api: workbook });
  assert.equal(result.stdout, "");
  assert.match(result.stderr, /hello from python/);
});

test("native python runtime surfaces captured stderr on failure without losing traceback", async () => {
  const workbook = new MockWorkbook();
  const runtime = new NativePythonRuntime({
    timeoutMs: 10_000,
    maxMemoryBytes: 256 * 1024 * 1024,
    permissions: { filesystem: "none", network: "none" },
  });

  const script = `
print("before boom")
raise Exception("boom")
`;

  await assert.rejects(
    () => runtime.execute(script, { api: workbook }),
    (err) => {
      assert.match(err.stack ?? "", /Traceback/);
      assert.match(err.stderr ?? "", /before boom/);
      return true;
    },
  );
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

test("native python sandbox blocks network bypass attempts via restored import (network=none)", async () => {
  const workbook = new MockWorkbook();
  const runtime = new NativePythonRuntime({
    timeoutMs: 10_000,
    maxMemoryBytes: 256 * 1024 * 1024,
    permissions: { filesystem: "none", network: "none" },
  });

  const script = `
import builtins
import sys
import __main__
import formula

sandbox_mod = sys.modules.get("formula.runtime.sandbox")
apply_sandbox_fn = getattr(__main__, "apply_sandbox", None)

if sandbox_mod is not None:
    builtins.__import__ = sandbox_mod._ORIGINAL_IMPORT
elif apply_sandbox_fn is not None:
    builtins.__import__ = apply_sandbox_fn.__globals__["_ORIGINAL_IMPORT"]

import socket
formula.active_sheet["A1"] = 1
`;

  await assert.rejects(() => runtime.execute(script, { api: workbook }), /Import of 'socket' is not permitted/);
});
