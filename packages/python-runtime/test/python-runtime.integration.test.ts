import { describe, expect, it } from "vitest";

import os from "node:os";
import path from "node:path";
import { promises as fs } from "node:fs";
import http from "node:http";

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

  it("returns captured stderr output (print)", async () => {
    const workbook = new MockWorkbook();
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    const result = await runtime.execute(`print("hello from python")\n`, { api: workbook });
    expect(result.stdout).toBe("");
    expect(result.stderr).toContain("hello from python");
  });

  it("includes captured stderr on thrown errors while preserving the Python traceback", async () => {
    const workbook = new MockWorkbook();
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    await expect(runtime.execute(`print("before boom")\nraise Exception("boom")\n`, { api: workbook })).rejects.toMatchObject(
      {
        stderr: expect.stringContaining("before boom"),
        stack: expect.stringMatching(/Traceback/),
      },
    );
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

  it("blocks low-level network modules by default (_socket)", async () => {
    const workbook = new MockWorkbook();
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    const script = `
import _socket
_socket.socket(_socket.AF_INET, _socket.SOCK_STREAM, 0)
`;

    await expect(runtime.execute(script, { api: workbook })).rejects.toThrow(/Import of '_socket' is not permitted/);
  });

  it("blocks importlib escape hatches for blocked builtins (network=none)", async () => {
    const workbook = new MockWorkbook();
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    const script = `
import importlib._bootstrap as ib
ib._builtin_from_name("_socket")
`;

    await expect(runtime.execute(script, { api: workbook })).rejects.toThrow(/Import of '_socket' is not permitted/);
  });

  it("blocks importlib BuiltinImporter direct load path (network=none)", async () => {
    const workbook = new MockWorkbook();
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    const script = `
import importlib._bootstrap as ib
spec = ib.BuiltinImporter.find_spec("_socket")
mod = ib.BuiltinImporter.create_module(spec)
ib.BuiltinImporter.exec_module(mod)
`;

    await expect(runtime.execute(script, { api: workbook })).rejects.toThrow(/Import of '_socket' is not permitted/);
  });

  it("blocks network access even if a script tries to use sys.modules for socket (network=none)", async () => {
    const workbook = new MockWorkbook();
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    // Regression guard: the sandbox should not pre-import `socket` under network=none.
    // If it is present in sys.modules for any reason, network operations should still
    // be blocked.
    const script = `
import sys

sock_mod = sys.modules.get("socket")
if sock_mod is None:
    import socket as sock_mod

s = sock_mod.socket(sock_mod.AF_INET, sock_mod.SOCK_DGRAM)
s.sendto(b"hi", ("127.0.0.1", 9))
`;

    await expect(runtime.execute(script, { api: workbook })).rejects.toThrow(/not permitted/i);
  });

  it("prevents bypassing network=none by restoring the original import function", async () => {
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

# Best-effort: attempt to find the runner's sandbox module and restore its original
# import implementation. The runtime should not expose this to user scripts.
sandbox_mod = sys.modules.get("formula.runtime.sandbox")
apply_sandbox_fn = getattr(__main__, "apply_sandbox", None)

if sandbox_mod is not None:
    builtins.__import__ = sandbox_mod._ORIGINAL_IMPORT
elif apply_sandbox_fn is not None:
    builtins.__import__ = apply_sandbox_fn.__globals__["_ORIGINAL_IMPORT"]

import socket
formula.active_sheet["A1"] = 1
`;

    await expect(runtime.execute(script, { api: workbook })).rejects.toThrow(/Import of 'socket' is not permitted/);
  });

  it("allows network access to allowlisted hosts (native allowlist sandbox)", async () => {
    const server = http.createServer((_req, res) => {
      res.writeHead(200, { "Content-Type": "text/plain" });
      res.end("ok");
    });
    await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
    const address = server.address();
    if (!address || typeof address === "string") {
      server.close();
      throw new Error("Failed to bind local test server");
    }

    const workbook = new MockWorkbook();
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    const script = `
import socket

sock = socket.create_connection(("127.0.0.1", ${address.port}), timeout=2)
sock.sendall(b"GET / HTTP/1.0\\r\\nHost: 127.0.0.1\\r\\n\\r\\n")
sock.recv(1024)
sock.close()
`;

    try {
      await runtime.execute(script, {
        api: workbook,
        permissions: { filesystem: "none", network: "allowlist", networkAllowlist: ["127.0.0.1"] },
      });
    } finally {
      await new Promise<void>((resolve) => server.close(() => resolve()));
    }
  });

  it("blocks network access to non-allowlisted hosts (native allowlist sandbox)", async () => {
    const server = http.createServer((_req, res) => {
      res.writeHead(200, { "Content-Type": "text/plain" });
      res.end("ok");
    });
    await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
    const address = server.address();
    if (!address || typeof address === "string") {
      server.close();
      throw new Error("Failed to bind local test server");
    }

    const workbook = new MockWorkbook();
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    const script = `
import socket

sock = socket.create_connection(("127.0.0.1", ${address.port}), timeout=2)
`;

    try {
      await expect(
        runtime.execute(script, {
          api: workbook,
          permissions: { filesystem: "none", network: "allowlist", networkAllowlist: ["example.com"] },
        }),
      ).rejects.toThrow(/Network access to ['"]?127\.0\.0\.1(?::\d+)?['"]? is not permitted/);
    } finally {
      await new Promise<void>((resolve) => server.close(() => resolve()));
    }
  });

  it("prevents allowlist bypasses via _socket (native allowlist sandbox)", async () => {
    const server = http.createServer((_req, res) => {
      res.writeHead(200, { "Content-Type": "text/plain" });
      res.end("ok");
    });
    await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
    const address = server.address();
    if (!address || typeof address === "string") {
      server.close();
      throw new Error("Failed to bind local test server");
    }

    const workbook = new MockWorkbook();
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    const script = `
import _socket

s = _socket.socket(_socket.AF_INET, _socket.SOCK_STREAM, 0)
s.settimeout(2)
s.connect(("127.0.0.1", ${address.port}))
`;

    try {
      await expect(
        runtime.execute(script, {
          api: workbook,
          permissions: { filesystem: "none", network: "allowlist", networkAllowlist: ["example.com"] },
        }),
      ).rejects.toThrow(/Network access to '127.0.0.1' is not permitted/);
    } finally {
      await new Promise<void>((resolve) => server.close(() => resolve()));
    }
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
