import test from "node:test";
import assert from "node:assert/strict";
import http from "node:http";
import net from "node:net";
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

test("native python runtime forwards create_sheet index to host API", async () => {
  const workbook = new MockWorkbook();
  // Populate a few sheets so insertion position is observable.
  const secondId = workbook.create_sheet({ name: "Second", index: 1 });
  workbook.create_sheet({ name: "Third", index: 2 });
  workbook.activeSheetId = secondId;
  workbook.selection.sheet_id = secondId;

  const runtime = new NativePythonRuntime({
    timeoutMs: 10_000,
    maxMemoryBytes: 256 * 1024 * 1024,
    permissions: { filesystem: "none", network: "none" },
  });

  const script = `
import formula

formula.create_sheet("Inserted")
formula.create_sheet("AtStart", index=0)
`;

  await runtime.execute(script, { api: workbook });

  const sheetNames = Array.from(workbook.sheets.values(), (sheet) => sheet.name);
  assert.deepEqual(sheetNames, ["AtStart", "Sheet1", "Second", "Inserted", "Third"]);
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

test("native python sandbox blocks low-level network modules by default (_socket)", async () => {
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

  await assert.rejects(() => runtime.execute(script, { api: workbook }), /Import of '_socket' is not permitted/);
});

test("native python sandbox blocks importlib escape hatches for blocked builtins (network=none)", async () => {
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

  await assert.rejects(() => runtime.execute(script, { api: workbook }), /Import of '_socket' is not permitted/);
});

test("native python sandbox blocks importlib BuiltinImporter direct load path (network=none)", async () => {
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

  await assert.rejects(() => runtime.execute(script, { api: workbook }), /Import of '_socket' is not permitted/);
});

test("native python sandbox blocks importlib.reload while sandboxed", async () => {
  const workbook = new MockWorkbook();
  const runtime = new NativePythonRuntime({
    timeoutMs: 10_000,
    maxMemoryBytes: 256 * 1024 * 1024,
    permissions: { filesystem: "none", network: "none" },
  });

  const script = `
import importlib
import math

importlib.reload(math)
`;

  await assert.rejects(() => runtime.execute(script, { api: workbook }), /Reload of 'math' is not permitted/);
});

test("native allowlist sandbox cannot be bypassed via _socket", async () => {
  const server = http.createServer((_req, res) => {
    res.writeHead(200, { "Content-Type": "text/plain" });
    res.end("ok");
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
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
    await assert.rejects(
      () =>
        runtime.execute(script, {
          api: workbook,
          permissions: { filesystem: "none", network: "allowlist", networkAllowlist: ["example.com"] },
        }),
      /Network access to '127.0.0.1' is not permitted/,
    );
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
});

test("native allowlist sandbox cannot be bypassed via sendmsg", async () => {
  const workbook = new MockWorkbook();
  const runtime = new NativePythonRuntime({
    timeoutMs: 10_000,
    maxMemoryBytes: 256 * 1024 * 1024,
    permissions: { filesystem: "none", network: "none" },
  });

  const script = `
import socket

s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
s.sendmsg([b"hi"], [], 0, ("127.0.0.1", 9))
`;

  await assert.rejects(
    () =>
      runtime.execute(script, {
        api: workbook,
        permissions: { filesystem: "none", network: "allowlist", networkAllowlist: ["example.com"] },
      }),
    /Network access to/,
  );
});

test("native allowlist sandbox cannot be bypassed via base socket connect", async () => {
  const server = http.createServer((_req, res) => {
    res.writeHead(200, { "Content-Type": "text/plain" });
    res.end("ok");
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
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

Base = socket.socket.__mro__[1]
s = socket.socket()
s.settimeout(2)
Base.connect(s, ("127.0.0.1", ${address.port}))
`;

  try {
    await assert.rejects(
      () =>
        runtime.execute(script, {
          api: workbook,
          permissions: { filesystem: "none", network: "allowlist", networkAllowlist: ["example.com"] },
        }),
      /Network access to '127.0.0.1' is not permitted/,
    );
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
});

test("native allowlist sandbox cannot be bypassed via original connect globals", async () => {
  const server = http.createServer((_req, res) => {
    res.writeHead(200, { "Content-Type": "text/plain" });
    res.end("ok");
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
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

orig = socket.socket.connect.__globals__.get("_ORIGINAL_SOCKET_CONNECT")
s = socket.socket()
s.settimeout(2)
orig(s, ("127.0.0.1", ${address.port}))
`;

  try {
    await assert.rejects(
      () =>
        runtime.execute(script, {
          api: workbook,
          permissions: { filesystem: "none", network: "allowlist", networkAllowlist: ["example.com"] },
        }),
      /Network access to '127.0.0.1' is not permitted/,
    );
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
});

test("native allowlist sandbox ignores monkeypatched getaddrinfo in create_connection", async () => {
  let hits1271 = 0;
  let hits1272 = 0;

  const server1272 = net.createServer((socket) => {
    hits1272 += 1;
    socket.end();
  });
  await new Promise((resolve, reject) => {
    server1272.listen(0, "127.0.0.2", resolve);
    server1272.on("error", reject);
  });
  const addr1272 = server1272.address();
  if (!addr1272 || typeof addr1272 === "string") {
    server1272.close();
    throw new Error("Failed to bind 127.0.0.2 server");
  }

  const server1271 = net.createServer((socket) => {
    hits1271 += 1;
    socket.end();
  });
  await new Promise((resolve, reject) => {
    server1271.listen(addr1272.port, "127.0.0.1", resolve);
    server1271.on("error", reject);
  });

  const workbook = new MockWorkbook();
  const runtime = new NativePythonRuntime({
    timeoutMs: 10_000,
    maxMemoryBytes: 256 * 1024 * 1024,
    permissions: { filesystem: "none", network: "none" },
  });

  const script = `
import socket

orig = socket.getaddrinfo

def fake_getaddrinfo(host, port, *args, **kwargs):
    # Attempt to redirect connections to a different host.
    return [(socket.AF_INET, socket.SOCK_STREAM, 0, "", ("127.0.0.2", port))]

socket.getaddrinfo = fake_getaddrinfo

sock = socket.create_connection(("127.0.0.1", ${addr1272.port}), timeout=1)
sock.close()
`;

  try {
    await runtime.execute(script, {
      api: workbook,
      permissions: { filesystem: "none", network: "allowlist", networkAllowlist: ["127.0.0.1"] },
    });
  } finally {
    await new Promise((resolve) => server1271.close(resolve));
    await new Promise((resolve) => server1272.close(resolve));
  }

  assert.ok(hits1271 >= 1);
  assert.equal(hits1272, 0);
});

test("native python sandbox blocks process creation escape hatch (os.fork)", async () => {
  const workbook = new MockWorkbook();
  const runtime = new NativePythonRuntime({
    timeoutMs: 10_000,
    maxMemoryBytes: 256 * 1024 * 1024,
    permissions: { filesystem: "none", network: "none" },
  });

  const script = `
import os

pid = os.fork()
if pid == 0:
    os._exit(0)
os.waitpid(pid, 0)
`;

  await assert.rejects(() => runtime.execute(script, { api: workbook }), /Process execution is not permitted/);
});

test("native python sandbox blocks posix_spawn escape hatch", async () => {
  const workbook = new MockWorkbook();
  const runtime = new NativePythonRuntime({
    timeoutMs: 10_000,
    maxMemoryBytes: 256 * 1024 * 1024,
    permissions: { filesystem: "none", network: "none" },
  });

  const script = `
import os
if hasattr(os, "posix_spawn"):
    os.posix_spawn("/bin/echo", ["echo", "should-not-run"], os.environ)
else:
    os.system("echo should-not-run")
`;

  await assert.rejects(() => runtime.execute(script, { api: workbook }), /Process execution is not permitted/);
});
