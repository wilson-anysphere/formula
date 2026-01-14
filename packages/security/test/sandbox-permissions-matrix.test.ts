import fs from "node:fs/promises";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { afterAll, beforeAll, describe, expect, it } from "vitest";

import { AuditLogger, PermissionManager, runExtension, runScript } from "../src/index.js";

function createInMemoryAuditLogger() {
  const events: any[] = [];
  const store = { append: (event: any) => events.push(event) };
  return { auditLogger: new AuditLogger({ store }), events };
}

async function startHttpServer() {
  const server = http.createServer((_req, res) => {
    res.writeHead(200, { "content-type": "text/plain" });
    res.end("ok");
  });

  await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
  const address = server.address();
  if (!address || typeof address === "string") throw new Error("Failed to bind http server");
  const url = `http://127.0.0.1:${address.port}/`;
  return { server, url };
}

// These sandbox tests can be CPU/IO sensitive under heavily parallelized CI shards.
// Keep the timeout comfortably above typical cold-start overhead so we don't flake.
const SANDBOX_TIMEOUT_MS = 30_000;

describe("Sandbox permissions matrix", () => {
  let server: http.Server;
  let serverUrl: string;

  beforeAll(async () => {
    const started = await startHttpServer();
    server = started.server;
    serverUrl = started.url;
  });

  afterAll(async () => {
    await new Promise<void>((resolve, reject) => {
      server.close((err) => (err ? reject(err) : resolve()));
    });
  });

  it("enforces filesystem read/write scopes for JavaScript extensions", async () => {
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-sec-ext-fs-"));
    const filePath = path.join(dir, "data.txt");
    await fs.writeFile(filePath, "hello", "utf8");

    const { auditLogger } = createInMemoryAuditLogger();
    const permissionManager = new PermissionManager({ auditLogger });

    const extensionId = "ext.fs.matrix";
    const principal = { type: "extension", id: extensionId };

    // Read allowed.
    permissionManager.grant(principal, { filesystem: { read: [dir] } });
    await expect(
      runExtension({
        extensionId,
        code: `const data = await fs.readFile(${JSON.stringify(filePath)}, "utf8");\nreturn data;`,
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS
      })
    ).resolves.toBe("hello");

    // Write denied with read-only grant.
    const outPath = path.join(dir, "out.txt");
    await expect(
      runExtension({
        extensionId,
        code: `await fs.writeFile(${JSON.stringify(outPath)}, "nope");`,
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS
      })
    ).rejects.toMatchObject({
      code: "PERMISSION_DENIED",
      request: { kind: "filesystem", access: "readwrite" }
    });

    // Write allowed with readwrite grant.
    permissionManager.grant(principal, { filesystem: { readwrite: [dir] } });
    await expect(
      runExtension({
        extensionId,
        code: `await fs.writeFile(${JSON.stringify(outPath)}, "ok");\nreturn await fs.readFile(${JSON.stringify(
          outPath
        )}, "utf8");`,
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS
      })
    ).resolves.toBe("ok");
  });

  it("enforces network allowlist for JavaScript extensions", async () => {
    const { auditLogger } = createInMemoryAuditLogger();
    const permissionManager = new PermissionManager({ auditLogger });

    const extensionId = "ext.net.allowlist";
    const principal = { type: "extension", id: extensionId };

    const origin = new URL(serverUrl).origin;
    permissionManager.grant(principal, { network: { mode: "allowlist", allowlist: [origin] } });

    await expect(
      runExtension({
        extensionId,
        code: `const res = await fetch(${JSON.stringify(serverUrl)});\nreturn res.status;`,
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS
      })
    ).resolves.toBe(200);

    await expect(
      runExtension({
        extensionId,
        code: `await fetch("http://127.0.0.1:65534/");`,
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS
      })
    ).rejects.toMatchObject({ code: "PERMISSION_DENIED", request: { kind: "network" } });
  });

  it("enforces filesystem/network permissions for Python scripts", async () => {
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-sec-py-"));
    const filePath = path.join(dir, "data.txt");
    await fs.writeFile(filePath, "hello-py", "utf8");
    const runnerPath = fileURLToPath(new URL("../src/sandbox/pythonSandbox.py", import.meta.url));

    const { auditLogger } = createInMemoryAuditLogger();
    const permissionManager = new PermissionManager({ auditLogger });

    const scriptId = "py.matrix";
    const principal = { type: "script", id: scriptId };

    permissionManager.grant(principal, { filesystem: { read: [dir] } });

    await expect(
      runScript({
        scriptId,
        language: "python",
        code: `with open(${JSON.stringify(filePath)}, "r") as f:\n    __result__ = f.read()\n`,
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS
      })
    ).resolves.toBe("hello-py");

    await expect(
      runScript({
        scriptId,
        language: "python",
        code: `import io\nwith io.open(${JSON.stringify(filePath)}, "r") as f:\n    __result__ = f.read()\n`,
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS
      })
    ).resolves.toBe("hello-py");

    if (process.platform !== "win32") {
      await expect(
        runScript({
          scriptId,
          language: "python",
          code: `import posix\nfd = posix.open(${JSON.stringify(filePath)}, posix.O_RDONLY)\ndata = posix.read(fd, 1024)\nposix.close(fd)\n__result__ = data.decode("utf8")\n`,
          permissionManager,
          auditLogger,
          timeoutMs: SANDBOX_TIMEOUT_MS
        })
      ).resolves.toBe("hello-py");
    }

    // Escape hatches must still be permission gated.
    await expect(
      runScript({
        scriptId: "py.fs.io.denied",
        language: "python",
        code: `import io\nio.open(${JSON.stringify(runnerPath)}, "r").read()\n`,
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS
      })
    ).rejects.toMatchObject({ code: "PERMISSION_DENIED", request: { kind: "filesystem", access: "read" } });

    await expect(
      runScript({
        scriptId: "py.fs._io.denied",
        language: "python",
        code: `import _io\n_io.open(${JSON.stringify(runnerPath)}, "r").read()\n`,
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS
      })
    ).rejects.toMatchObject({ code: "PERMISSION_DENIED", request: { kind: "filesystem", access: "read" } });

    if (process.platform !== "win32") {
      await expect(
        runScript({
          scriptId: "py.fs.posix.denied",
          language: "python",
          code: `import posix\nfd = posix.open(${JSON.stringify(runnerPath)}, posix.O_RDONLY)\nposix.read(fd, 10)\n`,
          permissionManager,
          auditLogger,
          timeoutMs: SANDBOX_TIMEOUT_MS
        })
      ).rejects.toMatchObject({ code: "PERMISSION_DENIED", request: { kind: "filesystem", access: "read" } });
    }

    if (process.platform !== "win32") {
      await expect(
        runScript({
          scriptId: "py.automation.denied",
          language: "python",
          code: `import os\nos.spawnv(os.P_WAIT, "/bin/true", ["true"])\n`,
          permissionManager,
          auditLogger,
          timeoutMs: SANDBOX_TIMEOUT_MS
        })
      ).rejects.toMatchObject({ code: "PERMISSION_DENIED", request: { kind: "automation" } });
    }

    await expect(
      runScript({
        scriptId: "py.net.udp.denied",
        language: "python",
        code: `import socket\nsock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)\nsock.sendto(b"hi", ("127.0.0.1", 9))\n`,
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS
      })
    ).rejects.toMatchObject({ code: "PERMISSION_DENIED", request: { kind: "network" } });

    await expect(
      runScript({
        scriptId: "py.net.bind.denied",
        language: "python",
        code: `import socket\nsock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)\nsock.bind(("127.0.0.1", 0))\n`,
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS
      })
    ).rejects.toMatchObject({ code: "PERMISSION_DENIED", request: { kind: "network" } });

    await expect(
      runScript({
        scriptId: "py.net.denied",
        language: "python",
        code: `import urllib.request\nurllib.request.urlopen(${JSON.stringify(serverUrl)}).read()\n`,
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS
      })
    ).rejects.toMatchObject({ code: "PERMISSION_DENIED", request: { kind: "network" } });

    const netPrincipal = { type: "script", id: "py.net.allowed" };
    permissionManager.grant(netPrincipal, {
      network: { mode: "allowlist", allowlist: [new URL(serverUrl).origin] }
    });

    await expect(
      runScript({
        scriptId: "py.net.allowed",
        language: "python",
        code: `import urllib.request\nwith urllib.request.urlopen(${JSON.stringify(
          serverUrl
        )}) as res:\n    __result__ = res.status\n`,
        permissionManager,
        auditLogger,
        timeoutMs: SANDBOX_TIMEOUT_MS
      })
    ).resolves.toBe(200);
  });
});
