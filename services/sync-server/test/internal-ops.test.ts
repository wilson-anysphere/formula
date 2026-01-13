import assert from "node:assert/strict";
import { createHash, randomBytes } from "node:crypto";
import { access, mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import net from "node:net";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";
import { WebsocketProvider, Y } from "./yjs-interop.ts";

import {
  startSyncServer,
  waitForCondition,
  waitForProviderSync,
} from "./test-helpers.ts";

async function fileExists(filePath: string): Promise<boolean> {
  try {
    await access(filePath);
    return true;
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === "ENOENT") return false;
    throw err;
  }
}

async function expectWsUpgradeStatus(
  url: string,
  expectedStatusCode: number
): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    const ws = new WebSocket(url);
    let finished = false;

    const finish = (cb: () => void) => {
      if (finished) return;
      finished = true;
      try {
        ws.terminate();
      } catch {
        // ignore
      }
      cb();
    };

    ws.on("open", () => {
      finish(() => reject(new Error("Expected WebSocket upgrade rejection")));
    });
    ws.on("unexpected-response", (_req, res) => {
      try {
        assert.equal(res.statusCode, expectedStatusCode);
        finish(resolve);
      } catch (err) {
        finish(() => reject(err));
      }
    });
    ws.on("error", (err) => {
      if (finished) return;
      reject(err);
    });
  });
}

async function rawHttpStatus(opts: { host: string; port: number; request: string }): Promise<number> {
  return await new Promise((resolve, reject) => {
    const socket = net.connect(opts.port, opts.host);
    socket.setTimeout(5_000);

    socket.on("timeout", () => {
      socket.destroy();
      reject(new Error("Timed out waiting for response"));
    });
    socket.on("error", reject);

    socket.on("connect", () => {
      socket.write(opts.request);
    });

    let buffer = "";
    socket.setEncoding("utf8");
    socket.on("data", (chunk) => {
      buffer += chunk;
      const headerEnd = buffer.indexOf("\r\n\r\n");
      if (headerEnd < 0) return;

      const header = buffer.slice(0, headerEnd);
      const statusLine = header.split("\r\n")[0] ?? "";
      const match = statusLine.match(/^HTTP\/1\.1\s+(\d+)/i);
      if (!match) {
        socket.destroy();
        reject(new Error(`Unexpected response: ${statusLine}`));
        return;
      }

      socket.destroy();
      resolve(Number(match[1]));
    });
  });
}

async function rawWebSocketUpgradeStatus(opts: {
  host: string;
  port: number;
  pathWithQuery: string;
}): Promise<number> {
  return await new Promise((resolve, reject) => {
    const socket = net.connect(opts.port, opts.host);
    socket.setTimeout(5_000);

    socket.on("timeout", () => {
      socket.destroy();
      reject(new Error("Timed out waiting for upgrade response"));
    });
    socket.on("error", reject);

    socket.on("connect", () => {
      const key = randomBytes(16).toString("base64");
      const request = [
        `GET ${opts.pathWithQuery} HTTP/1.1`,
        `Host: ${opts.host}:${opts.port}`,
        "Upgrade: websocket",
        "Connection: Upgrade",
        `Sec-WebSocket-Key: ${key}`,
        "Sec-WebSocket-Version: 13",
        "",
        "",
      ].join("\r\n");
      socket.write(request);
    });

    let buffer = "";
    socket.setEncoding("utf8");
    socket.on("data", (chunk) => {
      buffer += chunk;
      const headerEnd = buffer.indexOf("\r\n\r\n");
      if (headerEnd < 0) return;

      const header = buffer.slice(0, headerEnd);
      const statusLine = header.split("\r\n")[0] ?? "";
      const match = statusLine.match(/^HTTP\/1\.1\s+(\d+)/i);
      if (!match) {
        socket.destroy();
        reject(new Error(`Unexpected response: ${statusLine}`));
        return;
      }

      socket.destroy();
      resolve(Number(match[1]));
    });
  });
}

test("internal endpoints are disabled without admin token", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const res = await fetch(`${server.httpUrl}/internal/stats`);
  assert.equal(res.status, 404);
});

test("internal endpoints require x-internal-admin-token", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "admin-token",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const missingHeader = await fetch(`${server.httpUrl}/internal/stats`);
  assert.equal(missingHeader.status, 403);

  const wrongHeader = await fetch(`${server.httpUrl}/internal/stats`, {
    headers: { "x-internal-admin-token": "wrong-token" },
  });
  assert.equal(wrongHeader.status, 403);

  const ok = await fetch(`${server.httpUrl}/internal/stats`, {
    headers: { "x-internal-admin-token": "admin-token" },
  });
  assert.equal(ok.status, 200);
  assert.equal(ok.headers.get("cache-control"), "no-store");
  const body = (await ok.json()) as {
    tombstonesCount?: unknown;
    rssBytes?: unknown;
    heapUsedBytes?: unknown;
    heapTotalBytes?: unknown;
    eventLoopDelayMs?: unknown;
    connections?: { activeDocs?: unknown };
  };
  assert.equal(typeof body.tombstonesCount, "number");
  assert.equal(typeof body.rssBytes, "number");
  assert.equal(typeof body.heapUsedBytes, "number");
  assert.equal(typeof body.heapTotalBytes, "number");
  assert.equal(typeof body.eventLoopDelayMs, "number");
  assert.equal(typeof body.connections?.activeDocs, "number");
});

test("purge creates tombstone and prevents doc resurrection", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "admin-token",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const docName = "purge-doc";
  const doc = new Y.Doc();
  const provider = new WebsocketProvider(server.wsUrl, docName, doc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });

  t.after(() => {
    provider.destroy();
    doc.destroy();
  });

  await waitForProviderSync(provider);
  doc.getText("t").insert(0, "goodbye");

  // Give the provider a moment to send the update before disconnecting.
  await new Promise((r) => setTimeout(r, 50));

  provider.destroy();
  doc.destroy();

  const docKey = createHash("sha256").update(docName).digest("hex");
  const persistedPath = path.join(dataDir, `${docKey}.yjs`);

  await waitForCondition(() => fileExists(persistedPath), 10_000);
  assert.equal(await fileExists(persistedPath), true);
  if (process.platform !== "win32") {
    const st = await stat(persistedPath);
    assert.equal(st.mode & 0o777, 0o600);
  }

  const purgeRes = await fetch(
    `${server.httpUrl}/internal/docs/${encodeURIComponent(docName)}`,
    {
      method: "DELETE",
      headers: { "x-internal-admin-token": "admin-token" },
    }
  );
  assert.equal(purgeRes.status, 200);
  assert.deepEqual(await purgeRes.json(), { ok: true });

  await waitForCondition(async () => !(await fileExists(persistedPath)), 10_000);
  assert.equal(await fileExists(persistedPath), false);

  await expectWsUpgradeStatus(
    `${server.wsUrl}/${encodeURIComponent(docName)}?token=test-token`,
    410
  );
});

test("purge decodes url-encoded doc ids", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "admin-token",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  // y-websocket uses the room name as the URL path verbatim; slashes are valid
  // and must be handled by the internal purge endpoint via URL decoding.
  const docName = "purge/doc";

  const doc = new Y.Doc();
  const provider = new WebsocketProvider(server.wsUrl, docName, doc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });

  t.after(() => {
    provider.destroy();
    doc.destroy();
  });

  await waitForProviderSync(provider);
  doc.getText("t").insert(0, "goodbye");
  await new Promise((r) => setTimeout(r, 50));

  provider.destroy();
  doc.destroy();

  const docKey = createHash("sha256").update(docName).digest("hex");
  const persistedPath = path.join(dataDir, `${docKey}.yjs`);
  await waitForCondition(() => fileExists(persistedPath), 10_000);

  const purgeRes = await fetch(
    `${server.httpUrl}/internal/docs/${encodeURIComponent(docName)}`,
    {
      method: "DELETE",
      headers: { "x-internal-admin-token": "admin-token" },
    }
  );
  assert.equal(purgeRes.status, 200);
  assert.deepEqual(await purgeRes.json(), { ok: true });

  await waitForCondition(async () => !(await fileExists(persistedPath)), 10_000);

  await expectWsUpgradeStatus(`${server.wsUrl}/${docName}?token=test-token`, 410);
});

test("purge rejects doc ids that exceed the server's maximum length", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "admin-token",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const tooLongDocName = "a".repeat(1025);
  const res = await fetch(
    `${server.httpUrl}/internal/docs/${encodeURIComponent(tooLongDocName)}`,
    {
      method: "DELETE",
      headers: { "x-internal-admin-token": "admin-token" },
    }
  );

  assert.equal(res.status, 414);
  assert.deepEqual(await res.json(), { error: "doc_id_too_long" });
});

test("purge handles dot-segment doc ids without URL normalization", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "admin-token",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  // URL parsers normalize `.`/`..` segments, but y-websocket treats doc names as
  // raw path strings. The internal purge endpoint should do the same so operators
  // can purge maliciously crafted doc ids.
  const docName = "..";
  const docKey = createHash("sha256").update(docName).digest("hex");
  const persistedPath = path.join(dataDir, `${docKey}.yjs`);
  await writeFile(persistedPath, "dummy");
  assert.equal(await fileExists(persistedPath), true);

  const purgeStatus = await rawHttpStatus({
    host: "127.0.0.1",
    port: server.port,
    request: [
      `DELETE /internal/docs/${docName} HTTP/1.1`,
      `Host: 127.0.0.1:${server.port}`,
      "Connection: close",
      "x-internal-admin-token: admin-token",
      "",
      "",
    ].join("\r\n"),
  });
  assert.equal(purgeStatus, 200);

  await waitForCondition(async () => !(await fileExists(persistedPath)), 10_000);

  const tombstonesLogRaw = await readFile(path.join(dataDir, "tombstones.log"), "utf8");
  if (process.platform !== "win32") {
    const st = await stat(path.join(dataDir, "tombstones.log"));
    assert.equal(st.mode & 0o777, 0o600);
  }
  const records = tombstonesLogRaw
    .trim()
    .split("\n")
    .filter(Boolean)
    .map((line) => JSON.parse(line) as any);
  assert.ok(
    records.some(
      (r) =>
        r &&
        typeof r === "object" &&
        r.op === "set" &&
        r.docKey === docKey &&
        typeof r.deletedAtMs === "number"
    )
  );

  const wsStatus = await rawWebSocketUpgradeStatus({
    host: "127.0.0.1",
    port: server.port,
    pathWithQuery: `/${docName}?token=test-token`,
  });
  assert.equal(wsStatus, 410);
});
