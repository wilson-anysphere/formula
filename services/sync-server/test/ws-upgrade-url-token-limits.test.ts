import assert from "node:assert/strict";
import crypto from "node:crypto";
import { mkdtemp, rm } from "node:fs/promises";
import net from "node:net";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import { startSyncServer, waitForCondition } from "./test-helpers.ts";

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
      const key = crypto.randomBytes(16).toString("base64");
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

      const statusCode = Number(match[1]);
      socket.destroy();
      resolve(statusCode);
    });
  });
}

async function getRejectedCount(httpUrl: string, reason: string): Promise<number | null> {
  const res = await fetch(`${httpUrl}/metrics`);
  if (!res.ok) return null;
  const body = await res.text();
  const match = body.match(
    new RegExp(
      `^sync_server_ws_connections_rejected_total\\{[^}]*reason="${reason}"[^}]*\\}\\s+([0-9.]+)$`,
      "m"
    )
  );
  if (!match) return null;
  return Number(match[1]);
}

test("rejects websocket upgrade when token exceeds SYNC_SERVER_MAX_TOKEN_BYTES", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-token-limit-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      // Keep the URL limit disabled so we only exercise the token path.
      SYNC_SERVER_MAX_URL_BYTES: "0",
      SYNC_SERVER_MAX_TOKEN_BYTES: "16",
    },
  });
  t.after(async () => {
    await server.stop();
  });

  const tooLongToken = "x".repeat(17);
  const status = await rawWebSocketUpgradeStatus({
    host: "127.0.0.1",
    port: server.port,
    pathWithQuery: `/token-limit-doc?token=${tooLongToken}`,
  });
  assert.equal(status, 414);

  // Ensure the server stays alive after rejecting the oversized token.
  const health = await fetch(`${server.httpUrl}/healthz`);
  assert.ok(health.ok);

  await waitForCondition(async () => {
    const value = await getRejectedCount(server.httpUrl, "token_too_long");
    return typeof value === "number" && value >= 1;
  }, 5_000, 50);
});

test("rejects websocket upgrade when URL exceeds SYNC_SERVER_MAX_URL_BYTES", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-url-limit-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_MAX_URL_BYTES: "64",
      // Disable token length limit so the URL limit triggers first.
      SYNC_SERVER_MAX_TOKEN_BYTES: "0",
    },
  });
  t.after(async () => {
    await server.stop();
  });

  const longQuery = "x".repeat(80);
  const pathWithQuery = `/url-limit-doc?token=test-token&x=${longQuery}`;
  assert.ok(Buffer.byteLength(pathWithQuery, "utf8") > 64);

  const status = await rawWebSocketUpgradeStatus({
    host: "127.0.0.1",
    port: server.port,
    pathWithQuery,
  });
  assert.equal(status, 414);

  const health = await fetch(`${server.httpUrl}/healthz`);
  assert.ok(health.ok);

  await waitForCondition(async () => {
    const value = await getRejectedCount(server.httpUrl, "url_too_long");
    return typeof value === "number" && value >= 1;
  }, 5_000, 50);
});

