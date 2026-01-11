import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";
import { mkdtemp, rm } from "node:fs/promises";
import * as http from "node:http";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { createAuditEvent } from "../../audit-core/index.js";
import { SiemExporter } from "../siem/exporter.js";
import { NodeFsOfflineAuditQueue } from "../siem/queue/node_fs.js";

test("NodeFsOfflineAuditQueue flushes persisted events via SiemExporter (idempotency + redaction)", async () => {
  const received = [];

  const server = http.createServer((req, res) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", () => {
      received.push({
        headers: req.headers,
        body: Buffer.concat(chunks).toString("utf8"),
      });
      res.writeHead(200, { "Content-Type": "text/plain" });
      res.end("ok");
    });
  });

  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const { port } = server.address();

  const dir = await mkdtemp(path.join(os.tmpdir(), "siem-queue-integration-"));
  const queue = new NodeFsOfflineAuditQueue({ dirPath: dir, flushBatchSize: 2 });
  const exporter = new SiemExporter({
    endpointUrl: `http://127.0.0.1:${port}/ingest`,
    format: "json",
    flushIntervalMs: 0,
    idempotencyKeyHeader: "Idempotency-Key",
    retry: { maxAttempts: 2, baseDelayMs: 5, maxDelayMs: 20, jitter: false },
  });

  const events = [
    createAuditEvent({
      id: randomUUID(),
      timestamp: "2025-01-01T00:00:00.000Z",
      eventType: "document.opened",
      actor: { type: "user", id: "user_1" },
      context: { orgId: "org_1" },
      success: true,
      details: { token: "secret1" }
    }),
    createAuditEvent({
      id: randomUUID(),
      timestamp: "2025-01-01T00:00:01.000Z",
      eventType: "document.modified",
      actor: { type: "user", id: "user_1" },
      context: { orgId: "org_1" },
      success: true,
      details: { token: "secret2" }
    }),
    createAuditEvent({
      id: randomUUID(),
      timestamp: "2025-01-01T00:00:02.000Z",
      eventType: "document.deleted",
      actor: { type: "user", id: "user_1" },
      context: { orgId: "org_1" },
      success: true,
      details: { token: "secret3" }
    })
  ];

  for (const event of events) await queue.enqueue(event);

  const result = await queue.flushToExporter(exporter);
  assert.equal(result.sent, events.length);

  assert.equal(received.length, 2);
  for (const request of received) {
    assert.match(request.headers["idempotency-key"], /^[0-9a-f]{64}$/i);
    const payload = JSON.parse(request.body);
    for (const item of payload) {
      assert.equal(item.details.token, "[REDACTED]");
    }
  }

  const sentIds = received.flatMap((request) => JSON.parse(request.body).map((evt) => evt.id));
  assert.deepEqual(sentIds.sort(), events.map((evt) => evt.id).sort());

  await exporter.stop({ flush: false });
  await new Promise((resolve) => server.close(resolve));
  await rm(dir, { recursive: true, force: true });
});
