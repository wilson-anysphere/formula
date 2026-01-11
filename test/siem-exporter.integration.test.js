import assert from "node:assert/strict";
import * as http from "node:http";
import test from "node:test";

import { createAuditEvent } from "../packages/audit-core/index.js";
import { SiemExporter } from "../packages/security/siem/exporter.js";

test("SiemExporter batches events and retries failed requests", async () => {
  const received = [];
  let requestCount = 0;

  const server = http.createServer((req, res) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", () => {
      requestCount += 1;
      received.push({
        method: req.method,
        url: req.url,
        headers: req.headers,
        body: Buffer.concat(chunks).toString("utf8")
      });

      if (requestCount === 1) {
        res.writeHead(500, { "Content-Type": "text/plain" });
        res.end("try again");
        return;
      }

      res.writeHead(200, { "Content-Type": "text/plain" });
      res.end("ok");
    });
  });

  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const { port } = server.address();

  const exporter = new SiemExporter({
    endpointUrl: `http://127.0.0.1:${port}/ingest`,
    format: "json",
    batchSize: 2,
    flushIntervalMs: 0,
    idempotencyKeyHeader: "Idempotency-Key",
    retry: { maxAttempts: 3, baseDelayMs: 5, maxDelayMs: 50, jitter: false }
  });

  exporter.enqueue(
    createAuditEvent({
      id: "11111111-1111-4111-8111-111111111111",
      timestamp: "2025-01-01T00:00:00.000Z",
      eventType: "document.opened",
      actor: { type: "user", id: "user_1" },
      context: { orgId: "org_1" },
      success: true,
      details: { token: "supersecret" }
    })
  );
  exporter.enqueue(
    createAuditEvent({
      id: "22222222-2222-4222-8222-222222222222",
      timestamp: "2025-01-01T00:00:01.000Z",
      eventType: "document.modified",
      actor: { type: "user", id: "user_1" },
      context: { orgId: "org_1" },
      success: true,
      details: { password: "p@ssw0rd" }
    })
  );

  await exporter.flush();

  assert.equal(received.length, 2);
  assert.equal(received[0].body, received[1].body);
  assert.equal(received[0].headers["idempotency-key"], received[1].headers["idempotency-key"]);
  assert.match(received[1].headers["idempotency-key"], /^[0-9a-f]{64}$/i);
  assert.equal(received[0].method, "POST");
  assert.equal(received[0].url, "/ingest");
  assert.match(received[0].headers["content-type"], /application\/json/);

  const payload = JSON.parse(received[1].body);
  assert.equal(payload.length, 2);
  assert.equal(payload[0].details.token, "[REDACTED]");
  assert.equal(payload[1].details.password, "[REDACTED]");

  await exporter.stop({ flush: false });
  await new Promise((resolve) => server.close(resolve));
});
