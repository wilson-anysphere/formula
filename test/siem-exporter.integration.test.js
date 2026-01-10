import assert from "node:assert/strict";
import * as http from "node:http";
import test from "node:test";

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
    retry: { maxAttempts: 3, baseDelayMs: 5, maxDelayMs: 50, jitter: false }
  });

  exporter.enqueue({
    id: "evt_1",
    timestamp: "2025-01-01T00:00:00.000Z",
    orgId: "org_1",
    eventType: "document.opened",
    details: { token: "supersecret" }
  });
  exporter.enqueue({
    id: "evt_2",
    timestamp: "2025-01-01T00:00:01.000Z",
    orgId: "org_1",
    eventType: "document.modified",
    details: { password: "p@ssw0rd" }
  });

  await exporter.flush();

  assert.equal(received.length, 2);
  assert.equal(received[0].body, received[1].body);
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
