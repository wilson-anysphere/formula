import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import path from "node:path";
import { fileURLToPath } from "node:url";
import https from "node:https";
import type { IncomingMessage } from "node:http";
import crypto from "node:crypto";
import { runMigrations } from "../db/migrations";
import { createMetrics } from "../observability/metrics";
import type { EnabledSiemOrg, SiemConfigProvider } from "../siem/configProvider";
import type { SiemEndpointConfig } from "../siem/types";
import { SiemExportWorker } from "../siem/worker";
import { closeCachedOrgTlsAgentsForTests, sha256FingerprintHexFromCertRaw } from "../http/tls";

const TEST_CERT_PEM = `-----BEGIN CERTIFICATE-----
MIIDJTCCAg2gAwIBAgIUehlPQe+ayYtcheci4TTfm9Ds1RQwDQYJKoZIhvcNAQEL
BQAwFDESMBAGA1UEAwwJMTI3LjAuMC4xMB4XDTI2MDExMjAxMDc1M1oXDTM2MDEx
MDAxMDc1M1owFDESMBAGA1UEAwwJMTI3LjAuMC4xMIIBIjANBgkqhkiG9w0BAQEF
AAOCAQ8AMIIBCgKCAQEAm5Acyl8hQDP/X17wdQzTlumRNzbXpof8AVqgN4KITjA1
8FLlx6AWiTlkp2sTw8iqVlAw0rBvHlzA0sbkqyghtq6OXfjpDZFrBTY4nogIriCA
kX6rkGCi13DKrUcm6+4wQT4zRbc2HSFRI4Sq6tZoXx17EutNY96GRTOb7cAwAfUL
44i+It9NBkMRzfEwkFHAAMBDaN1oJVRxYr7matnH94u+VM4sKK+Erlmg9RXORNhv
BGSD7iqtgKnJwigtgIP7NZZwCzBC5t8BPC3wx04bwqERxqvOkOhgE9MT+CQ412Jc
xLJ87lyWZRIUbfwd+JpwCJG+hyogiPwcUoYYuCNAnQIDAQABo28wbTAdBgNVHQ4E
FgQU+hjrMxbjA4NUI12uptz2RygGavowHwYDVR0jBBgwFoAU+hjrMxbjA4NUI12u
ptz2RygGavowDwYDVR0TAQH/BAUwAwEB/zAaBgNVHREEEzARhwR/AAABgglsb2Nh
bGhvc3QwDQYJKoZIhvcNAQELBQADggEBAAiL/Rco5SffhNl60uMpFzgb5KrqZxKr
O1yXg9bWubQjqfJPjXAtQ7GGbcbj3pbD9A9s/2Z2am/DsGiuhRQ42t0fD6Zm2NZH
eJm4xHWusdYmrNFN7M7Cqj9ZDl0tpmbtTOcPd6gFyJ2a3yTXCr3mV1TmNtp2bgP7
6ibvupxyzMBti17Tdrw2EBY2oI5/GGCXPZ9SL+TUg11YnNqBWUzd3OTOMztKepde
SckxtIutzuEs4+DGq6mHTpXAF905PeD9Q0t9VYxxKBZknivh92ONEjcBc2z+n+li
3mEj9WnCLI35XC5YepctvYyI0VgDYEdvTQXeo3GZ0iTnUNuO/NMGnmI=
-----END CERTIFICATE-----`;

const TEST_KEY_PEM = `-----BEGIN PRIVATE KEY-----
MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCbkBzKXyFAM/9f
XvB1DNOW6ZE3Ntemh/wBWqA3gohOMDXwUuXHoBaJOWSnaxPDyKpWUDDSsG8eXMDS
xuSrKCG2ro5d+OkNkWsFNjieiAiuIICRfquQYKLXcMqtRybr7jBBPjNFtzYdIVEj
hKrq1mhfHXsS601j3oZFM5vtwDAB9QvjiL4i300GQxHN8TCQUcAAwENo3WglVHFi
vuZq2cf3i75Uziwor4SuWaD1Fc5E2G8EZIPuKq2AqcnCKC2Ag/s1lnALMELm3wE8
LfDHThvCoRHGq86Q6GAT0xP4JDjXYlzEsnzuXJZlEhRt/B34mnAIkb6HKiCI/BxS
hhi4I0CdAgMBAAECggEACIwGtOy8njMtKv+DnR6/ClnbX2n9N7pdc3KX/mzG0eru
r71SQCFQ06nKWLN2osll6Hef8xd8B3JHqt0AJ9I85fVZv5qDLXppo5/qxPUK6wxA
nB3WTcitccJR9GrGHezYjGEfPouWJswTkezCWkQ8+Erdnfi9KAlMHcW74bhvOtA2
njYMPuAuBMpwL4VV/PQTpCdxjf8dR2pVBc9BihDeRUe2HkcOk6sSdFFUTa8dKqL6
2f0xxbQNwDWIeh1WL0OaLW6qrgEkleuXDIBhaFfXlnFU4T8eoL635Zi0POw7xpB6
XNilheBTmN+X2ljQ/jm6/2CfNtKnd003vsEYVaQnOQKBgQDYjdOK083NLnlKDJba
4cpgzAx053W5sZ+h0YZTNyRV2Mm4bT3Ec/S+N5k/CPvdiWQECjKwOY6Nor8P/tzD
93OHhUUhEXibVpRthOlB7OnflzZf0L7Yu39QkKxg3dpX1mzaF2lWDmVyzeN4W2UV
sZime0Ixx7QXOd5QgYKz2YZzdQKBgQC35jMCJrVv62DYBX9Ewg7SRs7hJ527CK5B
rz1HMecC81ZNNwT+4FGdd4f6DLRxeEavn32CGB8a2omTu86ifoD71fBKJYHvVLM7
o4tJwYscossPn90URcDj1jhTU7Nsz6vXlBBNSHO8t7+YLzG2tYu727ix2KgDqPgT
hOFnV1W7iQKBgQDESsKey2B4BRFCMuknHJXCahM8gHXwzXXiSzcUBR61hh1LRBJC
Gc2WAnWxcqZC6H+1Pb02ieWSsxu3FxDrvUiGZiIEWH7XZ4KBR4HcFTDlUH6kGWZ0
tHgyAgGOiGqbRi1C/wenTsNcbg4rkcSuBl5VQdL9poSyrOy8Uriz54/85QKBgBpA
HungsqeONt2/MyKSfOEhQGi9afOH0rAHnryp7+ro51nQT8M+LAhJRry30Y4c+CIb
pyPJ467GoTrYZS+m1SydplY/MmQCeC88MibOHNhymH/bdwhsyJL9Qj8KxKL0pff4
57bQb8zxgcTsf7EwCwk+3QduANW86eSHZFGHEvLZAoGBANNYWvhD3huQDRbVAmk7
3rZl3vu1i1VWbof7onJQqMR3tLFUuejSIdC4zlhbaM8hclUg/wNJaN1qYvXdMQko
rkEiPZhcqbtqertrLLuOyAd4cCPIR2tieJGxYDgeX+eYYN5heSgAVyqyyuyM446f
L1t+7qETE5o02rb2iisNYBvR
-----END PRIVATE KEY-----`;

function getMigrationsDir(): string {
  const here = path.dirname(fileURLToPath(import.meta.url));
  // services/api/src/__tests__ -> services/api/migrations
  return path.resolve(here, "../../migrations");
}

type CapturedRequest = {
  headers: IncomingMessage["headers"];
  body: string;
  statusCode: number;
};

function toColonSeparatedHex(hex: string): string {
  return hex.match(/.{1,2}/g)?.join(":") ?? hex;
}

async function startSiemHttpsServer(): Promise<{
  url: string;
  requests: CapturedRequest[];
  close: () => Promise<void>;
}> {
  const requests: CapturedRequest[] = [];

  const server = https.createServer({ key: TEST_KEY_PEM, cert: TEST_CERT_PEM }, (req, res) => {
    const chunks: Buffer[] = [];
    req.on("data", (chunk) => chunks.push(Buffer.from(chunk)));
    req.on("end", () => {
      requests.push({
        headers: req.headers,
        body: Buffer.concat(chunks).toString("utf8"),
        statusCode: 200
      });

      res.writeHead(200, { "content-type": "text/plain" });
      res.end("ok");
    });
  });

  await new Promise<void>((resolve) => {
    server.listen(0, "127.0.0.1", () => resolve());
  });

  const address = server.address();
  if (!address || typeof address === "string") throw new Error("expected server to listen on tcp");

  return {
    url: `https://127.0.0.1:${address.port}/ingest`,
    requests,
    close: async () => {
      await new Promise<void>((resolve, reject) => {
        server.close((err) => (err ? reject(err) : resolve()));
      });
    }
  };
}

class StaticConfigProvider implements SiemConfigProvider {
  constructor(private readonly orgs: EnabledSiemOrg[]) {}
  async listEnabledOrgs(): Promise<EnabledSiemOrg[]> {
    return this.orgs;
  }
}

async function insertOrg(db: Pool, orgId: string): Promise<void> {
  await db.query("INSERT INTO organizations (id, name) VALUES ($1, $2)", [orgId, "Acme"]);
  await db.query("INSERT INTO org_settings (org_id) VALUES ($1)", [orgId]);
}

async function insertAuditEvent(options: {
  db: Pool;
  id: string;
  orgId: string;
  createdAt: Date;
  eventType: string;
}): Promise<void> {
  await options.db.query(
    `
      INSERT INTO audit_log (id, org_id, event_type, resource_type, success, details, created_at)
      VALUES ($1, $2, $3, 'organization', true, '{}'::jsonb, $4)
    `,
    [options.id, options.orgId, options.eventType, options.createdAt]
  );
}

describe("SIEM outbound TLS policy (certificate pinning)", () => {
  let db: Pool;
  const logger = { info: () => {}, warn: () => {}, error: () => {} };

  beforeAll(async () => {
    const mem = newDb({ autoCreateForeignKeyIndices: true });
    const pgAdapter = mem.adapters.createPg();
    db = new pgAdapter.Pool();
    await runMigrations(db, { migrationsDir: getMigrationsDir() });
  });

  afterAll(async () => {
    await db.end();
    await closeCachedOrgTlsAgentsForTests();
  });

  it("succeeds when certificate pinning is enabled and the pin matches", async () => {
    const orgId = crypto.randomUUID();
    await insertOrg(db, orgId);

    const cert = new crypto.X509Certificate(TEST_CERT_PEM);
    const pin = toColonSeparatedHex(sha256FingerprintHexFromCertRaw(cert.raw));
    await db.query(
      "UPDATE org_settings SET certificate_pinning_enabled = true, certificate_pins = $2 WHERE org_id = $1",
      [orgId, JSON.stringify([pin])]
    );

    const eventId = crypto.randomUUID();
    await insertAuditEvent({
      db,
      id: eventId,
      orgId,
      createdAt: new Date("2025-01-01T01:00:00.000Z"),
      eventType: "test.tls.ok"
    });

    const siem = await startSiemHttpsServer();
    try {
      const config: SiemEndpointConfig = {
        endpointUrl: siem.url,
        format: "json",
        retry: { maxAttempts: 2, baseDelayMs: 1, maxDelayMs: 5, jitter: false }
      };

      const metrics = createMetrics();
      const worker = new SiemExportWorker({
        db,
        configProvider: new StaticConfigProvider([{ orgId, config }]),
        metrics,
        logger,
        pollIntervalMs: 0
      });

      await worker.tick();

      expect(siem.requests).toHaveLength(1);

      const state = await db.query(
        "SELECT last_event_id, consecutive_failures FROM org_siem_export_state WHERE org_id = $1",
        [orgId]
      );
      expect(state.rows[0].last_event_id).toBe(eventId);
      expect(state.rows[0].consecutive_failures).toBe(0);
    } finally {
      await siem.close();
    }
  });

  it("fails hard when certificate pinning is enabled and the pin mismatches", async () => {
    const orgId = crypto.randomUUID();
    await insertOrg(db, orgId);

    await db.query(
      "UPDATE org_settings SET certificate_pinning_enabled = true, certificate_pins = $2 WHERE org_id = $1",
      [orgId, JSON.stringify(["00".repeat(32)])]
    );

    const eventId = crypto.randomUUID();
    await insertAuditEvent({
      db,
      id: eventId,
      orgId,
      createdAt: new Date("2025-01-01T02:00:00.000Z"),
      eventType: "test.tls.fail"
    });

    const siem = await startSiemHttpsServer();
    try {
      const config: SiemEndpointConfig = {
        endpointUrl: siem.url,
        format: "json",
        retry: { maxAttempts: 10, baseDelayMs: 1, maxDelayMs: 5, jitter: false }
      };

      const metrics = createMetrics();
      const worker = new SiemExportWorker({
        db,
        configProvider: new StaticConfigProvider([{ orgId, config }]),
        metrics,
        logger,
        pollIntervalMs: 0
      });

      await worker.tick();

      // Pinning failures occur during the TLS handshake, so we never receive the HTTP request.
      expect(siem.requests).toHaveLength(0);

      const state = await db.query(
        "SELECT last_event_id, consecutive_failures, last_error FROM org_siem_export_state WHERE org_id = $1",
        [orgId]
      );
      expect(state.rows[0].last_event_id).toBeNull();
      expect(state.rows[0].consecutive_failures).toBe(1);
      expect(String(state.rows[0].last_error)).toContain("Certificate pinning failed");

      const errors = await metrics.siemBatchErrorsTotal.get();
      expect(errors.values).toEqual(
        expect.arrayContaining([expect.objectContaining({ labels: { reason: "tls_pinning_failed" }, value: 1 })])
      );
    } finally {
      await siem.close();
    }
  });

  it("rejects http endpoints in production and emits a distinct error metric", async () => {
    const previousEnv = process.env.NODE_ENV;
    process.env.NODE_ENV = "production";
    try {
      const orgId = crypto.randomUUID();
      await insertOrg(db, orgId);

      const eventId = crypto.randomUUID();
      await insertAuditEvent({
        db,
        id: eventId,
        orgId,
        createdAt: new Date("2025-01-01T03:00:00.000Z"),
        eventType: "test.http_in_production"
      });

      const config: SiemEndpointConfig = {
        endpointUrl: "http://example.invalid/ingest",
        format: "json",
        retry: { maxAttempts: 10, baseDelayMs: 1, maxDelayMs: 5, jitter: false }
      };

      const metrics = createMetrics();
      const worker = new SiemExportWorker({
        db,
        configProvider: new StaticConfigProvider([{ orgId, config }]),
        metrics,
        logger,
        pollIntervalMs: 0
      });

      await worker.tick();

      const state = await db.query(
        "SELECT last_event_id, consecutive_failures, last_error FROM org_siem_export_state WHERE org_id = $1",
        [orgId]
      );
      expect(state.rows[0].last_event_id).toBeNull();
      expect(state.rows[0].consecutive_failures).toBe(1);
      expect(String(state.rows[0].last_error)).toContain("SIEM endpoint must use https in production");

      const errors = await metrics.siemBatchErrorsTotal.get();
      expect(errors.values).toEqual(
        expect.arrayContaining([expect.objectContaining({ labels: { reason: "insecure_http_endpoint" }, value: 1 })])
      );
    } finally {
      process.env.NODE_ENV = previousEnv;
    }
  });
});
