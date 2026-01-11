import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import tls from "node:tls";
import http from "node:http";
import https from "node:https";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import crypto from "node:crypto";
import {
  closeCachedOrgTlsAgentsForTests,
  createPinnedCheckServerIdentity,
  normalizeFingerprintHex,
  sha256FingerprintHexFromCertRaw,
  fetchWithOrgTls
} from "../http/tls";

function fixturesDir(): string {
  const here = path.dirname(fileURLToPath(import.meta.url));
  return path.join(here, "fixtures", "tls");
}

describe("TLS pinning helpers", () => {
  it("normalizes SHA-256 fingerprints (colon-separated vs hex)", () => {
    expect(normalizeFingerprintHex("AA:bb:CC")).toBe("aabbcc");
    expect(normalizeFingerprintHex("aabbcc")).toBe("aabbcc");
  });

  it("rejects pin mismatches", () => {
    const checkSpy = vi.spyOn(tls, "checkServerIdentity").mockReturnValue(undefined);
    try {
      const raw = Buffer.from("test-cert-raw");
      const fingerprint = sha256FingerprintHexFromCertRaw(raw);

      const check = createPinnedCheckServerIdentity({ pins: ["00".repeat(32)] });
      const err = check("example.com", { raw } as any);

      expect(err).toBeInstanceOf(Error);
      expect((err as any).retriable).toBe(false);
      expect((err as Error).message).toContain("fingerprint mismatch");
      expect(fingerprint).not.toBe("00".repeat(32));
    } finally {
      checkSpy.mockRestore();
    }
  });

  it("accepts pin matches", () => {
    const checkSpy = vi.spyOn(tls, "checkServerIdentity").mockReturnValue(undefined);
    try {
      const raw = Buffer.from("another-test-cert");
      const fingerprint = sha256FingerprintHexFromCertRaw(raw);
      const colonSeparated = fingerprint.match(/.{1,2}/g)!.join(":").toUpperCase();

      const check = createPinnedCheckServerIdentity({ pins: [colonSeparated] });
      const err = check("example.com", { raw } as any);
      expect(err).toBeUndefined();
    } finally {
      checkSpy.mockRestore();
    }
  });
});

describe("fetchWithOrgTls", () => {
  let serverUrl: string;
  let closeServer: (() => Promise<void>) | null = null;
  let certPem: Buffer;

  beforeAll(async () => {
    const dir = fixturesDir();
    certPem = fs.readFileSync(path.join(dir, "localhost.crt"));
    const keyPem = fs.readFileSync(path.join(dir, "localhost.key"));

    const server = https.createServer({ key: keyPem, cert: certPem }, (_req, res) => {
      res.writeHead(200, { "content-type": "text/plain" });
      res.end("ok");
    });

    await new Promise<void>((resolve) => {
      server.listen(0, "127.0.0.1", () => resolve());
    });
    const address = server.address();
    if (!address || typeof address === "string") throw new Error("expected https server to listen on tcp");

    serverUrl = `https://127.0.0.1:${address.port}/`;
    closeServer = async () => {
      await new Promise<void>((resolve, reject) => {
        server.close((err) => (err ? reject(err) : resolve()));
      });
    };
  });

  afterAll(async () => {
    await closeServer?.();
    await closeCachedOrgTlsAgentsForTests();
  });

  it("succeeds when pin matches", async () => {
    const x509 = new crypto.X509Certificate(certPem);
    const pin = sha256FingerprintHexFromCertRaw(x509.raw);

    const res = await fetchWithOrgTls(
      serverUrl,
      { method: "GET" },
      {
        tls: {
          certificatePinningEnabled: true,
          certificatePins: [pin],
          ca: certPem
        }
      }
    );

    expect(res.ok).toBe(true);
    await expect(res.text()).resolves.toBe("ok");
  });

  it("fails when pin mismatches", async () => {
    const res = fetchWithOrgTls(
      serverUrl,
      { method: "GET" },
      {
        tls: {
          certificatePinningEnabled: true,
          certificatePins: ["00".repeat(32)],
          ca: certPem
        }
      }
    );

    await expect(res).rejects.toThrow();

    await res.catch((err: any) => {
      expect(err?.cause?.retriable).toBe(false);
      const message =
        typeof err?.cause?.message === "string"
          ? err.cause.message
          : typeof err?.message === "string"
            ? err.message
            : String(err);
      expect(message).toContain("Certificate pinning failed");
    });
  });

  it("fails fast when pinning is enabled for non-https URLs", async () => {
    const server = http.createServer((_req, res) => {
      res.writeHead(200, { "content-type": "text/plain" });
      res.end("ok");
    });

    await new Promise<void>((resolve) => {
      server.listen(0, "127.0.0.1", () => resolve());
    });
    const address = server.address();
    if (!address || typeof address === "string") throw new Error("expected http server to listen on tcp");

    const url = `http://127.0.0.1:${address.port}/`;
    try {
      await fetchWithOrgTls(
        url,
        { method: "GET" },
        { tls: { certificatePinningEnabled: true, certificatePins: ["00".repeat(32)] } }
      );
      throw new Error("expected request to fail");
    } catch (err: any) {
      expect(err).toBeInstanceOf(Error);
      expect(err.message).toContain("requires an https URL");
      expect(err.retriable).toBe(false);
    } finally {
      await new Promise<void>((resolve, reject) => {
        server.close((closeErr) => (closeErr ? reject(closeErr) : resolve()));
      });
    }
  });
});
