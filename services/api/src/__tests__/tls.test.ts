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
  createTlsConnectOptions,
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

  it("createTlsConnectOptions enforces TLSv1.3 minimum", () => {
    const options = createTlsConnectOptions({ certificatePinningEnabled: false, certificatePins: [] });
    expect(options.minVersion).toBe("TLSv1.3");
  });

  it("createTlsConnectOptions marks default TLS certificate validation errors non-retriable", () => {
    const hostnameError = new Error("Hostname/IP does not match certificate");
    const checkSpy = vi.spyOn(tls, "checkServerIdentity").mockReturnValue(hostnameError);
    try {
      const options = createTlsConnectOptions({ certificatePinningEnabled: false, certificatePins: [] });
      expect(options.checkServerIdentity).toBeTypeOf("function");
      const err = options.checkServerIdentity?.("example.com", { raw: Buffer.from("irrelevant") } as any);
      expect(err).toBe(hostnameError);
      expect((err as any).retriable).toBe(false);
    } finally {
      checkSpy.mockRestore();
    }
  });

  it("createTlsConnectOptions rejects empty pin set when pinning enabled", () => {
    expect(() => createTlsConnectOptions({ certificatePinningEnabled: true, certificatePins: [] })).toThrow(
      /certificatePins must be non-empty/i
    );

    try {
      createTlsConnectOptions({ certificatePinningEnabled: true, certificatePins: [] });
    } catch (err: any) {
      expect(err.retriable).toBe(false);
    }
  });

  it("createTlsConnectOptions rejects invalid pins when pinning enabled", () => {
    expect(() =>
      createTlsConnectOptions({ certificatePinningEnabled: true, certificatePins: ["not-a-fingerprint"] })
    ).toThrow(/certificatePins must be SHA-256/i);

    try {
      createTlsConnectOptions({ certificatePinningEnabled: true, certificatePins: ["not-a-fingerprint"] });
    } catch (err: any) {
      expect(err.retriable).toBe(false);
    }
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
      expect((err as any).code).toBe("ERR_CERT_PINNING");
      expect((err as Error).message).toContain("fingerprint mismatch");
      expect(fingerprint).not.toBe("00".repeat(32));
    } finally {
      checkSpy.mockRestore();
    }
  });

  it("treats TLS hostname validation failures as non-retriable", () => {
    const hostnameError = new Error("Hostname/IP does not match certificate");
    const checkSpy = vi.spyOn(tls, "checkServerIdentity").mockReturnValue(hostnameError);
    try {
      const check = createPinnedCheckServerIdentity({ pins: ["00".repeat(32)] });
      const err = check("example.com", { raw: Buffer.from("irrelevant") } as any);
      expect(err).toBe(hostnameError);
      expect((err as any).retriable).toBe(false);
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
      expect(err?.cause?.code).toBe("ERR_CERT_PINNING");
      const message =
        typeof err?.cause?.message === "string"
          ? err.cause.message
          : typeof err?.message === "string"
            ? err.message
            : String(err);
      expect(message).toContain("Certificate pinning failed");
    });
  });

  it("still enforces hostname verification when pinning is enabled", async () => {
    const dir = fixturesDir();
    const certPemLocalhostOnly = fs.readFileSync(path.join(dir, "localhost-dns-only.crt"));
    const keyPemLocalhostOnly = fs.readFileSync(path.join(dir, "localhost-dns-only.key"));

    let requestCount = 0;
    const server = https.createServer({ key: keyPemLocalhostOnly, cert: certPemLocalhostOnly }, (_req, res) => {
      requestCount += 1;
      res.writeHead(200, { "content-type": "text/plain" });
      res.end("ok");
    });

    await new Promise<void>((resolve) => {
      server.listen(0, "127.0.0.1", () => resolve());
    });
    const address = server.address();
    if (!address || typeof address === "string") throw new Error("expected https server to listen on tcp");

    const url = `https://127.0.0.1:${address.port}/`;
    const x509 = new crypto.X509Certificate(certPemLocalhostOnly);
    const pin = sha256FingerprintHexFromCertRaw(x509.raw);

    try {
      const res = fetchWithOrgTls(
        url,
        { method: "GET" },
        {
          tls: {
            certificatePinningEnabled: true,
            certificatePins: [pin]
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
        expect(message).toContain("Hostname/IP does not match certificate");
      });
      expect(requestCount).toBe(0);
    } finally {
      await new Promise<void>((resolve, reject) => {
        server.close((closeErr) => (closeErr ? reject(closeErr) : resolve()));
      });
    }
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

  it("enforces TLSv1.3 minimum by rejecting TLSv1.2-only servers", async () => {
    const dir = fixturesDir();
    const keyPem = fs.readFileSync(path.join(dir, "localhost.key"));

    const server = https.createServer({ key: keyPem, cert: certPem, maxVersion: "TLSv1.2" }, (_req, res) => {
      res.writeHead(200, { "content-type": "text/plain" });
      res.end("ok");
    });

    await new Promise<void>((resolve) => {
      server.listen(0, "127.0.0.1", () => resolve());
    });
    const address = server.address();
    if (!address || typeof address === "string") throw new Error("expected https server to listen on tcp");

    const url = `https://127.0.0.1:${address.port}/`;
    try {
      await expect(
        fetchWithOrgTls(url, { method: "GET" }, { tls: { certificatePinningEnabled: false, certificatePins: [], ca: certPem } })
      ).rejects.toThrow();
    } finally {
      await new Promise<void>((resolve, reject) => {
        server.close((closeErr) => (closeErr ? reject(closeErr) : resolve()));
      });
    }
  });
});
