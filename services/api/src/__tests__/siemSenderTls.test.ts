import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../http/tls", () => {
  return {
    fetchWithOrgTls: vi.fn()
  };
});

import { createAuditEvent } from "@formula/audit-core";
import { sendSiemBatch } from "../siem/sender";
import { fetchWithOrgTls } from "../http/tls";

describe("SIEM sender TLS failures", () => {
  beforeEach(() => {
    (fetchWithOrgTls as unknown as ReturnType<typeof vi.fn>).mockReset();
  });

  it("does not retry on TLS certificate validation errors", async () => {
    const cause = Object.assign(new Error("self-signed certificate"), { code: "DEPTH_ZERO_SELF_SIGNED_CERT" });
    const err: any = new TypeError("fetch failed");
    err.cause = cause;

    (fetchWithOrgTls as unknown as ReturnType<typeof vi.fn>).mockRejectedValue(err);

    const config = {
      endpointUrl: "https://example.invalid/ingest",
      format: "json",
      retry: { maxAttempts: 3, baseDelayMs: 1, maxDelayMs: 1, jitter: false }
    };

    const event = createAuditEvent({
      eventType: "test.tls_error",
      actor: { type: "system", id: "unit-test" },
      context: { orgId: null },
      resource: { type: "integration", id: "siem", name: "siem" },
      success: true
    });

    await expect(
      sendSiemBatch(config as any, [event], {
        tls: { certificatePinningEnabled: false, certificatePins: [] }
      })
    ).rejects.toMatchObject({ retriable: false });

    expect(fetchWithOrgTls).toHaveBeenCalledTimes(1);

    const call = (fetchWithOrgTls as unknown as ReturnType<typeof vi.fn>).mock.calls[0];
    expect(call).toBeTruthy();
    const init = call![1] as any;
    expect(Buffer.isBuffer(init.body)).toBe(true);
    expect(init.headers?.["Content-Type"]).toBe("application/json");
  });

  it("marks certificate pinning failures as non-retriable and does not retry", async () => {
    const pinningError = Object.assign(new Error("Certificate pinning failed: server certificate fingerprint mismatch"), {
      retriable: false
    });
    const err: any = new TypeError("fetch failed");
    err.cause = pinningError;

    (fetchWithOrgTls as unknown as ReturnType<typeof vi.fn>).mockRejectedValue(err);

    const config = {
      endpointUrl: "https://example.invalid/ingest",
      format: "json",
      retry: { maxAttempts: 5, baseDelayMs: 1, maxDelayMs: 1, jitter: false }
    };

    const event = createAuditEvent({
      eventType: "test.tls_pinning_error",
      actor: { type: "system", id: "unit-test" },
      context: { orgId: null },
      resource: { type: "integration", id: "siem", name: "siem" },
      success: true
    });

    await expect(
      sendSiemBatch(config as any, [event], {
        tls: { certificatePinningEnabled: true, certificatePins: ["00".repeat(32)] }
      })
    ).rejects.toMatchObject({ retriable: false, siemErrorLabel: "tls_pinning_failed" });

    expect(fetchWithOrgTls).toHaveBeenCalledTimes(1);
  });

  it("rejects http endpoints in production without attempting a request", async () => {
    const previousEnv = process.env.NODE_ENV;
    process.env.NODE_ENV = "production";
    try {
      (fetchWithOrgTls as unknown as ReturnType<typeof vi.fn>).mockClear();

      const config = {
        endpointUrl: "http://example.invalid/ingest",
        format: "json",
        retry: { maxAttempts: 3, baseDelayMs: 1, maxDelayMs: 1, jitter: false }
      };

      const event = createAuditEvent({
        eventType: "test.http_in_production",
        actor: { type: "system", id: "unit-test" },
        context: { orgId: null },
        resource: { type: "integration", id: "siem", name: "siem" },
        success: true
      });

      await expect(
        sendSiemBatch(config as any, [event], {
          tls: { certificatePinningEnabled: false, certificatePins: [] }
        })
      ).rejects.toMatchObject({ retriable: false, siemErrorLabel: "insecure_http_endpoint" });

      expect(fetchWithOrgTls).not.toHaveBeenCalled();
    } finally {
      process.env.NODE_ENV = previousEnv;
    }
  });
});
