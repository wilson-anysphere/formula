import { describe, expect, it, vi } from "vitest";

vi.mock("../http/tls", () => {
  return {
    fetchWithOrgTls: vi.fn()
  };
});

import { sendSiemBatch } from "../siem/sender";
import { fetchWithOrgTls } from "../http/tls";

describe("SIEM sender TLS failures", () => {
  it("does not retry on TLS certificate validation errors", async () => {
    const cause = Object.assign(new Error("self-signed certificate"), { code: "DEPTH_ZERO_SELF_SIGNED_CERT" });
    const err: any = new TypeError("fetch failed");
    err.cause = cause;

    (fetchWithOrgTls as unknown as ReturnType<typeof vi.fn>).mockRejectedValue(err);

    const config = {
      endpointUrl: "https://example.invalid/ingest",
      retry: { maxAttempts: 3, baseDelayMs: 1, maxDelayMs: 1, jitter: false }
    };

    const event = {
      id: "11111111-1111-4111-8111-111111111111",
      timestamp: new Date("2025-01-01T00:00:00.000Z").toISOString(),
      orgId: null,
      userId: null,
      userEmail: null,
      eventType: "test.tls_error",
      resourceType: "organization",
      resourceId: null,
      ipAddress: null,
      userAgent: null,
      sessionId: null,
      success: true,
      errorCode: null,
      errorMessage: null,
      details: {}
    };

    await expect(
      sendSiemBatch(config as any, [event] as any, {
        tls: { certificatePinningEnabled: false, certificatePins: [] }
      })
    ).rejects.toMatchObject({ retriable: false });

    expect(fetchWithOrgTls).toHaveBeenCalledTimes(1);
  });
});

