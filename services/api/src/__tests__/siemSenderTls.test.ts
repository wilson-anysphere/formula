import { describe, expect, it, vi } from "vitest";

vi.mock("../http/tls", () => {
  return {
    fetchWithOrgTls: vi.fn()
  };
});

import { createAuditEvent } from "@formula/audit-core";
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
});
