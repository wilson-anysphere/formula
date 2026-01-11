import crypto from "node:crypto";
import { afterEach, beforeEach, expect, test } from "vitest";

import { MarketplaceClient } from "../src/marketplace/MarketplaceClient";

const originalFetch = globalThis.fetch;
const originalCrypto = (globalThis as any).crypto;

beforeEach(() => {
  if (!globalThis.crypto?.subtle) {
    try {
      // eslint-disable-next-line no-global-assign
      (globalThis as any).crypto = crypto.webcrypto;
    } catch {
      // Some Node versions expose `crypto` as a read-only getter.
    }
  }
});

afterEach(() => {
  // eslint-disable-next-line no-global-assign
  globalThis.fetch = originalFetch;
  try {
    // eslint-disable-next-line no-global-assign
    (globalThis as any).crypto = originalCrypto;
  } catch {
    // ignore
  }
});

function sha256Hex(bytes: Uint8Array): string {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

test("MarketplaceClient.downloadPackage verifies x-package-sha256", async () => {
  const payload = new Uint8Array([1, 2, 3, 4]);
  const expectedSha = sha256Hex(payload);

  // eslint-disable-next-line no-global-assign
  globalThis.fetch = async () =>
    ({
      ok: true,
      status: 200,
      headers: new Headers({
        "x-package-sha256": expectedSha,
        "x-package-format-version": "2",
        "x-publisher": "test"
      }),
      arrayBuffer: async () => payload.buffer.slice(payload.byteOffset, payload.byteOffset + payload.byteLength)
    }) as any;

  const client = new MarketplaceClient({ baseUrl: "/api" });
  const res = await client.downloadPackage("test.ext", "1.0.0");
  expect(res?.sha256).toBe(expectedSha);
});

test("MarketplaceClient.downloadPackage rejects sha mismatch", async () => {
  const payload = new Uint8Array([5, 6, 7]);

  // eslint-disable-next-line no-global-assign
  globalThis.fetch = async () =>
    ({
      ok: true,
      status: 200,
      headers: new Headers({
        "x-package-sha256": "0".repeat(64),
        "x-package-format-version": "2",
        "x-publisher": "test"
      }),
      arrayBuffer: async () => payload.buffer.slice(payload.byteOffset, payload.byteOffset + payload.byteLength)
    }) as any;

  const client = new MarketplaceClient({ baseUrl: "/api" });
  await expect(client.downloadPackage("test.ext", "1.0.0")).rejects.toThrow(/sha256 mismatch/i);
});

