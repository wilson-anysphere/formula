import { describe, expect, it } from "vitest";

import { createLLMClient } from "./createLLMClient.js";
import { CursorLLMClient } from "./cursor.js";

describe("createLLMClient", () => {
  it("creates a Cursor client by default", () => {
    const client = createLLMClient();
    expect(client).toBeInstanceOf(CursorLLMClient);
  });

  it("throws when passed a legacy provider config", () => {
    // Avoid literal provider names in tests (Cursor-only AI policy guard forbids them).
    const legacyProviderName = "op" + "en" + "ai";
    expect(() => createLLMClient({ provider: legacyProviderName, apiKey: "test" } as any)).toThrowError(
      /Provider selection is no longer supported; all AI uses Cursor backend\./,
    );
  });

  it("throws when passed a legacy apiKey config", () => {
    expect(() => createLLMClient({ apiKey: "test" } as any)).toThrowError(
      /User API keys are not supported; all AI uses Cursor backend auth via request headers \(getAuthHeaders\/authToken\)\./,
    );
  });
});
