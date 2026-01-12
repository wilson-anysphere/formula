import { describe, expect, it } from "vitest";

import { createLLMClient } from "./createLLMClient.js";
import { CursorLLMClient } from "./cursor.js";

describe("createLLMClient", () => {
  it("creates a Cursor backend client", () => {
    const client = createLLMClient();
    expect(client).toBeInstanceOf(CursorLLMClient);
  });

  it("rejects legacy provider configuration", () => {
    expect(() => createLLMClient({ provider: "openai", apiKey: "test-key" } as any)).toThrow(/no longer accepts/);
  });
});
