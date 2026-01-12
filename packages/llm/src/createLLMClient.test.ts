import { describe, expect, it } from "vitest";

import { createLLMClient } from "./createLLMClient.js";
import { CursorLLMClient } from "./cursor.js";

describe("createLLMClient", () => {
  it("creates a Cursor client by default", () => {
    const client = createLLMClient();
    expect(client).toBeInstanceOf(CursorLLMClient);
  });

  it("throws when passed a legacy provider config", () => {
    expect(() => createLLMClient({ provider: "openai", apiKey: "test" } as any)).toThrowError(
      /Provider selection is no longer supported; all AI uses Cursor backend\./,
    );
  });
});
