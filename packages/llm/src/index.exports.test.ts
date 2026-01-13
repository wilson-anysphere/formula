import { describe, expect, it } from "vitest";

import { serializeToolResultForModel } from "./index.js";

describe("llm index exports", () => {
  it("re-exports serializeToolResultForModel", () => {
    expect(typeof serializeToolResultForModel).toBe("function");
  });
});

