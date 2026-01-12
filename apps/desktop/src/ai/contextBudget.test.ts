import { describe, expect, it } from "vitest";

import { getModeContextWindowTokens, getModelContextWindowTokens } from "./contextBudget.js";

describe("contextBudget", () => {
  it("uses a large default context window for Cursor-managed models", () => {
    expect(getModelContextWindowTokens("cursor")).toBe(128_000);
    expect(getModelContextWindowTokens("Cursor-Default")).toBe(128_000);
  });

  it("parses explicit context window hints from the model name", () => {
    expect(getModelContextWindowTokens("model-32k")).toBe(32_000);
    expect(getModelContextWindowTokens("MODEL-200K-v2")).toBe(200_000);
  });

  it("falls back to a conservative default for unknown models", () => {
    expect(getModelContextWindowTokens("unit-test-model")).toBe(16_000);
  });

  it("caps inline_edit context windows to keep prompts small", () => {
    expect(getModeContextWindowTokens("inline_edit", "cursor")).toBe(4_096);
    expect(getModeContextWindowTokens("inline_edit", "model-32k")).toBe(4_096);
    expect(getModeContextWindowTokens("chat", "model-32k")).toBe(32_000);
  });
});

