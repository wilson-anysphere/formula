// @vitest-environment jsdom

import { afterEach, describe, expect, it } from "vitest";

import { purgeLegacyDesktopLLMSettings } from "./desktopLLMClient.js";

describe("purgeLegacyDesktopLLMSettings", () => {
  afterEach(() => {
    try {
      window.localStorage.clear();
    } catch {
      // ignore
    }
  });

  it("removes legacy LLM provider + API key settings from localStorage", () => {
    window.localStorage.setItem("formula:openaiApiKey", "sk-legacy-test");
    window.localStorage.setItem("formula:llm:provider", "openai");

    purgeLegacyDesktopLLMSettings();

    expect(window.localStorage.getItem("formula:openaiApiKey")).toBeNull();
    expect(window.localStorage.getItem("formula:llm:provider")).toBeNull();
  });
});

