// @vitest-environment jsdom

import { act } from "react";
import { afterEach, describe, expect, it } from "vitest";

import { createPanelBodyRenderer } from "../panelBodyRenderer.js";
import { PanelIds } from "../panelRegistry.js";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

function clearApiKeyStorage() {
  // Node 25 ships an experimental `globalThis.localStorage` accessor that throws
  // unless Node is started with `--localstorage-file`. Guard all access so our
  // jsdom-based UI tests don't crash on newer Node versions.
  try {
    globalThis.localStorage?.removeItem("formula:openaiApiKey");
  } catch {
    // ignore
  }
  try {
    globalThis.localStorage?.clear();
  } catch {
    // ignore
  }
}

describe("AI chat panel", () => {
  afterEach(() => {
    document.body.innerHTML = "";
    clearApiKeyStorage();
  });

  it("mounts via renderPanelBody and shows setup state when no API key is set", async () => {
    clearApiKeyStorage();

    const renderer = createPanelBodyRenderer({
      getDocumentController: () => {
        throw new Error("document controller should not be requested when API key is missing");
      },
    });

    const body = document.createElement("div");
    document.body.appendChild(body);

    await act(async () => {
      renderer.renderPanelBody(PanelIds.AI_CHAT, body);
    });

    expect(body.textContent).toContain("AI chat setup");

    act(() => {
      renderer.cleanup([]);
    });
  });
});
