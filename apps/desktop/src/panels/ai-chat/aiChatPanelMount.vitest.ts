// @vitest-environment jsdom

import { act } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { createPanelBodyRenderer } from "../panelBodyRenderer.js";
import { PanelIds } from "../panelRegistry.js";
import { DocumentController } from "../../document/documentController.js";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

function setNativeInputValue(input: HTMLInputElement, value: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
  if (!setter) throw new Error("Missing HTMLInputElement.value setter");
  setter.call(input, value);
}

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

  it("can save an API key and transition into the chat UI", async () => {
    clearApiKeyStorage();

    const doc = new DocumentController();
    const getDocumentController = vi.fn(() => doc);

    const renderer = createPanelBodyRenderer({
      getDocumentController,
    });

    const body = document.createElement("div");
    document.body.appendChild(body);

    await act(async () => {
      renderer.renderPanelBody(PanelIds.AI_CHAT, body);
    });

    const keyInput = body.querySelector("input");
    expect(keyInput).toBeInstanceOf(HTMLInputElement);

    await act(async () => {
      setNativeInputValue(keyInput as HTMLInputElement, "sk-test-key");
      keyInput?.dispatchEvent(new Event("input", { bubbles: true }));
      keyInput?.dispatchEvent(new Event("change", { bubbles: true }));
    });

    const saveBtn = Array.from(body.querySelectorAll("button")).find((b) => b.textContent === "Save key");
    expect(saveBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      (saveBtn as HTMLButtonElement).click();
    });

    expect(getDocumentController).toHaveBeenCalled();
    expect(body.textContent).toContain("Chat");

    act(() => {
      renderer.cleanup([]);
    });
  });
});
