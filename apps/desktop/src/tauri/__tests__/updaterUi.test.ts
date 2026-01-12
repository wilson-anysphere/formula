/**
 * @vitest-environment jsdom
 */

import { afterEach, describe, expect, it, vi } from "vitest";

import * as ui from "../../extensions/ui";
import { handleUpdaterEvent } from "../updaterUi";

describe("updaterUi", () => {
  afterEach(() => {
    try {
      vi.runOnlyPendingTimers();
    } catch {
      // Timers weren't mocked.
    }
    vi.useRealTimers();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    document.body.replaceChildren();
  });

  it("shows + focuses the main window before rendering manual-check feedback", async () => {
    vi.useFakeTimers();
    document.body.innerHTML = '<div id="toast-root"></div>';

    const show = vi.fn(async () => {});
    const setFocus = vi.fn(async () => {});
    const handle = { show, setFocus };

    vi.stubGlobal("__TAURI__", {
      window: {
        getCurrentWindow: () => handle,
      },
    });

    const toastSpy = vi.spyOn(ui, "showToast");

    await handleUpdaterEvent("update-not-available", { source: "manual" });

    expect(show).toHaveBeenCalledTimes(1);
    expect(setFocus).toHaveBeenCalledTimes(1);
    expect(toastSpy).toHaveBeenCalledTimes(1);

    const toast = document.querySelector('[data-testid="toast"]');
    expect(toast?.textContent).toContain("up to date");

    expect(show.mock.invocationCallOrder[0]).toBeLessThan(toastSpy.mock.invocationCallOrder[0]);
    expect(setFocus.mock.invocationCallOrder[0]).toBeLessThan(toastSpy.mock.invocationCallOrder[0]);
  });
});
