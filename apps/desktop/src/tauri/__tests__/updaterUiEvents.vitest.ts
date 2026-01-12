/**
 * @vitest-environment jsdom
 */

import { afterEach, describe, expect, it, vi } from "vitest";

import * as ui from "../../extensions/ui";
import { handleUpdaterEvent, installUpdaterUi } from "../updaterUi";

describe("updaterUi (events)", () => {
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

  it("renders an 'already checking' toast for repeated manual update checks", async () => {
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

    const handlers = new Map<string, (event: any) => void>();
    const listen = vi.fn(async (eventName: string, handler: (event: any) => void) => {
      handlers.set(eventName, handler);
      return () => {};
    });

    installUpdaterUi(listen);

    expect(listen).toHaveBeenCalledWith("update-check-already-running", expect.any(Function));

    handlers.get("update-check-already-running")?.({ payload: { source: "manual" } });

    // `installUpdaterUi` wires handlers using `void handleUpdaterEvent(...)`; flush a few
    // microtasks so async window show/focus completes before we assert.
    for (let idx = 0; idx < 10; idx++) {
      // eslint-disable-next-line no-await-in-loop
      await Promise.resolve();
    }

    expect(show).toHaveBeenCalledTimes(1);
    expect(setFocus).toHaveBeenCalledTimes(1);
    expect(toastSpy).toHaveBeenCalledTimes(1);

    const toast = document.querySelector('[data-testid="toast"]');
    expect(toast?.textContent).toContain("Already checking");

    expect(show.mock.invocationCallOrder[0]).toBeLessThan(toastSpy.mock.invocationCallOrder[0]);
    expect(setFocus.mock.invocationCallOrder[0]).toBeLessThan(toastSpy.mock.invocationCallOrder[0]);
  });

  it("ignores 'already running' events emitted during startup checks", async () => {
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

    await handleUpdaterEvent("update-check-already-running", { source: "startup" });

    expect(show).not.toHaveBeenCalled();
    expect(setFocus).not.toHaveBeenCalled();
    expect(toastSpy).not.toHaveBeenCalled();
  });
});

