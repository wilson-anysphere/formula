// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";

import * as nativeDialogs from "./nativeDialogs";

describe("tauri/nativeDialogs", () => {
  const originalTauriDescriptor = Object.getOwnPropertyDescriptor(globalThis, "__TAURI__");

  afterEach(() => {
    if (originalTauriDescriptor) {
      Object.defineProperty(globalThis, "__TAURI__", originalTauriDescriptor);
    } else {
      // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
      delete (globalThis as any).__TAURI__;
    }
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("uses Tauri dialog APIs when __TAURI__.dialog is available", async () => {
    const tauriConfirm = vi.fn(async () => true);
    const tauriMessage = vi.fn(async () => undefined);

    const windowConfirm = vi.fn(() => false);
    const windowAlert = vi.fn(() => undefined);

    vi.stubGlobal("__TAURI__", { dialog: { confirm: tauriConfirm, message: tauriMessage } });
    vi.stubGlobal("window", { confirm: windowConfirm, alert: windowAlert });

    const ok = await nativeDialogs.confirm("Discard?", { title: "Formula", okLabel: "Discard", cancelLabel: "Cancel" });
    expect(ok).toBe(true);
    expect(tauriConfirm).toHaveBeenCalledWith("Discard?", { title: "Formula", okLabel: "Discard", cancelLabel: "Cancel" });
    expect(windowConfirm).not.toHaveBeenCalled();

    await nativeDialogs.alert("Failed to open workbook", { title: "Formula" });
    expect(tauriMessage).toHaveBeenCalledWith("Failed to open workbook", { title: "Formula" });
    expect(windowAlert).not.toHaveBeenCalled();
  });

  it("supports the __TAURI__.plugin.dialog API shape", async () => {
    const tauriConfirm = vi.fn(async () => true);
    const tauriMessage = vi.fn(async () => undefined);

    const windowConfirm = vi.fn(() => false);
    const windowAlert = vi.fn(() => undefined);

    vi.stubGlobal("__TAURI__", { plugin: { dialog: { confirm: tauriConfirm, message: tauriMessage } } });
    vi.stubGlobal("window", { confirm: windowConfirm, alert: windowAlert });

    const ok = await nativeDialogs.confirm("Discard?", { title: "Formula" });
    expect(ok).toBe(true);
    expect(tauriConfirm).toHaveBeenCalledWith("Discard?", { title: "Formula" });
    expect(windowConfirm).not.toHaveBeenCalled();

    await nativeDialogs.alert("Failed to open workbook", { title: "Formula" });
    expect(tauriMessage).toHaveBeenCalledWith("Failed to open workbook", { title: "Formula" });
    expect(windowAlert).not.toHaveBeenCalled();
  });

  it("supports the __TAURI__.plugins.dialog API shape", async () => {
    const tauriConfirm = vi.fn(async () => true);
    const tauriMessage = vi.fn(async () => undefined);

    const windowConfirm = vi.fn(() => false);
    const windowAlert = vi.fn(() => undefined);

    vi.stubGlobal("__TAURI__", { plugins: { dialog: { confirm: tauriConfirm, message: tauriMessage } } });
    vi.stubGlobal("window", { confirm: windowConfirm, alert: windowAlert });

    const ok = await nativeDialogs.confirm("Discard?", { title: "Formula" });
    expect(ok).toBe(true);
    expect(tauriConfirm).toHaveBeenCalledWith("Discard?", { title: "Formula" });
    expect(windowConfirm).not.toHaveBeenCalled();

    await nativeDialogs.alert("Failed to open workbook", { title: "Formula" });
    expect(tauriMessage).toHaveBeenCalledWith("Failed to open workbook", { title: "Formula" });
    expect(windowAlert).not.toHaveBeenCalled();
  });

  it("treats a throwing __TAURI__ getter as an unavailable API and falls back safely", async () => {
    const windowConfirm = vi.fn(() => true);
    const windowAlert = vi.fn(() => undefined);

    Object.defineProperty(globalThis, "__TAURI__", {
      configurable: true,
      get() {
        throw new Error("Blocked __TAURI__");
      },
    });
    vi.stubGlobal("window", { confirm: windowConfirm, alert: windowAlert });

    const ok = await nativeDialogs.confirm("Discard?");
    expect(ok).toBe(true);
    expect(windowConfirm).toHaveBeenCalledWith("Discard?");

    await nativeDialogs.alert("Something went wrong");
    expect(windowAlert).toHaveBeenCalledWith("Something went wrong");
  });

  it("falls back to window.confirm/alert when Tauri dialog APIs are unavailable", async () => {
    const windowConfirm = vi.fn(() => true);
    const windowAlert = vi.fn(() => undefined);

    vi.stubGlobal("__TAURI__", undefined);
    vi.stubGlobal("window", { confirm: windowConfirm, alert: windowAlert });

    const ok = await nativeDialogs.confirm("Discard?");
    expect(ok).toBe(true);
    expect(windowConfirm).toHaveBeenCalledWith("Discard?");

    await nativeDialogs.alert("Something went wrong");
    expect(windowAlert).toHaveBeenCalledWith("Something went wrong");
  });

  it("treats throwing window.confirm/alert as unavailable APIs", async () => {
    const windowConfirm = vi.fn(() => {
      throw new Error("Not implemented: window.confirm");
    });
    const windowAlert = vi.fn(() => {
      throw new Error("Not implemented: window.alert");
    });

    vi.stubGlobal("__TAURI__", undefined);
    vi.stubGlobal("window", { confirm: windowConfirm, alert: windowAlert });

    await expect(nativeDialogs.confirm("Discard?")).resolves.toBe(false);
    await expect(nativeDialogs.confirm("Discard?", { fallbackValue: true })).resolves.toBe(true);
    await expect(nativeDialogs.alert("Something went wrong")).resolves.toBeUndefined();
  });

  it("avoids calling browser-native window.confirm/alert and falls back to a non-blocking <dialog>", async () => {
    const originalToString = Function.prototype.toString;

    const confirmSpy = vi.fn(() => true);
    const alertSpy = vi.fn(() => undefined);

    // Make the stubs appear "native" so nativeDialogs skips them.
    vi.spyOn(Function.prototype, "toString").mockImplementation(function (this: unknown): string {
      if (this === confirmSpy) return "function confirm() { [native code] }";
      if (this === alertSpy) return "function alert() { [native code] }";
      return originalToString.call(this as any);
    });

    vi.stubGlobal("__TAURI__", undefined);
    vi.stubGlobal("window", { confirm: confirmSpy, alert: alertSpy });

    const confirmPromise = nativeDialogs.confirm("Discard?");
    const confirmDialog = document.querySelector('dialog[data-testid="quick-pick"]') as HTMLDialogElement | null;
    expect(confirmDialog).not.toBeNull();
    const ok = confirmDialog!.querySelector('[data-testid="quick-pick-item-0"]') as HTMLButtonElement | null;
    expect(ok).not.toBeNull();
    ok!.click();
    await expect(confirmPromise).resolves.toBe(true);
    expect(confirmSpy).not.toHaveBeenCalled();

    const alertPromise = nativeDialogs.alert("Something went wrong");
    const alertDialog = document.querySelector('dialog[data-testid="quick-pick"]') as HTMLDialogElement | null;
    expect(alertDialog).not.toBeNull();
    const alertOk = alertDialog!.querySelector('[data-testid="quick-pick-item-0"]') as HTMLButtonElement | null;
    expect(alertOk).not.toBeNull();
    alertOk!.click();
    await expect(alertPromise).resolves.toBeUndefined();
    expect(alertSpy).not.toHaveBeenCalled();
  });
});
