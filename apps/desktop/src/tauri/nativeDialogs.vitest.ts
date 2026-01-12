import { afterEach, describe, expect, it, vi } from "vitest";

import * as nativeDialogs from "./nativeDialogs";

describe("tauri/nativeDialogs", () => {
  afterEach(() => {
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
});

