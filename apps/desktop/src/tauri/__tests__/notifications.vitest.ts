import { afterEach, describe, expect, it, vi } from "vitest";

import { notify } from "../notifications";

describe("tauri/notifications", () => {
  const originalTauri = (globalThis as any).__TAURI__;
  const originalNotification = (globalThis as any).Notification;

  afterEach(() => {
    (globalThis as any).__TAURI__ = originalTauri;
    (globalThis as any).Notification = originalNotification;
    vi.restoreAllMocks();
  });

  it("calls the Tauri notification API when __TAURI__ notification is available", async () => {
    const tauriNotify = vi.fn();
    (globalThis as any).__TAURI__ = { notification: { notify: tauriNotify } };

    await notify({ title: "Hello", body: "World" });

    expect(tauriNotify).toHaveBeenCalledTimes(1);
    expect(tauriNotify).toHaveBeenCalledWith({ title: "Hello", body: "World" });
  });

  it("falls back to invoke(show_system_notification) when running in Tauri without a direct notification API", async () => {
    const invoke = vi.fn().mockResolvedValue(null);
    (globalThis as any).__TAURI__ = { core: { invoke } };

    await notify({ title: "Hello", body: "World" });

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("show_system_notification", { title: "Hello", body: "World" });
  });

  it("falls back to the Web Notification API when running in the web target", async () => {
    (globalThis as any).__TAURI__ = undefined;

    const created: Array<{ title: string; options?: NotificationOptions }> = [];
    class MockNotification {
      static permission = "granted";

      constructor(title: string, options?: NotificationOptions) {
        created.push({ title, options });
      }
    }

    (globalThis as any).Notification = MockNotification;

    await notify({ title: "Hi", body: "There" });

    expect(created).toEqual([{ title: "Hi", options: { body: "There" } }]);
  });
});
