import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function flushMicrotasks(page: Page, times = 6): Promise<void> {
  await page.evaluate(async (n) => {
    for (let idx = 0; idx < n; idx += 1) {
      await new Promise<void>((resolve) => queueMicrotask(resolve));
    }
  }, times);
}

async function waitForTauriListeners(page: Page, eventName: string): Promise<void> {
  await page.waitForFunction(
    (name) => {
      const listeners = (window as any).__tauriListeners?.[name];
      return Array.isArray(listeners) && listeners.length > 0;
    },
    eventName,
  );
}

async function dispatchTauriEvent(page: Page, eventName: string, payload: any): Promise<void> {
  await page.evaluate(
    ({ name, payload: eventPayload }) => {
      const listeners = (window as any).__tauriListeners?.[name];
      if (!Array.isArray(listeners) || listeners.length === 0) {
        throw new Error(`Missing Tauri listeners for ${name}`);
      }
      for (const handler of listeners) {
        try {
          handler({ payload: eventPayload });
        } catch (err) {
          console.error(`Failed to deliver Tauri event ${name}:`, err);
        }
      }
    },
    { name: eventName, payload },
  );
}

test.describe("desktop updater UI wiring (tauri)", () => {
  test("manual update events surface toasts, focus the window, and open the updater dialog", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, Array<(event: any) => void>> = {};
      const emitted: Array<{ event: string; payload: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriEmittedEvents = emitted;
      (window as any).__tauriShowCalls = 0;
      (window as any).__tauriFocusCalls = 0;

      const windowHandle = {
        show: async () => {
          (window as any).__tauriShowCalls += 1;
        },
        setFocus: async () => {
          (window as any).__tauriFocusCalls += 1;
        },
      };

      (window as any).__TAURI__ = {
        core: {
          invoke: async (_cmd: string, _args: any) => {
            // Keep the invoke surface flexible; updater UI wiring should not depend on
            // unrelated backend commands.
            return null;
          },
        },
        event: {
          listen: async (name: string, handler: any) => {
            if (!Array.isArray(listeners[name])) {
              listeners[name] = [];
            }
            listeners[name].push(handler);
            return () => {
              const arr = listeners[name];
              if (!Array.isArray(arr)) return;
              const idx = arr.indexOf(handler);
              if (idx >= 0) arr.splice(idx, 1);
            };
          },
          emit: async (event: string, payload?: any) => {
            emitted.push({ event, payload: payload ?? null });
          },
        },
        window: {
          // Provide all common handle accessors used by our Tauri abstractions so this
          // test stays resilient to future refactors.
          getCurrentWebviewWindow: () => windowHandle,
          getCurrentWindow: () => windowHandle,
          getCurrent: () => windowHandle,
          appWindow: windowHandle,
        },
        notification: {
          notify: async (_payload: { title: string; body?: string }) => {},
        },
      };
    });

    await gotoDesktop(page);

    // Ensure the frontend handshake fired (used by the Rust side to flush queued updater events).
    await page.waitForFunction(() =>
      Boolean((window as any).__tauriEmittedEvents?.some((entry: any) => entry?.event === "updater-ui-ready")),
    );

    await waitForTauriListeners(page, "update-check-started");
    await dispatchTauriEvent(page, "update-check-started", { source: "manual" });

    await expect(page.locator("#toast-root")).toContainText("Checking for updates…");
    await page.waitForFunction(() => (window as any).__tauriShowCalls >= 1);
    await page.waitForFunction(() => (window as any).__tauriFocusCalls >= 1);

    await waitForTauriListeners(page, "update-check-error");
    await dispatchTauriEvent(page, "update-check-error", { source: "manual", error: "network down" });
    await expect(page.locator("#toast-root")).toContainText("Error: network down");

    await waitForTauriListeners(page, "update-check-already-running");
    await dispatchTauriEvent(page, "update-check-already-running", { source: "manual" });
    await expect(page.locator("#toast-root")).toContainText("Already checking for updates…");

    await waitForTauriListeners(page, "update-available");
    await dispatchTauriEvent(page, "update-available", {
      source: "manual",
      version: "9.9.9",
      body: "Release notes",
    });

    const dialog = page.getByTestId("updater-dialog");
    await expect(dialog).toBeVisible();
    await expect(dialog).toHaveJSProperty("open", true);
    await expect(page.getByTestId("updater-version")).toContainText("Version 9.9.9");
    await expect(page.getByTestId("updater-body")).toContainText("Release notes");

    await waitForTauriListeners(page, "update-not-available");
    await dispatchTauriEvent(page, "update-not-available", { source: "manual" });
    await expect(page.locator("#toast-root")).toContainText("You're up to date.");
  });

  test("startup update completion is treated as manual when a manual check was queued behind an in-flight startup check", async ({
    page,
  }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, Array<(event: any) => void>> = {};
      const emitted: Array<{ event: string; payload: any }> = [];
      const notifications: Array<{ title: string; body?: string }> = [];
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriEmittedEvents = emitted;
      (window as any).__tauriNotifications = notifications;
      (window as any).__tauriInvokes = invokes;
      (window as any).__tauriShowCalls = 0;
      (window as any).__tauriFocusCalls = 0;

      const windowHandle = {
        show: async () => {
          (window as any).__tauriShowCalls += 1;
        },
        setFocus: async () => {
          (window as any).__tauriFocusCalls += 1;
        },
      };

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });
            return null;
          },
        },
        event: {
          listen: async (name: string, handler: any) => {
            if (!Array.isArray(listeners[name])) {
              listeners[name] = [];
            }
            listeners[name].push(handler);
            return () => {
              const arr = listeners[name];
              if (!Array.isArray(arr)) return;
              const idx = arr.indexOf(handler);
              if (idx >= 0) arr.splice(idx, 1);
            };
          },
          emit: async (event: string, payload?: any) => {
            emitted.push({ event, payload: payload ?? null });
          },
        },
        window: {
          getCurrentWebviewWindow: () => windowHandle,
          getCurrentWindow: () => windowHandle,
          getCurrent: () => windowHandle,
          appWindow: windowHandle,
        },
        notification: {
          notify: async (payload: { title: string; body?: string }) => {
            notifications.push({ title: payload.title, body: payload.body });
          },
        },
      };
    });

    await gotoDesktop(page);

    await page.waitForFunction(() =>
      Boolean((window as any).__tauriEmittedEvents?.some((entry: any) => entry?.event === "updater-ui-ready")),
    );

    // User clicks "Check for Updates" while a startup check is already in flight.
    await waitForTauriListeners(page, "update-check-already-running");
    await dispatchTauriEvent(page, "update-check-already-running", { source: "manual" });
    await expect(page.locator("#toast-root")).toContainText("Already checking for updates…");
    await page.waitForFunction(() => (window as any).__tauriShowCalls >= 1);
    await page.waitForFunction(() => (window as any).__tauriFocusCalls >= 1);

    // The backend eventually reports the completion with `source: "startup"`, but the
    // UI should still surface the dialog (treat as a manual follow-up) and should not
    // request a system notification.
    await waitForTauriListeners(page, "update-available");
    await dispatchTauriEvent(page, "update-available", { source: "startup", version: "9.9.9", body: "Notes" });

    const dialog = page.getByTestId("updater-dialog");
    await expect(dialog).toBeVisible();
    await expect(page.getByTestId("updater-version")).toContainText("Version 9.9.9");

    // Should not spam `show()`/`setFocus()` twice (startup completion shouldn't re-focus).
    const counts = await page.evaluate(() => ({
      show: (window as any).__tauriShowCalls,
      focus: (window as any).__tauriFocusCalls,
      notifications: (window as any).__tauriNotifications?.length ?? 0,
      notifiedViaInvoke: ((window as any).__tauriInvokes ?? []).some((entry: any) => entry?.cmd === "show_system_notification"),
    }));
    expect(counts.show).toBe(1);
    expect(counts.focus).toBe(1);
    expect(counts.notifications).toBe(0);
    expect(counts.notifiedViaInvoke).toBe(false);
  });

  test("startup update notifications are suppressed after the user dismisses the same version", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, Array<(event: any) => void>> = {};
      const emitted: Array<{ event: string; payload: any }> = [];
      const notifications: Array<{ title: string; body?: string }> = [];
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriEmittedEvents = emitted;
      (window as any).__tauriNotifications = notifications;
      (window as any).__tauriInvokes = invokes;
      (window as any).__tauriShowCalls = 0;
      (window as any).__tauriFocusCalls = 0;

      const windowHandle = {
        show: async () => {
          (window as any).__tauriShowCalls += 1;
        },
        setFocus: async () => {
          (window as any).__tauriFocusCalls += 1;
        },
      };

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });
            return null;
          },
        },
        event: {
          listen: async (name: string, handler: any) => {
            if (!Array.isArray(listeners[name])) {
              listeners[name] = [];
            }
            listeners[name].push(handler);
            return () => {
              const arr = listeners[name];
              if (!Array.isArray(arr)) return;
              const idx = arr.indexOf(handler);
              if (idx >= 0) arr.splice(idx, 1);
            };
          },
          emit: async (event: string, payload?: any) => {
            emitted.push({ event, payload: payload ?? null });
          },
        },
        window: {
          getCurrentWebviewWindow: () => windowHandle,
          getCurrentWindow: () => windowHandle,
          getCurrent: () => windowHandle,
          appWindow: windowHandle,
        },
        notification: {
          notify: async (payload: { title: string; body?: string }) => {
            notifications.push({ title: payload.title, body: payload.body });
          },
        },
      };
    });

    await gotoDesktop(page);

    await page.waitForFunction(() =>
      Boolean((window as any).__tauriEmittedEvents?.some((entry: any) => entry?.event === "updater-ui-ready")),
    );

    // Trigger a manual update dialog and dismiss it (which should persist suppression).
    await waitForTauriListeners(page, "update-available");
    await dispatchTauriEvent(page, "update-available", { source: "manual", version: "9.9.9", body: "Notes" });

    const dialog = page.getByTestId("updater-dialog");
    await expect(dialog).toBeVisible();
    await expect(page.getByTestId("updater-later")).toBeVisible();
    await page.getByTestId("updater-later").click();
    await expect(dialog).toBeHidden();

    const dismissal = await page.evaluate(() => ({
      version: localStorage.getItem("formula.updater.dismissedVersion"),
      dismissedAt: localStorage.getItem("formula.updater.dismissedAt"),
    }));
    expect(dismissal.version).toBe("9.9.9");
    expect(Number(dismissal.dismissedAt)).toBeGreaterThan(0);

    // Clear any earlier notification side-effects from test setup.
    await page.evaluate(() => {
      const notifications = (window as any).__tauriNotifications;
      if (Array.isArray(notifications)) notifications.length = 0;
      const invokes = (window as any).__tauriInvokes;
      if (Array.isArray(invokes)) invokes.length = 0;
    });

    // Now deliver the same version as a startup update. The UX should not open the dialog
    // and should not request a system notification because the user dismissed it recently.
    await dispatchTauriEvent(page, "update-available", { source: "startup", version: "9.9.9", body: "Notes" });
    await flushMicrotasks(page);

    await expect(dialog).toBeHidden();

    const notificationStatus = await page.evaluate(() => {
      const notifications = (window as any).__tauriNotifications ?? [];
      const invokes = (window as any).__tauriInvokes ?? [];
      const invoked = Array.isArray(invokes) && invokes.some((entry: any) => entry?.cmd === "show_system_notification");
      return { notificationsCount: Array.isArray(notifications) ? notifications.length : 0, invoked };
    });
    expect(notificationStatus.notificationsCount).toBe(0);
    expect(notificationStatus.invoked).toBe(false);
  });

  test("startup update-available events do not open the in-app dialog and instead request a system notification", async ({
    page,
  }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, Array<(event: any) => void>> = {};
      const emitted: Array<{ event: string; payload: any }> = [];
      const notifications: Array<{ title: string; body?: string }> = [];
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriEmittedEvents = emitted;
      (window as any).__tauriNotifications = notifications;
      (window as any).__tauriInvokes = invokes;

      const windowHandle = {
        show: async () => {},
        setFocus: async () => {},
      };

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });
            return null;
          },
        },
        event: {
          listen: async (name: string, handler: any) => {
            if (!Array.isArray(listeners[name])) {
              listeners[name] = [];
            }
            listeners[name].push(handler);
            return () => {
              const arr = listeners[name];
              if (!Array.isArray(arr)) return;
              const idx = arr.indexOf(handler);
              if (idx >= 0) arr.splice(idx, 1);
            };
          },
          emit: async (event: string, payload?: any) => {
            emitted.push({ event, payload: payload ?? null });
          },
        },
        window: {
          getCurrentWebviewWindow: () => windowHandle,
          getCurrentWindow: () => windowHandle,
          getCurrent: () => windowHandle,
          appWindow: windowHandle,
        },
        notification: {
          notify: async (payload: { title: string; body?: string }) => {
            notifications.push({ title: payload.title, body: payload.body });
          },
        },
      };
    });

    await gotoDesktop(page);

    await page.waitForFunction(() =>
      Boolean((window as any).__tauriEmittedEvents?.some((entry: any) => entry?.event === "updater-ui-ready")),
    );

    await waitForTauriListeners(page, "update-available");
    await dispatchTauriEvent(page, "update-available", {
      source: "startup",
      version: "9.9.9",
      body: "Notes",
    });

    // Startup checks should remain silent in-app (no modal dialog). `toBeHidden()` passes
    // for both "not present" and "present but not visible" cases.
    await expect(page.getByTestId("updater-dialog")).toBeHidden();

    // A system notification should be requested either through the direct notification
    // plugin API or via the `show_system_notification` invoke command.
    await page.waitForFunction(() => {
      const notifications = (window as any).__tauriNotifications;
      if (Array.isArray(notifications) && notifications.length > 0) return true;
      const invokes = (window as any).__tauriInvokes;
      if (Array.isArray(invokes) && invokes.some((entry: any) => entry?.cmd === "show_system_notification")) return true;
      return false;
    });

    const details = await page.evaluate(() => {
      const notifications = (window as any).__tauriNotifications ?? [];
      const invokes = (window as any).__tauriInvokes ?? [];
      return { notifications, invokes };
    });

    if (Array.isArray(details.notifications) && details.notifications.length > 0) {
      expect(String(details.notifications[0]?.title ?? "")).toContain("Update available");
      expect(String(details.notifications[0]?.body ?? "")).toContain("9.9.9");
    } else {
      const invoke = (details.invokes as any[]).find((entry: any) => entry?.cmd === "show_system_notification");
      expect(invoke).toBeTruthy();
      expect(String(invoke?.args?.title ?? "")).toContain("Update available");
      expect(String(invoke?.args?.body ?? "")).toContain("9.9.9");
    }
  });
});
