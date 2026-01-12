import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

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
          getCurrentWebviewWindow: () => ({
            show: async () => {
              (window as any).__tauriShowCalls += 1;
            },
            setFocus: async () => {
              (window as any).__tauriFocusCalls += 1;
            },
          }),
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

  test("startup update-available events do not open the in-app dialog and instead request a system notification", async ({
    page,
  }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, Array<(event: any) => void>> = {};
      const emitted: Array<{ event: string; payload: any }> = [];
      const notifications: Array<{ title: string; body?: string }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriEmittedEvents = emitted;
      (window as any).__tauriNotifications = notifications;

      (window as any).__TAURI__ = {
        core: {
          invoke: async (_cmd: string, _args: any) => null,
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
          getCurrentWebviewWindow: () => ({
            show: async () => {},
            setFocus: async () => {},
          }),
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

    await expect(page.getByTestId("updater-dialog")).toHaveCount(0);

    await page.waitForFunction(() => ((window as any).__tauriNotifications?.length ?? 0) >= 1);
    const notifications = await page.evaluate(() => (window as any).__tauriNotifications ?? []);
    expect(Array.isArray(notifications)).toBe(true);
    expect(notifications.length).toBeGreaterThan(0);
    expect(String(notifications[0]?.title ?? "")).toBe("Update available");
    expect(String(notifications[0]?.body ?? "")).toContain("9.9.9");
  });
});
