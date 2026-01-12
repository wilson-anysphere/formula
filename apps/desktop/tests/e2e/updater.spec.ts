import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

function installTauriStubForUpdaterTests(): void {
  const listeners: Record<string, any[]> = {};
  (window as any).__tauriListeners = listeners;

  (window as any).__tauriWindowHidden = false;
  (window as any).__tauriWindowShowCalls = 0;
  (window as any).__tauriWindowFocusCalls = 0;

  const winHandle = {
    isVisible: async () => !(window as any).__tauriWindowHidden,
    show: async () => {
      (window as any).__tauriWindowShowCalls += 1;
      (window as any).__tauriWindowHidden = false;
    },
    setFocus: async () => {
      (window as any).__tauriWindowFocusCalls += 1;
    },
    hide: async () => {
      (window as any).__tauriWindowHidden = true;
    },
  };

  (window as any).__TAURI__ = {
    core: {
      invoke: async (_cmd: string, _args: any) => null,
    },
    event: {
      listen: async (name: string, handler: any) => {
        const entry = listeners[name] ?? [];
        entry.push(handler);
        listeners[name] = entry;
        return () => {
          const handlers = listeners[name];
          if (!handlers) return;
          const idx = handlers.indexOf(handler);
          if (idx >= 0) handlers.splice(idx, 1);
          if (handlers.length === 0) delete listeners[name];
        };
      },
      emit: async () => {},
    },
    window: {
      getCurrentWebviewWindow: () => winHandle,
    },
    dialog: {
      open: async () => null,
      save: async () => null,
    },
  };
}

async function fireTauriEvent(page: Page, name: string, payload: unknown): Promise<void> {
  await page.waitForFunction((eventName) => {
    const handlers = (window as any).__tauriListeners?.[String(eventName)];
    if (!handlers) return false;
    if (typeof handlers === "function") return true;
    if (Array.isArray(handlers)) return handlers.length > 0;
    return false;
  }, name);
  await page.evaluate(
    ({ eventName, eventPayload }) => {
      const handlers = (window as any).__tauriListeners?.[String(eventName)];
      if (!handlers) return;
      const event = { payload: eventPayload };
      if (typeof handlers === "function") {
        handlers(event);
        return;
      }
      if (Array.isArray(handlers)) {
        for (const handler of handlers) {
          handler(event);
        }
      }
    },
    { eventName: name, eventPayload: payload },
  );
}

test.describe("updater UI wiring", () => {
  test.beforeEach(async ({ page }) => {
    await page.addInitScript(installTauriStubForUpdaterTests);
  });

  test("manual update check shows toast feedback", async ({ page }) => {
    await gotoDesktop(page);

    await fireTauriEvent(page, "update-check-started", { source: "manual" });
    await expect(page.getByTestId("toast").filter({ hasText: /checking for updates/i })).toBeVisible();

    await fireTauriEvent(page, "update-not-available", { source: "manual" });
    await expect(page.getByTestId("toast").filter({ hasText: /up to date/i })).toBeVisible();
  });

  test("update available shows a modal containing version + release notes", async ({ page }) => {
    await gotoDesktop(page);

    await fireTauriEvent(page, "update-available", { source: "manual", version: "9.9.9", body: "Notes" });

    const dialog = page.getByTestId("updater-dialog");
    await expect(dialog).toBeVisible();
    await expect(page.getByTestId("updater-version")).toContainText("9.9.9");
    await expect(page.getByTestId("updater-body")).toContainText("Notes");
  });

  test("manual update events show + focus the window when hidden-to-tray", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      (window as any).__tauriWindowHidden = true;
    });

    await fireTauriEvent(page, "update-check-started", { source: "manual" });

    await page.waitForFunction(
      () => (window as any).__tauriWindowShowCalls > 0 && (window as any).__tauriWindowFocusCalls > 0,
    );
  });
});
