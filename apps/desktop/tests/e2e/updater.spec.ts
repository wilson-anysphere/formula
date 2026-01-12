import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

function installTauriStubForUpdaterTests(): void {
  const listeners: Record<string, any[]> = {};
  (window as any).__tauriListeners = listeners;

  (window as any).__tauriInvokeCalls = [];

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
      invoke: async (cmd: string, args: any) => {
        const calls = (window as any).__tauriInvokeCalls;
        if (Array.isArray(calls)) {
          calls.push({ cmd, args });
        } else {
          (window as any).__tauriInvokeCalls = [{ cmd, args }];
        }
        return null;
      },
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

    await page.evaluate(() => {
      (window as any).__tauriInvokeCalls = [];
    });
    await page.getByTestId("updater-view-versions").click();
    await page.waitForFunction(() => {
      const calls = (window as any).__tauriInvokeCalls;
      return Array.isArray(calls) && calls.some((call) => call && call.cmd === "open_external_url");
    });
    const openExternalCalls = await page.evaluate(() => {
      const calls = (window as any).__tauriInvokeCalls;
      return Array.isArray(calls) ? calls.filter((call) => call?.cmd === "open_external_url") : [];
    });
    expect(openExternalCalls.at(-1)?.args?.url).toBe("https://github.com/wilson-anysphere/formula/releases");
    await expect(dialog).toBeHidden();
  });

  test("startup updater events are surfaced after a manual 'already running' follow-up", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      (window as any).__tauriWindowHidden = true;
    });

    await fireTauriEvent(page, "update-check-already-running", { source: "manual" });
    await expect(page.getByTestId("toast").filter({ hasText: /already checking/i })).toBeVisible();

    await page.waitForFunction(
      () => (window as any).__tauriWindowShowCalls > 0 && (window as any).__tauriWindowFocusCalls > 0,
    );

    const initialShowFocusCounts = await page.evaluate(() => {
      return {
        show: (window as any).__tauriWindowShowCalls,
        focus: (window as any).__tauriWindowFocusCalls,
      };
    });

    await fireTauriEvent(page, "update-available", { source: "startup", version: "9.9.9", body: "Notes" });
    await expect(page.getByTestId("updater-dialog")).toBeVisible();

    const finalShowFocusCounts = await page.evaluate(() => {
      return {
        show: (window as any).__tauriWindowShowCalls,
        focus: (window as any).__tauriWindowFocusCalls,
      };
    });

    expect(finalShowFocusCounts).toEqual(initialShowFocusCounts);
  });

  test("update-downloaded shows a ready-to-restart toast when dialog is closed", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      (window as any).__tauriInvokeCalls = [];
    });

    await fireTauriEvent(page, "update-downloaded", { source: "startup", version: "9.9.9" });
    await expect(page.getByTestId("update-ready-toast")).toBeVisible();
    await expect(page.getByTestId("update-ready-toast")).toContainText(/download/i);

    await page.getByTestId("update-ready-view-versions").click();
    await expect(page.getByTestId("update-ready-toast")).toHaveCount(0);
    await page.waitForFunction(() => {
      const calls = (window as any).__tauriInvokeCalls;
      return Array.isArray(calls) && calls.some((call) => call && call.cmd === "open_external_url");
    });
    const openExternalCalls = await page.evaluate(() => {
      const calls = (window as any).__tauriInvokeCalls;
      return Array.isArray(calls) ? calls.filter((call) => call?.cmd === "open_external_url") : [];
    });
    expect(openExternalCalls.at(-1)?.args?.url).toBe("https://github.com/wilson-anysphere/formula/releases");

    await fireTauriEvent(page, "update-downloaded", { source: "startup", version: "9.9.8" });
    await expect(page.getByTestId("update-ready-toast")).toBeVisible();
    await page.getByTestId("update-ready-later").click();
    await expect(page.getByTestId("update-ready-toast")).toHaveCount(0);
    const dismissal = await page.evaluate(() => {
      return {
        version: localStorage.getItem("formula.updater.dismissedVersion"),
        dismissedAt: localStorage.getItem("formula.updater.dismissedAt"),
      };
    });
    expect(dismissal.version).toBe("9.9.8");
    expect(Number(dismissal.dismissedAt)).toBeGreaterThan(0);

    // Restart should run the install + restart commands (and close the toast) without
    // requiring the dialog to be open.
    await fireTauriEvent(page, "update-downloaded", { source: "startup", version: "9.9.7" });
    await expect(page.getByTestId("update-ready-toast")).toBeVisible();

    await page.evaluate(() => {
      (window as any).__tauriInvokeCalls = [];
    });

    await page.getByTestId("update-ready-restart").click();
    await expect(page.getByTestId("update-ready-toast")).toHaveCount(0);

    await page.waitForFunction(() => {
      const calls = (window as any).__tauriInvokeCalls;
      if (!Array.isArray(calls)) return false;
      const cmds = calls.map((call) => call?.cmd);
      return cmds.includes("install_downloaded_update") && (cmds.includes("restart_app") || cmds.includes("quit_app"));
    });

    const restartCalls = await page.evaluate(() => {
      const calls = (window as any).__tauriInvokeCalls;
      return Array.isArray(calls) ? calls : [];
    });
    const installIdx = restartCalls.findIndex((call) => call?.cmd === "install_downloaded_update");
    const restartIdx = restartCalls.findIndex((call) => call?.cmd === "restart_app" || call?.cmd === "quit_app");
    expect(installIdx).toBeGreaterThanOrEqual(0);
    expect(restartIdx).toBeGreaterThan(installIdx);
  });

  test("download error events update the dialog when open", async ({ page }) => {
    await gotoDesktop(page);

    await fireTauriEvent(page, "update-available", { source: "manual", version: "9.9.9", body: "Notes" });
    await expect(page.getByTestId("updater-dialog")).toBeVisible();

    await fireTauriEvent(page, "update-download-started", { source: "startup", version: "9.9.9" });
    await expect(page.getByTestId("updater-progress-wrap")).toBeVisible();

    await fireTauriEvent(page, "update-download-error", {
      source: "startup",
      version: "9.9.9",
      message: "network down",
    });
    await expect(page.getByTestId("updater-progress-wrap")).toBeHidden();
    await expect(page.getByTestId("updater-status")).toContainText("network down");
    await expect(page.getByTestId("updater-view-versions")).toHaveClass(/updater-dialog__view-versions--primary/);
  });

  test("download progress events update the dialog and show a restart CTA when open", async ({ page }) => {
    await gotoDesktop(page);

    await fireTauriEvent(page, "update-available", { source: "manual", version: "9.9.9", body: "Notes" });
    const dialog = page.getByTestId("updater-dialog");
    await expect(dialog).toBeVisible();

    await fireTauriEvent(page, "update-download-started", { source: "startup", version: "9.9.9" });
    await expect(page.getByTestId("updater-progress-wrap")).toBeVisible();

    await fireTauriEvent(page, "update-download-progress", { source: "startup", version: "9.9.9", percent: 42 });
    await expect(page.getByTestId("updater-progress-text")).toContainText("42%");

    await fireTauriEvent(page, "update-downloaded", { source: "startup", version: "9.9.9" });
    await expect(page.getByTestId("updater-restart")).toBeVisible();
    await expect(page.getByTestId("update-ready-toast")).toHaveCount(0);

    await page.evaluate(() => {
      (window as any).__tauriInvokeCalls = [];
    });

    await page.getByTestId("updater-restart").click();
    await expect(dialog).toBeHidden();

    await page.waitForFunction(() => {
      const calls = (window as any).__tauriInvokeCalls;
      if (!Array.isArray(calls)) return false;
      const cmds = calls.map((call) => call?.cmd);
      return cmds.includes("install_downloaded_update") && (cmds.includes("restart_app") || cmds.includes("quit_app"));
    });
  });
});
