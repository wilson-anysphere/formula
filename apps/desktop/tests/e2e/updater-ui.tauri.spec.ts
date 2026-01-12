import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function flushMicrotasks(page: Page, times = 6): Promise<void> {
  await page.evaluate(async (n) => {
    for (let idx = 0; idx < n; idx += 1) {
      await new Promise<void>((resolve) => queueMicrotask(resolve));
    }
  }, times);
}

function lastToast(page: Page) {
  return page.locator('#toast-root [data-testid="toast"]').last();
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
      const invokes: Array<{ cmd: string; args: any }> = [];
      const callOrder: Array<{ kind: "listen" | "listen-registered" | "emit"; name: string; seq: number }> = [];
      let seq = 0;

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriEmittedEvents = emitted;
      (window as any).__tauriInvokes = invokes;
      (window as any).__tauriCallOrder = callOrder;
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
            callOrder.push({ kind: "listen", name, seq: ++seq });
            // Simulate async backend confirmation for handler registration so we can
            // assert `updater-ui-ready` is emitted only after all listeners are installed.
            await Promise.resolve();
            if (!Array.isArray(listeners[name])) {
              listeners[name] = [];
            }
            listeners[name].push(handler);
            callOrder.push({ kind: "listen-registered", name, seq: ++seq });
            return () => {
              const arr = listeners[name];
              if (!Array.isArray(arr)) return;
              const idx = arr.indexOf(handler);
              if (idx >= 0) arr.splice(idx, 1);
            };
          },
          emit: async (event: string, payload?: any) => {
            callOrder.push({ kind: "emit", name: event, seq: ++seq });
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
    const readyOrdering = await page.evaluate(() => {
      const calls = (window as any).__tauriCallOrder as Array<{ kind: string; name: string; seq: number }> | undefined;
      if (!Array.isArray(calls)) return null;
      const readySeq = calls.find((c) => c.kind === "emit" && c.name === "updater-ui-ready")?.seq ?? null;
      const required = [
        "update-check-already-running",
        "update-check-started",
        "update-not-available",
        "update-check-error",
        "update-available",
      ];
      const registered = Object.fromEntries(
        required.map((event) => [
          event,
          calls.find((c) => c.kind === "listen-registered" && c.name === event)?.seq ?? null,
        ]),
      );
      return { readySeq, registered };
    });
    expect(readyOrdering).not.toBeNull();
    expect(readyOrdering!.readySeq).not.toBeNull();
    for (const seqValue of Object.values(readyOrdering!.registered as Record<string, number | null>)) {
      expect(seqValue).not.toBeNull();
      expect(seqValue!).toBeLessThan(readyOrdering!.readySeq!);
    }

    await waitForTauriListeners(page, "update-check-started");
    await dispatchTauriEvent(page, "update-check-started", { source: "manual" });

    await expect(lastToast(page)).toHaveText("Checking for updates…");
    await page.waitForFunction(() => (window as any).__tauriShowCalls >= 1);
    await page.waitForFunction(() => (window as any).__tauriFocusCalls >= 1);

    await waitForTauriListeners(page, "update-check-error");
    await dispatchTauriEvent(page, "update-check-error", { source: "manual", error: "network down" });
    await expect(lastToast(page)).toHaveText("Error: network down");

    await waitForTauriListeners(page, "update-check-already-running");
    await dispatchTauriEvent(page, "update-check-already-running", { source: "manual" });
    await expect(lastToast(page)).toHaveText("Already checking for updates…");

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

    // Verify the dialog action wiring uses the Tauri shell open command.
    const openBefore = await page.evaluate(() => {
      const invokes = (window as any).__tauriInvokes ?? [];
      if (!Array.isArray(invokes)) return 0;
      return invokes.filter((entry: any) => entry?.cmd === "open_external_url").length;
    });
    await page.getByTestId("updater-view-versions").click();
    await page.waitForFunction(
      (expectedCount) => {
        const invokes = (window as any).__tauriInvokes ?? [];
        if (!Array.isArray(invokes)) return false;
        const count = invokes.filter((entry: any) => entry?.cmd === "open_external_url").length;
        return count >= expectedCount;
      },
      openBefore + 1,
    );
    const openExternal = await page.evaluate(
      () => {
        const invokes = (window as any).__tauriInvokes ?? [];
        if (!Array.isArray(invokes)) return null;
        const matches = invokes.filter((entry: any) => entry?.cmd === "open_external_url");
        return matches.length > 0 ? matches[matches.length - 1] : null;
      },
    );
    expect(openExternal).not.toBeNull();
    expect(String(openExternal.args?.url ?? "")).toContain("github.com/wilson-anysphere/formula/releases");

    await waitForTauriListeners(page, "update-not-available");
    await dispatchTauriEvent(page, "update-not-available", { source: "manual" });
    await expect(lastToast(page)).toHaveText("You're up to date.");
  });

  test("can download an update and trigger restart from the updater dialog", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, Array<(event: any) => void>> = {};
      const emitted: Array<{ event: string; payload: any }> = [];
      const invokes: Array<{ cmd: string; args: any }> = [];
      const actionOrder: Array<{ kind: string; name: string }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriEmittedEvents = emitted;
      (window as any).__tauriInvokes = invokes;
      (window as any).__tauriActionOrder = actionOrder;
      (window as any).__tauriUpdateDownloadCalls = 0;
      (window as any).__tauriUpdateInstallCalls = 0;

      // Avoid any prompts blocking the restart flow.
      window.confirm = () => true;

      const windowHandle = { show: async () => {}, setFocus: async () => {} };

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });
            actionOrder.push({ kind: "invoke", name: cmd });
            return null;
          },
        },
        event: {
          listen: async (name: string, handler: any) => {
            if (!Array.isArray(listeners[name])) listeners[name] = [];
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
        updater: {
          check: async () => ({
            version: "9.9.9",
            body: "Notes",
            download: async (onProgress?: any) => {
              (window as any).__tauriUpdateDownloadCalls += 1;
              actionOrder.push({ kind: "download", name: "update.download" });
              onProgress?.({ downloaded: 100, total: 100 });
            },
            install: async () => {
              (window as any).__tauriUpdateInstallCalls += 1;
              actionOrder.push({ kind: "install", name: "update.install" });
            },
          }),
        },
      };
    });

    await gotoDesktop(page);

    await page.waitForFunction(() =>
      Boolean((window as any).__tauriEmittedEvents?.some((entry: any) => entry?.event === "updater-ui-ready")),
    );

    await waitForTauriListeners(page, "update-available");
    await dispatchTauriEvent(page, "update-available", { source: "manual", version: "9.9.9", body: "Notes" });

    const dialog = page.getByTestId("updater-dialog");
    await expect(dialog).toBeVisible();

    await page.getByTestId("updater-download").click();
    await expect(page.getByTestId("updater-progress-wrap")).toBeVisible();

    await page.waitForFunction(() => (window as any).__tauriUpdateDownloadCalls >= 1);

    await expect(page.getByTestId("updater-restart")).toBeVisible();
    await expect(page.getByTestId("updater-status")).toContainText("Download complete.");
    await expect(page.getByTestId("updater-progress-wrap")).toBeHidden();

    const restartCountBefore = await page.evaluate(() => {
      const invokes = (window as any).__tauriInvokes ?? [];
      if (!Array.isArray(invokes)) return 0;
      return invokes.filter((entry: any) => entry?.cmd === "restart_app").length;
    });

    await page.getByTestId("updater-restart").click();

    await page.waitForFunction(
      (before) => {
        const invokes = (window as any).__tauriInvokes ?? [];
        if (!Array.isArray(invokes)) return false;
        const count = invokes.filter((entry: any) => entry?.cmd === "restart_app").length;
        return count >= before + 1;
      },
      restartCountBefore,
    );

    await page.waitForFunction(() => (window as any).__tauriUpdateInstallCalls >= 1);
    await expect(dialog).toBeHidden();

    const order = await page.evaluate(() => (window as any).__tauriActionOrder ?? []);
    expect(Array.isArray(order)).toBe(true);
    const installIdx = (order as any[]).findIndex((entry: any) => entry?.kind === "install" && entry?.name === "update.install");
    const restartIdx = (order as any[]).findIndex((entry: any) => entry?.kind === "invoke" && entry?.name === "restart_app");
    expect(installIdx).toBeGreaterThanOrEqual(0);
    expect(restartIdx).toBeGreaterThanOrEqual(0);
    expect(installIdx).toBeLessThan(restartIdx);
  });

  test("updater dialog cannot be dismissed while a download is in flight", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, Array<(event: any) => void>> = {};
      const emitted: Array<{ event: string; payload: any }> = [];
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriEmittedEvents = emitted;
      (window as any).__tauriInvokes = invokes;
      (window as any).__tauriUpdateDownloadCalls = 0;

      // Avoid any prompts blocking the flow.
      window.confirm = () => true;

      const windowHandle = { show: async () => {}, setFocus: async () => {} };

      let resolveDownload: (() => void) | null = null;
      const downloadGate = new Promise<void>((resolve) => {
        resolveDownload = resolve;
      });
      (window as any).__tauriResolveDownload = () => resolveDownload?.();

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });
            return null;
          },
        },
        event: {
          listen: async (name: string, handler: any) => {
            if (!Array.isArray(listeners[name])) listeners[name] = [];
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
        updater: {
          check: async () => ({
            version: "9.9.9",
            body: "Notes",
            download: async (onProgress?: any) => {
              (window as any).__tauriUpdateDownloadCalls += 1;
              onProgress?.({ downloaded: 10, total: 100 });
              await downloadGate;
              onProgress?.({ downloaded: 100, total: 100 });
            },
            install: async () => {},
          }),
        },
      };
    });

    await gotoDesktop(page);
    await page.waitForFunction(() =>
      Boolean((window as any).__tauriEmittedEvents?.some((entry: any) => entry?.event === "updater-ui-ready")),
    );

    await waitForTauriListeners(page, "update-available");
    await dispatchTauriEvent(page, "update-available", { source: "manual", version: "9.9.9", body: "Notes" });

    const dialog = page.getByTestId("updater-dialog");
    await expect(dialog).toBeVisible();

    await page.getByTestId("updater-download").click();
    await expect(page.getByTestId("updater-progress-wrap")).toBeVisible();
    await page.waitForFunction(() => (window as any).__tauriUpdateDownloadCalls >= 1);

    // Cancel (Escape) should be ignored while download is in progress.
    await page.evaluate(() => {
      const dialogEl = document.querySelector("dialog[data-testid=\"updater-dialog\"]") as HTMLDialogElement | null;
      dialogEl?.dispatchEvent(new Event("cancel", { cancelable: true }));
    });
    await expect(dialog).toBeVisible();
    await expect(dialog).toHaveJSProperty("open", true);

    // View versions should still open an external URL but should not close the dialog during download.
    const openExternalBefore = await page.evaluate(() => {
      const invokes = (window as any).__tauriInvokes ?? [];
      if (!Array.isArray(invokes)) return 0;
      return invokes.filter((entry: any) => entry?.cmd === "open_external_url").length;
    });
    await page.getByTestId("updater-view-versions").click();
    await page.waitForFunction(
      (expectedCount) => {
        const invokes = (window as any).__tauriInvokes ?? [];
        if (!Array.isArray(invokes)) return false;
        const count = invokes.filter((entry: any) => entry?.cmd === "open_external_url").length;
        return count >= expectedCount;
      },
      openExternalBefore + 1,
    );
    await expect(dialog).toBeVisible();
    await expect(dialog).toHaveJSProperty("open", true);

    // Finish the download.
    await page.evaluate(() => {
      (window as any).__tauriResolveDownload?.();
    });
    await expect(page.getByTestId("updater-progress-wrap")).toBeHidden();
    await expect(page.getByTestId("updater-restart")).toBeVisible();
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
    await expect(lastToast(page)).toHaveText("Already checking for updates…");
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

  test("startup completion toasts are surfaced after the user queues a manual check behind an in-flight startup check", async ({
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
            if (!Array.isArray(listeners[name])) listeners[name] = [];
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

    await waitForTauriListeners(page, "update-check-already-running");
    await dispatchTauriEvent(page, "update-check-already-running", { source: "manual" });
    await expect(lastToast(page)).toHaveText("Already checking for updates…");
    await page.waitForFunction(() => (window as any).__tauriShowCalls >= 1);
    await page.waitForFunction(() => (window as any).__tauriFocusCalls >= 1);

    // The backend reports completion as "startup", but because the user requested a manual check
    // while the startup check was in flight, we should surface the result toast (without re-focusing).
    await waitForTauriListeners(page, "update-not-available");
    await dispatchTauriEvent(page, "update-not-available", { source: "startup" });
    await expect(lastToast(page)).toHaveText("You're up to date.");

    const afterUpToDate = await page.evaluate(() => ({
      show: (window as any).__tauriShowCalls,
      focus: (window as any).__tauriFocusCalls,
      notifications: (window as any).__tauriNotifications?.length ?? 0,
      notifiedViaInvoke: ((window as any).__tauriInvokes ?? []).some((entry: any) => entry?.cmd === "show_system_notification"),
    }));
    expect(afterUpToDate.show).toBe(1);
    expect(afterUpToDate.focus).toBe(1);
    expect(afterUpToDate.notifications).toBe(0);
    expect(afterUpToDate.notifiedViaInvoke).toBe(false);

    // Repeat for the error completion case.
    await dispatchTauriEvent(page, "update-check-already-running", { source: "manual" });
    await expect(lastToast(page)).toHaveText("Already checking for updates…");

    await waitForTauriListeners(page, "update-check-error");
    await dispatchTauriEvent(page, "update-check-error", { source: "startup", error: "network down" });
    await expect(lastToast(page)).toHaveText("Error: network down");

    const afterError = await page.evaluate(() => ({
      show: (window as any).__tauriShowCalls,
      focus: (window as any).__tauriFocusCalls,
    }));
    expect(afterError.show).toBe(2);
    expect(afterError.focus).toBe(2);
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
    // Exercise the ESC/cancel path: it should behave like clicking "Later" and persist suppression.
    await page.evaluate(() => {
      const dialogEl = document.querySelector("dialog[data-testid=\"updater-dialog\"]");
      dialogEl?.dispatchEvent(new Event("cancel", { cancelable: true }));
    });
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

    // TTL expiry should re-enable startup notifications.
    await page.evaluate(() => {
      const eightDaysAgoMs = Date.now() - 8 * 24 * 60 * 60 * 1000;
      localStorage.setItem("formula.updater.dismissedVersion", "9.9.9");
      localStorage.setItem("formula.updater.dismissedAt", String(eightDaysAgoMs));

      const notifications = (window as any).__tauriNotifications;
      if (Array.isArray(notifications)) notifications.length = 0;
      const invokes = (window as any).__tauriInvokes;
      if (Array.isArray(invokes)) invokes.length = 0;
    });

    await dispatchTauriEvent(page, "update-available", { source: "startup", version: "9.9.9", body: "Notes" });
    await flushMicrotasks(page);
    await expect(dialog).toBeHidden();

    const ttlResult = await page.evaluate(() => ({
      version: localStorage.getItem("formula.updater.dismissedVersion"),
      dismissedAt: localStorage.getItem("formula.updater.dismissedAt"),
      notificationsCount: Array.isArray((window as any).__tauriNotifications) ? (window as any).__tauriNotifications.length : 0,
      invoked: Array.isArray((window as any).__tauriInvokes) &&
        (window as any).__tauriInvokes.some((entry: any) => entry?.cmd === "show_system_notification"),
    }));
    expect(ttlResult.version).toBeNull();
    expect(ttlResult.dismissedAt).toBeNull();
    expect(ttlResult.notificationsCount > 0 || ttlResult.invoked).toBe(true);

    // A different startup version should also clear any stored dismissal and re-notify.
    await page.evaluate(() => {
      localStorage.setItem("formula.updater.dismissedVersion", "9.9.9");
      localStorage.setItem("formula.updater.dismissedAt", String(Date.now()));

      const notifications = (window as any).__tauriNotifications;
      if (Array.isArray(notifications)) notifications.length = 0;
      const invokes = (window as any).__tauriInvokes;
      if (Array.isArray(invokes)) invokes.length = 0;
    });

    await dispatchTauriEvent(page, "update-available", { source: "startup", version: "9.9.10", body: "Notes" });
    await flushMicrotasks(page);

    const versionResult = await page.evaluate(() => ({
      version: localStorage.getItem("formula.updater.dismissedVersion"),
      dismissedAt: localStorage.getItem("formula.updater.dismissedAt"),
      notificationsCount: Array.isArray((window as any).__tauriNotifications) ? (window as any).__tauriNotifications.length : 0,
      invoked: Array.isArray((window as any).__tauriInvokes) &&
        (window as any).__tauriInvokes.some((entry: any) => entry?.cmd === "show_system_notification"),
    }));
    expect(versionResult.version).toBeNull();
    expect(versionResult.dismissedAt).toBeNull();
    expect(versionResult.notificationsCount > 0 || versionResult.invoked).toBe(true);
  });

  test("clicking Download clears any stored dismissal (even when the updater API is unavailable)", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, Array<(event: any) => void>> = {};
      const emitted: Array<{ event: string; payload: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriEmittedEvents = emitted;

      const windowHandle = { show: async () => {}, setFocus: async () => {} };

      // Intentionally omit `__TAURI__.updater` so the UI takes the "unavailable" path.
      (window as any).__TAURI__ = {
        core: {
          invoke: async (_cmd: string, _args: any) => null,
        },
        event: {
          listen: async (name: string, handler: any) => {
            if (!Array.isArray(listeners[name])) listeners[name] = [];
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
      };
    });

    await gotoDesktop(page);

    await page.waitForFunction(() =>
      Boolean((window as any).__tauriEmittedEvents?.some((entry: any) => entry?.event === "updater-ui-ready")),
    );

    await waitForTauriListeners(page, "update-available");

    // Seed a stored dismissal for this version.
    await page.evaluate(() => {
      localStorage.setItem("formula.updater.dismissedVersion", "9.9.9");
      localStorage.setItem("formula.updater.dismissedAt", String(Date.now()));
    });

    await dispatchTauriEvent(page, "update-available", { source: "manual", version: "9.9.9", body: "Notes" });
    const dialog = page.getByTestId("updater-dialog");
    await expect(dialog).toBeVisible();

    const before = await page.evaluate(() => ({
      version: localStorage.getItem("formula.updater.dismissedVersion"),
      dismissedAt: localStorage.getItem("formula.updater.dismissedAt"),
    }));
    expect(before.version).toBe("9.9.9");
    expect(Number(before.dismissedAt)).toBeGreaterThan(0);

    // User initiates an update download; this should clear the persisted suppression even if
    // the download cannot start in this environment.
    await page.getByTestId("updater-download").click();

    await page.waitForFunction(() => localStorage.getItem("formula.updater.dismissedVersion") === null);
    await page.waitForFunction(() => localStorage.getItem("formula.updater.dismissedAt") === null);

    await expect(dialog).toBeVisible();
    await expect(lastToast(page)).toHaveText("Auto-updater is unavailable in this build.");
    await expect(page.getByTestId("updater-view-versions")).toHaveText("Download manually");
  });

  test("menu-check-updates triggers a fallback check_for_updates invoke (and suppresses duplicates)", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, Array<(event: any) => void>> = {};
      const emitted: Array<{ event: string; payload: any }> = [];
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriEmittedEvents = emitted;
      (window as any).__tauriInvokes = invokes;

      const windowHandle = { show: async () => {}, setFocus: async () => {} };

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });
            return null;
          },
        },
        event: {
          listen: async (name: string, handler: any) => {
            if (!Array.isArray(listeners[name])) listeners[name] = [];
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
      };
    });

    await gotoDesktop(page);

    // Ensure all desktop listeners installed before we emit.
    await page.waitForFunction(() =>
      Boolean((window as any).__tauriEmittedEvents?.some((entry: any) => entry?.event === "updater-ui-ready")),
    );

    await page.evaluate(() => {
      const invokes = (window as any).__tauriInvokes;
      if (Array.isArray(invokes)) invokes.length = 0;
    });

    await waitForTauriListeners(page, "menu-check-updates");
    await dispatchTauriEvent(page, "menu-check-updates", null);

    // The listener uses a small setTimeout before invoking `check_for_updates`.
    await page.waitForFunction(
      () => Array.isArray((window as any).__tauriInvokes) && (window as any).__tauriInvokes.some((e: any) => e?.cmd === "check_for_updates"),
      undefined,
      { timeout: 5_000 },
    );
    const firstInvoke = await page.evaluate(() =>
      (window as any).__tauriInvokes?.find((entry: any) => entry?.cmd === "check_for_updates") ?? null,
    );
    expect(firstInvoke).not.toBeNull();
    expect(firstInvoke.args?.source).toBe("manual");

    // If the backend has already kicked off a manual update check, menu-check-updates should not
    // trigger another frontend-initiated check.
    await page.evaluate(() => {
      const invokes = (window as any).__tauriInvokes;
      if (Array.isArray(invokes)) invokes.length = 0;
    });

    await waitForTauriListeners(page, "update-check-started");
    await dispatchTauriEvent(page, "update-check-started", { source: "manual" });
    await dispatchTauriEvent(page, "menu-check-updates", null);

    let invoked = false;
    try {
      await page.waitForFunction(
        () =>
          Array.isArray((window as any).__tauriInvokes) &&
          (window as any).__tauriInvokes.some((e: any) => e?.cmd === "check_for_updates"),
        undefined,
        { timeout: 400 },
      );
      invoked = true;
    } catch {
      invoked = false;
    }
    expect(invoked).toBe(false);
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

    const baseline = await page.evaluate(() => ({
      toastCount: document.querySelectorAll('#toast-root [data-testid="toast"]').length,
      showCalls: (window as any).__tauriShowCalls ?? 0,
      focusCalls: (window as any).__tauriFocusCalls ?? 0,
    }));

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

    // Startup update events should not force-show or focus the main window and should not create
    // in-app toasts.
    const afterUpdateAvailable = await page.evaluate(() => ({
      toastCount: document.querySelectorAll('#toast-root [data-testid="toast"]').length,
      showCalls: (window as any).__tauriShowCalls ?? 0,
      focusCalls: (window as any).__tauriFocusCalls ?? 0,
      notificationsCount: (window as any).__tauriNotifications?.length ?? 0,
      invokeNotificationCount: Array.isArray((window as any).__tauriInvokes)
        ? (window as any).__tauriInvokes.filter((entry: any) => entry?.cmd === "show_system_notification").length
        : 0,
    }));
    expect(afterUpdateAvailable.toastCount).toBe(baseline.toastCount);
    expect(afterUpdateAvailable.showCalls).toBe(baseline.showCalls);
    expect(afterUpdateAvailable.focusCalls).toBe(baseline.focusCalls);

    // Startup update-check lifecycle events should also be silent (no toast/focus).
    await waitForTauriListeners(page, "update-check-started");
    await dispatchTauriEvent(page, "update-check-started", { source: "startup" });
    await waitForTauriListeners(page, "update-check-already-running");
    await dispatchTauriEvent(page, "update-check-already-running", { source: "startup" });
    await flushMicrotasks(page);

    const afterStartupCheckEvents = await page.evaluate(() => ({
      toastCount: document.querySelectorAll('#toast-root [data-testid="toast"]').length,
      showCalls: (window as any).__tauriShowCalls ?? 0,
      focusCalls: (window as any).__tauriFocusCalls ?? 0,
      notificationsCount: (window as any).__tauriNotifications?.length ?? 0,
      invokeNotificationCount: Array.isArray((window as any).__tauriInvokes)
        ? (window as any).__tauriInvokes.filter((entry: any) => entry?.cmd === "show_system_notification").length
        : 0,
    }));
    expect(afterStartupCheckEvents.toastCount).toBe(baseline.toastCount);
    expect(afterStartupCheckEvents.showCalls).toBe(baseline.showCalls);
    expect(afterStartupCheckEvents.focusCalls).toBe(baseline.focusCalls);
    expect(afterStartupCheckEvents.notificationsCount).toBe(afterUpdateAvailable.notificationsCount);
    expect(afterStartupCheckEvents.invokeNotificationCount).toBe(afterUpdateAvailable.invokeNotificationCount);

    // Startup completion events should also remain silent unless a manual follow-up is pending.
    await waitForTauriListeners(page, "update-not-available");
    await dispatchTauriEvent(page, "update-not-available", { source: "startup" });
    await flushMicrotasks(page);

    const afterNotAvailable = await page.evaluate(() => ({
      toastCount: document.querySelectorAll('#toast-root [data-testid="toast"]').length,
      notificationsCount: (window as any).__tauriNotifications?.length ?? 0,
      invokeNotificationCount: Array.isArray((window as any).__tauriInvokes)
        ? (window as any).__tauriInvokes.filter((entry: any) => entry?.cmd === "show_system_notification").length
        : 0,
    }));
    expect(afterNotAvailable.toastCount).toBe(baseline.toastCount);
    // No additional system notifications for "up to date".
    expect(afterNotAvailable.notificationsCount).toBe(afterUpdateAvailable.notificationsCount);
    expect(afterNotAvailable.invokeNotificationCount).toBe(afterUpdateAvailable.invokeNotificationCount);

    await waitForTauriListeners(page, "update-check-error");
    await dispatchTauriEvent(page, "update-check-error", { source: "startup", error: "network down" });
    await flushMicrotasks(page);
    const afterStartupError = await page.evaluate(() => ({
      toastCount: document.querySelectorAll('#toast-root [data-testid="toast"]').length,
      notificationsCount: (window as any).__tauriNotifications?.length ?? 0,
      invokeNotificationCount: Array.isArray((window as any).__tauriInvokes)
        ? (window as any).__tauriInvokes.filter((entry: any) => entry?.cmd === "show_system_notification").length
        : 0,
    }));
    expect(afterStartupError.toastCount).toBe(baseline.toastCount);
    expect(afterStartupError.notificationsCount).toBe(afterUpdateAvailable.notificationsCount);
    expect(afterStartupError.invokeNotificationCount).toBe(afterUpdateAvailable.invokeNotificationCount);
  });
});
