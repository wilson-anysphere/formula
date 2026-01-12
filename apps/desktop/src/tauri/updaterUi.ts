import { showToast } from "../extensions/ui.js";

import { requestAppRestart } from "./appQuit";
import { installUpdateAndRestart } from "./updater";

type UpdaterEventName = "update-check-started" | "update-not-available" | "update-check-error" | "update-available";

type UpdaterEventPayload = {
  source?: string;
  version?: string;
  body?: string | null;
  message?: string;
};

type TauriListen = (event: string, handler: (event: any) => void) => Promise<() => void>;

function getTauriListen(): TauriListen | null {
  const listen = (globalThis as any).__TAURI__?.event?.listen as TauriListen | undefined;
  if (typeof listen !== "function") return null;
  return listen;
}

function getTauriWindowHandle(): any | null {
  const winApi = (globalThis as any).__TAURI__?.window;
  if (!winApi) return null;

  // Mirrors the flexible handle lookup used in `main.ts`. We intentionally avoid
  // a hard dependency on `@tauri-apps/api`.
  const handle =
    (typeof winApi.getCurrentWebviewWindow === "function" ? winApi.getCurrentWebviewWindow() : null) ??
    (typeof winApi.getCurrentWindow === "function" ? winApi.getCurrentWindow() : null) ??
    (typeof winApi.getCurrent === "function" ? winApi.getCurrent() : null) ??
    winApi.appWindow ??
    null;

  return handle ?? null;
}

async function showMainWindowBestEffort(): Promise<void> {
  const win = getTauriWindowHandle();
  if (!win) return;

  try {
    if (typeof win.show === "function") {
      await win.show();
    }
  } catch {
    // Best-effort.
  }

  try {
    if (typeof win.setFocus === "function") {
      await win.setFocus();
    }
  } catch {
    // Best-effort.
  }
}

export async function handleUpdaterEvent(name: UpdaterEventName, payload: UpdaterEventPayload): Promise<void> {
  const source = payload?.source;

  // Tray-triggered manual checks can happen while the app is hidden to tray. Ensure the
  // window is visible before rendering any toast/dialog feedback.
  if (source === "manual") {
    await showMainWindowBestEffort();
  }

  if (source !== "manual") {
    // Preserve startup/automatic behavior: only show user-facing UI for explicit manual checks.
    return;
  }

  switch (name) {
    case "update-check-started": {
      showToast("Checking for updates...", "info");
      break;
    }
    case "update-not-available": {
      showToast("You're up to date.", "info");
      break;
    }
    case "update-check-error": {
      const message = typeof payload?.message === "string" && payload.message.trim() !== "" ? payload.message : "Unknown error";
      showToast(`Update check failed: ${message}`, "error");
      break;
    }
    case "update-available": {
      const version = typeof payload?.version === "string" && payload.version.trim() !== "" ? payload.version : "unknown";
      showToast(`Update available: ${version}`, "info");
      break;
    }
  }
}

export function installUpdaterUi(listenArg?: TauriListen): void {
  const listen = listenArg ?? getTauriListen();
  if (!listen) return;

  const events: UpdaterEventName[] = [
    "update-check-started",
    "update-not-available",
    "update-check-error",
    "update-available",
  ];

  for (const eventName of events) {
    void listen(eventName, (event) => {
      const payload = (event as any)?.payload as UpdaterEventPayload;
      void handleUpdaterEvent(eventName, payload);
    });
  }
}

/**
 * Called by the updater UI when the user confirms "Restart now".
 *
 * This routes through the normal quit flow (Workbook_BeforeClose macros, backend-sync drain,
 * and the unsaved-changes confirm prompt) before triggering the updater install step.
 */
export async function restartToInstallUpdate(): Promise<boolean> {
  return await requestAppRestart({
    beforeQuit: installUpdateAndRestart,
    beforeQuitErrorToast: "Failed to restart to install the update.",
  });
}

