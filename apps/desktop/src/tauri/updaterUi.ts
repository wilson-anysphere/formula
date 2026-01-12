import { showToast } from "../extensions/ui.js";

import { requestAppRestart } from "./appQuit";
import { installUpdateAndRestart } from "./updater";

export const FORMULA_RELEASES_URL = "https://github.com/wilson-anysphere/formula/releases";

type UpdaterEventName =
  | "update-check-already-running"
  | "update-check-started"
  | "update-not-available"
  | "update-check-error"
  | "update-available";

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

let updateDialogShownForVersion: string | null = null;

async function openExternalUrl(url: string): Promise<void> {
  const tauri = (globalThis as any).__TAURI__;
  const tauriOpen = tauri?.shell?.open ?? tauri?.plugin?.shell?.open;
  if (typeof tauriOpen === "function") {
    await tauriOpen(url);
    return;
  }

  if (typeof window !== "undefined" && typeof window.open === "function") {
    window.open(url, "_blank", "noopener,noreferrer");
  }
}

function styleDialogButton(btn: HTMLButtonElement, variant: "primary" | "secondary" = "secondary") {
  btn.style.border = "1px solid var(--border)";
  btn.style.borderRadius = "10px";
  btn.style.padding = "8px 12px";
  btn.style.cursor = "pointer";
  btn.style.background = variant === "primary" ? "var(--accent)" : "var(--bg-primary)";
  btn.style.color = variant === "primary" ? "var(--text-on-accent)" : "var(--text-primary)";
}

function showUpdateAvailableDialog(update: { version: string; body: string | null }): void {
  // Fall back to a toast in environments that don't support `<dialog>`.
  if (typeof document === "undefined" || typeof document.createElement !== "function" || typeof (window as any).HTMLDialogElement === "undefined") {
    showToast(`Update available: ${update.version}`, "info");
    return;
  }

  const dialog = document.createElement("dialog");
  dialog.className = "dialog";
  dialog.dataset.testid = "update-dialog";
  dialog.style.maxWidth = "min(640px, calc(100vw - 32px))";
  dialog.style.width = "520px";
  dialog.style.boxShadow = "var(--dialog-shadow)";

  const title = document.createElement("div");
  title.className = "dialog__title";
  title.textContent = "Update available";

  const intro = document.createElement("div");
  intro.textContent = `Formula ${update.version} is available.`;
  intro.style.marginBottom = "10px";

  const notes = document.createElement("div");
  notes.style.whiteSpace = "pre-wrap";
  notes.style.fontSize = "13px";
  notes.style.lineHeight = "18px";
  notes.style.color = "var(--text-secondary)";
  notes.style.marginBottom = "12px";
  notes.textContent = update.body ?? "";

  const rollbackHint = document.createElement("div");
  rollbackHint.style.fontSize = "12px";
  rollbackHint.style.lineHeight = "16px";
  rollbackHint.style.color = "var(--text-secondary)";
  rollbackHint.textContent =
    "Need to downgrade? Use “View all versions” to download and install any previous release for your platform.";

  const controls = document.createElement("div");
  controls.style.display = "flex";
  controls.style.justifyContent = "flex-end";
  controls.style.gap = "8px";
  controls.style.marginTop = "14px";

  const laterBtn = document.createElement("button");
  laterBtn.type = "button";
  laterBtn.textContent = "Later";
  laterBtn.dataset.testid = "update-dialog-later";
  styleDialogButton(laterBtn, "secondary");

  const viewVersionsBtn = document.createElement("button");
  viewVersionsBtn.type = "button";
  viewVersionsBtn.textContent = "View all versions";
  viewVersionsBtn.dataset.testid = "update-dialog-view-versions";
  styleDialogButton(viewVersionsBtn, "secondary");

  const restartBtn = document.createElement("button");
  restartBtn.type = "button";
  restartBtn.textContent = "Restart to update";
  restartBtn.dataset.testid = "update-dialog-restart";
  styleDialogButton(restartBtn, "primary");

  controls.appendChild(laterBtn);
  controls.appendChild(viewVersionsBtn);
  controls.appendChild(restartBtn);

  dialog.appendChild(title);
  dialog.appendChild(intro);
  if (update.body) dialog.appendChild(notes);
  dialog.appendChild(rollbackHint);
  dialog.appendChild(controls);

  document.body.appendChild(dialog);

  const cleanup = () => {
    dialog.remove();
  };

  dialog.addEventListener(
    "close",
    () => {
      cleanup();
    },
    { once: true },
  );

  dialog.addEventListener("cancel", (e) => {
    e.preventDefault();
    dialog.close();
  });

  laterBtn.addEventListener("click", () => dialog.close());
  viewVersionsBtn.addEventListener("click", () => {
    void openExternalUrl(FORMULA_RELEASES_URL);
    dialog.close();
  });
  restartBtn.addEventListener("click", () => {
    void restartToInstallUpdate();
    dialog.close();
  });

  dialog.showModal();
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
    case "update-check-already-running": {
      showToast("Already checking for updates...", "info");
      break;
    }
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
      const version = typeof payload?.version === "string" && payload.version.trim() !== "" ? payload.version.trim() : "unknown";
      const body = typeof payload?.body === "string" && payload.body.trim() !== "" ? payload.body : null;
      if (updateDialogShownForVersion === version) break;
      updateDialogShownForVersion = version;
      showUpdateAvailableDialog({ version, body });
      break;
    }
  }
}

export function installUpdaterUi(listenArg?: TauriListen): void {
  const listen = listenArg ?? getTauriListen();
  if (!listen) return;

  const events: UpdaterEventName[] = [
    "update-check-already-running",
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
