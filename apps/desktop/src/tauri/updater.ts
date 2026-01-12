import { showToast } from "../extensions/ui.js";

type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;

function getTauriInvoke(): TauriInvoke | null {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  return typeof invoke === "function" ? invoke : null;
}

export type UpdateCheckSource = "manual";

/**
 * Thin wrapper around Tauri's updater plugin so the UI can trigger installation/restart
 * without importing `@tauri-apps/*` directly.
 */
export async function installUpdateAndRestart(): Promise<void> {
  // We intentionally avoid importing `@tauri-apps/plugin-updater` because the desktop
  // frontend leans on global `__TAURI__` bindings (see `src/main.ts`).
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const tauri = (globalThis as any).__TAURI__;
  const updater = tauri?.updater ?? tauri?.plugin?.updater;

  // Support a few likely API shapes (Tauri v2 plugin and potential wrappers).
  const install =
    updater?.install ??
    updater?.installUpdate ??
    updater?.downloadAndInstall ??
    updater?.downloadAndInstallUpdate ??
    null;

  if (typeof install !== "function") {
    throw new Error("Updater install API not available");
  }

  await install();
}

/**
 * Trigger an in-app update check (desktop/Tauri only).
 *
 * The backend emits updater events that the UI listens to for toasts/prompts; this function
 * intentionally avoids showing additional UI (except when running outside the desktop app).
 */
export async function checkForUpdatesFromCommandPalette(source: UpdateCheckSource = "manual"): Promise<void> {
  const invoke = getTauriInvoke();
  if (!invoke) {
    try {
      showToast("Update checks are only available in the desktop app.");
    } catch (err) {
      // Avoid crashing lightweight embedders/tests that don't render a toast root.
      console.warn("Update checks are only available in the desktop app.", err);
    }
    return;
  }

  await invoke("check_for_updates", { source });
}

