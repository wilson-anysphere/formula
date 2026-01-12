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

  if (!updater) {
    throw new Error("Updater API not available");
  }

  // Support a few likely API shapes:
  // - direct install/downloadAndInstall functions
  // - "check() -> update object -> downloadAndInstall()" style
  const directInstall =
    updater.install ??
    updater.installUpdate ??
    updater.downloadAndInstall ??
    updater.downloadAndInstallUpdate ??
    updater.downloadAndInstallUpdate ??
    null;

  if (typeof directInstall === "function") {
    await directInstall.call(updater);
    return;
  }

  const check = updater.check ?? updater.checkUpdate ?? updater.checkForUpdate ?? null;
  if (typeof check === "function") {
    const result = await check.call(updater);
    if (!result) {
      throw new Error("No update available");
    }

    // Some APIs return `{ available: boolean, ... }`, others return an update object directly.
    const update =
      typeof result === "object" && result && "available" in result
        ? (result as any).available
          ? result
          : null
        : result;
    if (!update) {
      throw new Error("No update available");
    }

    const updateInstall =
      (update as any).downloadAndInstall ??
      (update as any).downloadAndInstallUpdate ??
      (update as any).install ??
      (update as any).installUpdate ??
      null;
    if (typeof updateInstall === "function") {
      await updateInstall.call(update);
      return;
    }

    const download = (update as any).download ?? null;
    const install = (update as any).install ?? (update as any).installUpdate ?? null;
    if (typeof download === "function" && typeof install === "function") {
      await download.call(update);
      await install.call(update);
      return;
    }
  }

  throw new Error("Updater install API not available");
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

