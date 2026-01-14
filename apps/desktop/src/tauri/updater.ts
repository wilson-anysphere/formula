import { showToast } from "../extensions/ui.js";
import { t } from "../i18n/index.js";

import { getTauriInvokeOrNull } from "./api";

export type UpdateCheckSource = "manual";

function getTauriGlobalOrNull(): any | null {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (globalThis as any).__TAURI__ ?? null;
  } catch {
    // Some hardened host environments (or tests) may define `__TAURI__` with a throwing getter.
    // Treat that as "unavailable" so best-effort callsites can fall back cleanly.
    return null;
  }
}

function safeGetProp(obj: any, prop: string): any | undefined {
  if (!obj) return undefined;
  try {
    return obj[prop];
  } catch {
    return undefined;
  }
}

/**
 * Thin wrapper around Tauri's updater plugin so the UI can trigger installation/restart
 * without importing `@tauri-apps/*` directly.
 */
export async function installUpdateAndRestart(): Promise<void> {
  // Prefer our backend command which installs the already-downloaded update (from the startup
  // background download). This avoids waiting for a second download when the user approves a restart.
  const invoke = getTauriInvokeOrNull();
  if (invoke) {
    try {
      await invoke("install_downloaded_update");
      return;
    } catch (err) {
      // Fall back to the plugin API below (e.g. if the command isn't available yet or the
      // invoke is blocked by capabilities).
      console.warn("[formula][updater] Failed to invoke install_downloaded_update; falling back to updater plugin:", err);
    }
  }

  // We intentionally avoid importing `@tauri-apps/plugin-updater` because the desktop
  // frontend leans on global `__TAURI__` bindings (see `src/main.ts`).
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const tauri = getTauriGlobalOrNull();
  // Tauri's global API shape can vary across versions/builds.
  const plugin = safeGetProp(tauri, "plugin");
  const plugins = safeGetProp(tauri, "plugins");
  const updater = safeGetProp(tauri, "updater") ?? safeGetProp(plugin, "updater") ?? safeGetProp(plugins, "updater") ?? null;

  if (!updater) {
    throw new Error(t("updater.unavailable"));
  }

  // Tauri v2 updater plugin API shape (tauri-plugin-updater 2.x):
  // - `updater.check()` -> `Update | null`
  // - `Update.download()` -> downloads update package
  // - `Update.install()` -> installs downloaded package
  //
  // We intentionally avoid `Update.downloadAndInstall()` so we don't need to grant the
  // extra `updater:allow-download-and-install` capability permission.
  const check =
    (safeGetProp(updater, "check") ??
      safeGetProp(updater, "checkUpdate") ??
      safeGetProp(updater, "checkForUpdate") ??
      null) as unknown;
  if (typeof check !== "function") {
    throw new Error(t("updater.unavailable"));
  }

  const result = await (check as any).call(updater);
  if (!result) {
    throw new Error(t("updater.updateNoLongerAvailable"));
  }

  // Some APIs return `{ available: boolean, ... }`, others return an update object directly.
  const update =
    typeof result === "object" && result && "available" in result ? ((result as any).available ? result : null) : result;
  if (!update) {
    throw new Error(t("updater.updateNoLongerAvailable"));
  }

  const download = safeGetProp(update, "download") ?? null;
  const install = (safeGetProp(update, "install") ?? safeGetProp(update, "installUpdate") ?? null) as unknown;
  if (typeof download !== "function" || typeof install !== "function") {
    throw new Error(t("updater.unavailable"));
  }

  await (download as any).call(update);
  await (install as any).call(update);
}

/**
 * Trigger an in-app update check (desktop/Tauri only).
 *
 * The backend emits updater events that the UI listens to for toasts/prompts; this function
 * intentionally avoids showing additional UI (except when running outside the desktop app).
 */
export async function checkForUpdatesFromCommandPalette(source: UpdateCheckSource = "manual"): Promise<void> {
  const invoke = getTauriInvokeOrNull();
  if (!invoke) {
    try {
      showToast(t("updater.desktopOnly"));
    } catch (err) {
      // Avoid crashing lightweight embedders/tests that don't render a toast root.
      console.warn(t("updater.desktopOnly"), err);
    }
    return;
  }

  await invoke("check_for_updates", { source });
}
