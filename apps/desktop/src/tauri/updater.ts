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
