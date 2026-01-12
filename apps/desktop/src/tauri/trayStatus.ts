export type TrayStatus = "idle" | "syncing" | "error";

type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;

function getTauriInvokeOrNull(): TauriInvoke | null {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  return typeof invoke === "function" ? invoke : null;
}

/**
 * Best-effort tray status update.
 *
 * - No-ops when running outside Tauri (web build, tests without a mocked `__TAURI__`).
 * - Swallows invoke errors so older backends don't break the UI.
 */
export async function setTrayStatus(status: TrayStatus): Promise<void> {
  const invoke = getTauriInvokeOrNull();
  if (!invoke) return;
  try {
    await invoke("set_tray_status", { status });
  } catch {
    // Graceful degradation: ignore missing command / invoke failures.
  }
}

