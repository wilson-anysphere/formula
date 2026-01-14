import { getTauriInvokeOrNull } from "./api";

export type NotifyPayload = {
  title: string;
  body?: string;
};

type TauriNotify = (payload: { title: string; body?: string }) => Promise<void> | void;

function getTauriDirectNotify(): TauriNotify | null {
  const tauri = (globalThis as any).__TAURI__;
  const direct =
    (tauri?.notification?.notify as TauriNotify | undefined) ??
    (tauri?.notification?.sendNotification as TauriNotify | undefined) ??
    (tauri?.plugin?.notification?.notify as TauriNotify | undefined) ??
    (tauri?.plugin?.notification?.sendNotification as TauriNotify | undefined) ??
    (tauri?.plugins?.notification?.notify as TauriNotify | undefined) ??
    (tauri?.plugins?.notification?.sendNotification as TauriNotify | undefined) ??
    null;
  return typeof direct === "function" ? direct : null;
}

async function notifyWeb(payload: { title: string; body?: string }): Promise<void> {
  if (typeof Notification === "undefined") return;
  if (Notification.permission !== "granted") return;
  const options = payload.body ? { body: payload.body } : undefined;
  // `new Notification(...)` is synchronous but we keep the async shape for callers.
  // eslint-disable-next-line no-new
  new Notification(payload.title, options);
}

export async function notify(payload: NotifyPayload): Promise<void> {
  if (!payload || typeof payload.title !== "string") return;
  const title = payload.title.trim();
  if (!title) return;
  const body = typeof payload.body === "string" ? payload.body : undefined;
  const tauri = (globalThis as any).__TAURI__ ?? null;

  const directNotify = getTauriDirectNotify();
  if (directNotify) {
    try {
      await directNotify({ title, body });
      return;
    } catch {
      // Fall back to the invoke-based command below (useful if the direct API is
      // present but blocked by permissions or differs across Tauri versions).
    }
  }

  const invoke = getTauriInvokeOrNull();
  if (invoke) {
    try {
      await invoke("show_system_notification", { title, body: body ?? null });
      return;
    } catch {
      // Fall back to the web API below.
    }
  }

  // Only fall back to the Web Notification API in *web builds* (no Tauri runtime). In the desktop
  // shell we intentionally do **not** fall back to the web API because that would allow remote or
  // navigated-to content to create system notifications outside the hardened Rust command path.
  if (!tauri) {
    await notifyWeb({ title, body });
  }
}
