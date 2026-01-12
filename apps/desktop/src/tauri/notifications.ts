export type NotifyPayload = {
  title: string;
  body?: string;
};

type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;
type TauriNotify = (payload: { title: string; body?: string }) => Promise<void> | void;

function getTauriInvoke(): TauriInvoke | null {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  return typeof invoke === "function" ? invoke : null;
}

function getTauriNotify(): TauriNotify | null {
  const tauri = (globalThis as any).__TAURI__;

  const direct =
    (tauri?.notification?.notify as TauriNotify | undefined) ??
    (tauri?.notification?.sendNotification as TauriNotify | undefined) ??
    (tauri?.plugin?.notification?.notify as TauriNotify | undefined) ??
    (tauri?.plugin?.notification?.sendNotification as TauriNotify | undefined) ??
    null;

  if (typeof direct === "function") return direct;

  const invoke = getTauriInvoke();
  if (!invoke) return null;

  return (payload) =>
    invoke("show_system_notification", {
      title: payload.title,
      body: payload.body ?? null,
    }).then(() => undefined);
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

  const native = getTauriNotify();
  if (native) {
    await native({ title, body });
    return;
  }

  await notifyWeb({ title, body });
}

