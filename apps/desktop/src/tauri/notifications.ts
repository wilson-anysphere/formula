import { getTauriInvokeOrNull } from "./api";

export type NotifyPayload = {
  title: string;
  body?: string;
};

type TauriNotify = (payload: { title: string; body?: string }) => Promise<void> | void;

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

function getTauriDirectNotify(): TauriNotify | null {
  const tauri = getTauriGlobalOrNull();
  const plugin = safeGetProp(tauri, "plugin");
  const plugins = safeGetProp(tauri, "plugins");
  const notification = safeGetProp(tauri, "notification");
  const pluginNotification = safeGetProp(plugin, "notification");
  const pluginsNotification = safeGetProp(plugins, "notification");

  const candidates: Array<{ owner: any; fn: unknown }> = [
    { owner: notification, fn: safeGetProp(notification, "notify") },
    { owner: notification, fn: safeGetProp(notification, "sendNotification") },
    { owner: pluginNotification, fn: safeGetProp(pluginNotification, "notify") },
    { owner: pluginNotification, fn: safeGetProp(pluginNotification, "sendNotification") },
    { owner: pluginsNotification, fn: safeGetProp(pluginsNotification, "notify") },
    { owner: pluginsNotification, fn: safeGetProp(pluginsNotification, "sendNotification") },
  ];

  for (const { owner, fn } of candidates) {
    if (typeof fn !== "function") continue;
    const notify = fn as TauriNotify;
    return async (payload) => {
      await notify.call(owner, payload);
    };
  }

  return null;
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
  const tauri = getTauriGlobalOrNull();

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
