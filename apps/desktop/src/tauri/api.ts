import {
  getTauriInvokeOrNull as getTauriInvokeOrNullRuntime,
  getTauriInvokeOrThrow as getTauriInvokeOrThrowRuntime,
  hasTauriInvoke as hasTauriInvokeRuntime,
} from "./invoke.js";

export type TauriDialogOpen = (options?: Record<string, unknown>) => Promise<string | string[] | null>;
export type TauriDialogSave = (options?: Record<string, unknown>) => Promise<string | null>;
export type TauriDialogConfirm = (message: string, options?: Record<string, unknown>) => Promise<boolean>;
export type TauriDialogMessage = (message: string, options?: Record<string, unknown>) => Promise<void>;

export type TauriInvoke = (cmd: string, args?: any) => Promise<any>;

export type TauriAppGetName = () => Promise<string>;
export type TauriAppGetVersion = () => Promise<string>;

export type TauriDialogApi = {
  open: TauriDialogOpen;
  save: TauriDialogSave;
};

export type TauriListen = (event: string, handler: (event: any) => void) => Promise<() => void>;
export type TauriEmit = (event: string, payload?: any) => Promise<void> | void;

export type TauriEventApi = {
  listen: TauriListen;
  emit: TauriEmit | null;
};

function safeGetProp(obj: any, prop: string): any | undefined {
  if (!obj) return undefined;
  try {
    return obj[prop];
  } catch {
    return undefined;
  }
}

function getTauriGlobalOrNull(): any | null {
  try {
    return (globalThis as any).__TAURI__ ?? null;
  } catch {
    // Some hardened host environments (or tests) may define `__TAURI__` with a throwing getter.
    // Treat that as "unavailable" so best-effort callsites can fall back cleanly.
    return null;
  }
}

export function hasTauri(): boolean {
  return getTauriGlobalOrNull() != null;
}

export function getTauriInvokeOrNull(): TauriInvoke | null {
  return getTauriInvokeOrNullRuntime() as TauriInvoke | null;
}

export function getTauriInvokeOrThrow(): TauriInvoke {
  return getTauriInvokeOrThrowRuntime() as TauriInvoke;
}

export function hasTauriInvoke(): boolean {
  return hasTauriInvokeRuntime();
}

function getTauriAppNamespaceOrNull(): any | null {
  const tauri = getTauriGlobalOrNull();
  const plugin = safeGetProp(tauri, "plugin");
  const plugins = safeGetProp(tauri, "plugins");
  return safeGetProp(tauri, "app") ?? safeGetProp(plugin, "app") ?? safeGetProp(plugins, "app") ?? null;
}

export function getTauriAppGetNameOrNull(): TauriAppGetName | null {
  const app = getTauriAppNamespaceOrNull();
  const getName = safeGetProp(app, "getName") as (() => Promise<unknown>) | undefined;
  if (typeof getName !== "function") return null;
  return async () => {
    const resolved = await getName.call(app);
    if (typeof resolved !== "string") {
      throw new Error("Tauri app.getName returned a non-string value");
    }
    return resolved;
  };
}

export function getTauriAppGetVersionOrNull(): TauriAppGetVersion | null {
  const app = getTauriAppNamespaceOrNull();
  const getVersion = safeGetProp(app, "getVersion") as (() => Promise<unknown>) | undefined;
  if (typeof getVersion !== "function") return null;
  return async () => {
    const resolved = await getVersion.call(app);
    if (typeof resolved !== "string") {
      throw new Error("Tauri app.getVersion returned a non-string value");
    }
    return resolved;
  };
}

function getTauriDialogNamespaceOrNull(): any | null {
  const tauri = getTauriGlobalOrNull();
  const plugin = safeGetProp(tauri, "plugin");
  const plugins = safeGetProp(tauri, "plugins");
  return safeGetProp(tauri, "dialog") ?? safeGetProp(plugin, "dialog") ?? safeGetProp(plugins, "dialog") ?? null;
}

export function getTauriDialogOpenOrNull(): TauriDialogOpen | null {
  const dialog = getTauriDialogNamespaceOrNull();
  const open = safeGetProp(dialog, "open") as TauriDialogOpen | undefined;
  return typeof open === "function" ? open : null;
}

export function getTauriDialogSaveOrNull(): TauriDialogSave | null {
  const dialog = getTauriDialogNamespaceOrNull();
  const save = safeGetProp(dialog, "save") as TauriDialogSave | undefined;
  return typeof save === "function" ? save : null;
}

export function getTauriDialogConfirmOrNull(): TauriDialogConfirm | null {
  const dialog = getTauriDialogNamespaceOrNull();
  const confirm = safeGetProp(dialog, "confirm") as TauriDialogConfirm | undefined;
  return typeof confirm === "function" ? confirm : null;
}

export function getTauriDialogMessageOrNull(): TauriDialogMessage | null {
  const dialog = getTauriDialogNamespaceOrNull();
  const message = (safeGetProp(dialog, "message") ?? safeGetProp(dialog, "alert")) as
    | TauriDialogMessage
    | undefined;
  return typeof message === "function" ? message : null;
}

/**
 * Access the Tauri dialog plugin API (open/save) without a hard dependency on
 * `@tauri-apps/api`.
 *
 * Supports both legacy shapes:
 * - `__TAURI__.dialog.*`
 * - `__TAURI__.plugin.dialog.*`
 * - `__TAURI__.plugins.dialog.*`
 */
export function getTauriDialogOrNull(): TauriDialogApi | null {
  const open = getTauriDialogOpenOrNull();
  const save = getTauriDialogSaveOrNull();
  if (!open || !save) return null;
  return { open, save };
}

export function getTauriDialogOrThrow(): TauriDialogApi {
  const dialog = getTauriDialogOrNull();
  if (!dialog) {
    throw new Error("Tauri dialog API not available");
  }
  return dialog;
}

/**
 * Access the Tauri event API (listen/emit) without a hard dependency on
 * `@tauri-apps/api`.
 */
export function getTauriEventApiOrNull(): TauriEventApi | null {
  const tauri = getTauriGlobalOrNull();
  const plugin = safeGetProp(tauri, "plugin");
  const plugins = safeGetProp(tauri, "plugins");
  const eventApi = safeGetProp(tauri, "event") ?? safeGetProp(plugin, "event") ?? safeGetProp(plugins, "event") ?? null;
  const listen = safeGetProp(eventApi, "listen") as TauriListen | undefined;
  if (typeof listen !== "function") return null;
  const emit = safeGetProp(eventApi, "emit") as TauriEmit | undefined;
  return { listen, emit: typeof emit === "function" ? emit : null };
}

export function getTauriEventApiOrThrow(): TauriEventApi {
  const api = getTauriEventApiOrNull();
  if (!api) {
    throw new Error("Tauri event API not available");
  }
  return api;
}

export function hasTauriWindowApi(): boolean {
  return Boolean(safeGetProp(getTauriGlobalOrNull(), "window"));
}

/**
 * Returns true when the runtime exposes an API surface that can *produce* a window handle.
 *
 * This intentionally does not call the `getCurrent*()` accessors (some callsites only want
 * feature-detection without invoking the underlying bindings).
 */
export function hasTauriWindowHandleApi(): boolean {
  const winApi = safeGetProp(getTauriGlobalOrNull(), "window") as any;
  if (!winApi) return false;
  const hasAppWindow = (() => {
    try {
      return Boolean(winApi.appWindow);
    } catch {
      return false;
    }
  })();
  const getCurrentWebviewWindow = safeGetProp(winApi, "getCurrentWebviewWindow");
  const getCurrentWindow = safeGetProp(winApi, "getCurrentWindow");
  const getCurrent = safeGetProp(winApi, "getCurrent");
  return (
    typeof getCurrentWebviewWindow === "function" ||
    typeof getCurrentWindow === "function" ||
    typeof getCurrent === "function" ||
    hasAppWindow
  );
}

export function getTauriWindowHandleOrNull(): any | null {
  const winApi = safeGetProp(getTauriGlobalOrNull(), "window") as any;
  if (!winApi) return null;

  // Tauri v2 exposes window handles via helper functions; keep this flexible since
  // we intentionally avoid a hard dependency on `@tauri-apps/api`.
  const tryCall = (fn: unknown): any | null => {
    if (typeof fn !== "function") return null;
    try {
      return (fn as (...args: any[]) => any).call(winApi);
    } catch {
      return null;
    }
  };

  const handle =
    tryCall(safeGetProp(winApi, "getCurrentWebviewWindow")) ??
    tryCall(safeGetProp(winApi, "getCurrentWindow")) ??
    tryCall(safeGetProp(winApi, "getCurrent")) ??
    (() => {
      try {
        return winApi.appWindow ?? null;
      } catch {
        return null;
      }
    })();

  return handle ?? null;
}

export function getTauriWindowHandleOrThrow(): any {
  const winApi = safeGetProp(getTauriGlobalOrNull(), "window") as any;
  if (!winApi) {
    throw new Error("Tauri window API not available");
  }

  const handle = getTauriWindowHandleOrNull();
  if (!handle) {
    throw new Error("Tauri window handle not available");
  }
  return handle;
}
