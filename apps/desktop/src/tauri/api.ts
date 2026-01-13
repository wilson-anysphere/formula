export type TauriDialogOpen = (options?: Record<string, unknown>) => Promise<string | string[] | null>;
export type TauriDialogSave = (options?: Record<string, unknown>) => Promise<string | null>;

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

/**
 * Access the Tauri dialog plugin API (open/save) without a hard dependency on
 * `@tauri-apps/api`.
 *
 * Supports both legacy shapes:
 * - `__TAURI__.dialog.*`
 * - `__TAURI__.plugin.dialog.*`
 */
export function getTauriDialogOrNull(): TauriDialogApi | null {
  const tauri = (globalThis as any).__TAURI__ as any;
  const dialog = tauri?.dialog ?? tauri?.plugin?.dialog ?? null;
  const open = dialog?.open as TauriDialogOpen | undefined;
  const save = dialog?.save as TauriDialogSave | undefined;
  if (typeof open !== "function" || typeof save !== "function") return null;
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
  const eventApi = (globalThis as any).__TAURI__?.event as any;
  const listen = eventApi?.listen as TauriListen | undefined;
  if (typeof listen !== "function") return null;
  const emit = eventApi?.emit as TauriEmit | undefined;
  return { listen, emit: typeof emit === "function" ? emit : null };
}

export function getTauriEventApiOrThrow(): TauriEventApi {
  const api = getTauriEventApiOrNull();
  if (!api) {
    throw new Error("Tauri event API not available");
  }
  return api;
}

export function getTauriWindowHandleOrNull(): any | null {
  const winApi = (globalThis as any).__TAURI__?.window as any;
  if (!winApi) return null;

  // Tauri v2 exposes window handles via helper functions; keep this flexible since
  // we intentionally avoid a hard dependency on `@tauri-apps/api`.
  const handle =
    (typeof winApi.getCurrentWebviewWindow === "function" ? winApi.getCurrentWebviewWindow() : null) ??
    (typeof winApi.getCurrentWindow === "function" ? winApi.getCurrentWindow() : null) ??
    (typeof winApi.getCurrent === "function" ? winApi.getCurrent() : null) ??
    winApi.appWindow ??
    null;

  return handle ?? null;
}

export function getTauriWindowHandleOrThrow(): any {
  const winApi = (globalThis as any).__TAURI__?.window as any;
  if (!winApi) {
    throw new Error("Tauri window API not available");
  }

  const handle = getTauriWindowHandleOrNull();
  if (!handle) {
    throw new Error("Tauri window handle not available");
  }
  return handle;
}

