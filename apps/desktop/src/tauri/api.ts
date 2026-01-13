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

function getTauriDialogNamespaceOrNull(): any | null {
  const tauri = (globalThis as any).__TAURI__ as any;
  return tauri?.dialog ?? tauri?.plugin?.dialog ?? tauri?.plugins?.dialog ?? null;
}

export function getTauriDialogOpenOrNull(): TauriDialogOpen | null {
  const dialog = getTauriDialogNamespaceOrNull();
  const open = dialog?.open as TauriDialogOpen | undefined;
  return typeof open === "function" ? open : null;
}

export function getTauriDialogSaveOrNull(): TauriDialogSave | null {
  const dialog = getTauriDialogNamespaceOrNull();
  const save = dialog?.save as TauriDialogSave | undefined;
  return typeof save === "function" ? save : null;
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
  const tauri = (globalThis as any).__TAURI__ as any;
  const eventApi = tauri?.event ?? tauri?.plugin?.event ?? tauri?.plugins?.event ?? null;
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

export function hasTauriWindowApi(): boolean {
  return Boolean((globalThis as any).__TAURI__?.window);
}

/**
 * Returns true when the runtime exposes an API surface that can *produce* a window handle.
 *
 * This intentionally does not call the `getCurrent*()` accessors (some callsites only want
 * feature-detection without invoking the underlying bindings).
 */
export function hasTauriWindowHandleApi(): boolean {
  const winApi = (globalThis as any).__TAURI__?.window as any;
  if (!winApi) return false;
  return (
    typeof winApi.getCurrentWebviewWindow === "function" ||
    typeof winApi.getCurrentWindow === "function" ||
    typeof winApi.getCurrent === "function" ||
    Boolean(winApi.appWindow)
  );
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
