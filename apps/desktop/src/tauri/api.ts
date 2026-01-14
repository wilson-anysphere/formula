export type TauriDialogOpen = (options?: Record<string, unknown>) => Promise<string | string[] | null>;
export type TauriDialogSave = (options?: Record<string, unknown>) => Promise<string | null>;
export type TauriDialogConfirm = (message: string, options?: Record<string, unknown>) => Promise<boolean>;
export type TauriDialogMessage = (message: string, options?: Record<string, unknown>) => Promise<void>;

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

function getTauriDialogNamespaceOrNull(): any | null {
  const tauri = getTauriGlobalOrNull();
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

export function getTauriDialogConfirmOrNull(): TauriDialogConfirm | null {
  const dialog = getTauriDialogNamespaceOrNull();
  const confirm = dialog?.confirm as TauriDialogConfirm | undefined;
  return typeof confirm === "function" ? confirm : null;
}

export function getTauriDialogMessageOrNull(): TauriDialogMessage | null {
  const dialog = getTauriDialogNamespaceOrNull();
  const message = (dialog?.message ?? dialog?.alert) as TauriDialogMessage | undefined;
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
  return Boolean(getTauriGlobalOrNull()?.window);
}

/**
 * Returns true when the runtime exposes an API surface that can *produce* a window handle.
 *
 * This intentionally does not call the `getCurrent*()` accessors (some callsites only want
 * feature-detection without invoking the underlying bindings).
 */
export function hasTauriWindowHandleApi(): boolean {
  const winApi = getTauriGlobalOrNull()?.window as any;
  if (!winApi) return false;
  const hasAppWindow = (() => {
    try {
      return Boolean(winApi.appWindow);
    } catch {
      return false;
    }
  })();
  return (
    typeof winApi.getCurrentWebviewWindow === "function" ||
    typeof winApi.getCurrentWindow === "function" ||
    typeof winApi.getCurrent === "function" ||
    hasAppWindow
  );
}

export function getTauriWindowHandleOrNull(): any | null {
  const winApi = getTauriGlobalOrNull()?.window as any;
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
    tryCall(winApi.getCurrentWebviewWindow) ??
    tryCall(winApi.getCurrentWindow) ??
    tryCall(winApi.getCurrent) ??
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
  const winApi = getTauriGlobalOrNull()?.window as any;
  if (!winApi) {
    throw new Error("Tauri window API not available");
  }

  const handle = getTauriWindowHandleOrNull();
  if (!handle) {
    throw new Error("Tauri window handle not available");
  }
  return handle;
}
