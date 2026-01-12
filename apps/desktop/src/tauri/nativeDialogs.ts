export type ConfirmDialogOptions = {
  title?: string;
  okLabel?: string;
  cancelLabel?: string;
  /**
   * When no dialog API is available (e.g. non-browser env or jsdom "Not implemented"
   * stubs), return this value instead of the default `false`.
   *
   * This is useful for non-destructive prompts whose safest behavior is to proceed
   * when prompting isn't possible.
   */
  fallbackValue?: boolean;
};

export type AlertDialogOptions = {
  title?: string;
};

type TauriDialogConfirm = (message: string, options?: Record<string, unknown>) => Promise<boolean>;
type TauriDialogMessage = (message: string, options?: Record<string, unknown>) => Promise<void>;

function getTauriDialogApi():
  | {
      confirm?: TauriDialogConfirm;
      message?: TauriDialogMessage;
      alert?: TauriDialogMessage;
    }
  | null {
  const tauri = (globalThis as any).__TAURI__;
  const dialog = (tauri?.dialog ?? tauri?.plugin?.dialog) as unknown;
  if (!dialog || typeof dialog !== "object") return null;
  return dialog as any;
}

function getWindowConfirm(): ((message: string) => boolean) | null {
  const win = (globalThis as any).window as unknown;
  const confirmFn = (win as any)?.confirm as ((message: string) => boolean) | undefined;
  if (typeof confirmFn !== "function") return null;
  return (message: string) => confirmFn.call(win, message);
}

function getWindowAlert(): ((message: string) => void) | null {
  const win = (globalThis as any).window as unknown;
  const alertFn = (win as any)?.alert as ((message: string) => void) | undefined;
  if (typeof alertFn !== "function") return null;
  return (message: string) => alertFn.call(win, message);
}

export async function confirm(message: string, opts: ConfirmDialogOptions = {}): Promise<boolean> {
  const { fallbackValue, ...dialogOpts } = opts;
  const fallback = fallbackValue ?? false;

  const dialog = getTauriDialogApi();
  const tauriConfirm = dialog?.confirm;
  if (typeof tauriConfirm === "function") {
    try {
      return await tauriConfirm(message, dialogOpts);
    } catch {
      // Fall through to `window.confirm` below if the Tauri call fails.
    }
  }

  const windowConfirm = getWindowConfirm();
  if (windowConfirm) {
    try {
      return windowConfirm(message);
    } catch {
      // Some test/host environments (e.g. jsdom) define `window.confirm` but throw
      // a "Not implemented" error. Treat that the same as an unavailable API.
      return fallback;
    }
  }

  // Non-browser environment (e.g. unit tests) without a stubbed `window.confirm`.
  // Default to `false` to avoid accidentally confirming destructive actions.
  return fallback;
}

export async function alert(message: string, opts: AlertDialogOptions = {}): Promise<void> {
  const dialog = getTauriDialogApi();
  const tauriMessage = dialog?.message ?? dialog?.alert;
  if (typeof tauriMessage === "function") {
    try {
      await tauriMessage(message, opts);
      return;
    } catch {
      // Fall through to `window.alert` below if the Tauri call fails.
    }
  }

  const windowAlert = getWindowAlert();
  if (windowAlert) {
    try {
      windowAlert(message);
    } catch {
      // Some test/host environments (e.g. jsdom) define `window.alert` but throw
      // a "Not implemented" error. Ignore and continue.
    }
  }
}
