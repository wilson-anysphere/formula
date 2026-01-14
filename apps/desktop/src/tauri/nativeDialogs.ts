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

import { showQuickPick } from "../extensions/ui.js";

import { getTauriDialogConfirmOrNull, getTauriDialogMessageOrNull } from "./api";

function isNativeDialogFn(fn: unknown): boolean {
  if (typeof fn !== "function") return false;
  try {
    // Native browser implementations typically include `[native code]` in their
    // Function#toString output. Avoid calling those because they block the UI
    // thread. (In unit tests, `window.confirm` is often stubbed with `vi.fn()`
    // which does not include `[native code]`, and is safe to call synchronously.)
    const text = Function.prototype.toString.call(fn);
    return /\[native code\]/.test(text);
  } catch {
    // Be conservative: if we cannot inspect it, assume it's native.
    return true;
  }
}

function getWindowConfirm(): ((message: string) => boolean) | null {
  const win = (globalThis as any).window as unknown;
  const confirmFn = (win as any)?.confirm as ((message: string) => boolean) | undefined;
  if (typeof confirmFn !== "function") return null;
  // Avoid calling the browser-native confirm dialog (blocks the UI thread).
  if (isNativeDialogFn(confirmFn)) return null;
  return (message: string) => confirmFn.call(win, message);
}

function getWindowAlert(): ((message: string) => void) | null {
  const win = (globalThis as any).window as unknown;
  const alertFn = (win as any)?.alert as ((message: string) => void) | undefined;
  if (typeof alertFn !== "function") return null;
  // Avoid calling the browser-native alert dialog (blocks the UI thread).
  if (isNativeDialogFn(alertFn)) return null;
  return (message: string) => alertFn.call(win, message);
}

export async function confirm(message: string, opts: ConfirmDialogOptions = {}): Promise<boolean> {
  const { fallbackValue, ...dialogOpts } = opts;
  const fallback = fallbackValue ?? false;

  const tauriConfirm = getTauriDialogConfirmOrNull();
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

  // Web fallback: use a non-blocking <dialog>-based quick pick instead of the
  // browser-native `window.confirm` dialog.
  if (typeof document !== "undefined" && document.body) {
    const okLabel = dialogOpts.okLabel ?? "OK";
    const cancelLabel = dialogOpts.cancelLabel ?? "Cancel";
    const prefix = dialogOpts.title ? `${dialogOpts.title}: ` : "";
    const choice = await showQuickPick(
      [
        { label: okLabel, value: true },
        { label: cancelLabel, value: false },
      ],
      { placeHolder: `${prefix}${message}` },
    );
    return choice === true;
  }

  // Non-browser environment (e.g. unit tests) without a stubbed `window.confirm`.
  // Default to `false` to avoid accidentally confirming destructive actions.
  return fallback;
}

export async function alert(message: string, opts: AlertDialogOptions = {}): Promise<void> {
  const tauriMessage = getTauriDialogMessageOrNull();
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
    return;
  }

  // Web fallback: use a non-blocking <dialog>-based quick pick instead of the
  // browser-native `window.alert` dialog.
  if (typeof document !== "undefined" && document.body) {
    const prefix = opts.title ? `${opts.title}: ` : "";
    const dialogMessage = `${prefix}${message}`;
    const okLabel = "OK";
    // Reuse the quick-pick dialog styling so we don't need a dedicated alert UI.
    await showQuickPick([{ label: okLabel, value: true }], { placeHolder: dialogMessage });
  }
}
