import { showToast } from "../extensions/ui.js";
import { t } from "../i18n/index.js";

import * as nativeDialogs from "./nativeDialogs";

export type BeforeQuitHook = () => Promise<void> | void;

export type AppQuitHandlers = {
  /**
   * Returns whether there are unsaved changes that should gate quitting/restarting.
   */
  isDirty: () => boolean;
  /**
   * Best-effort Workbook_BeforeClose handling (macros + any workbook-sync flush it needs).
   */
  runWorkbookBeforeClose?: () => Promise<void>;
  /**
   * Flush any remaining workbook-sync operations that should complete before we exit.
   */
  drainBackendSync: () => Promise<void>;
  /**
   * Perform the final quit operation. (On desktop this invokes the `quit_app` backend command,
   * which hard-exits the process.)
   */
  quitApp: () => Promise<void> | void;
  /**
   * Perform the final restart operation.
   *
   * For updater-driven restarts, prefer this over `quitApp` so the desktop shell can use
   * Tauri-managed restart/exit semantics (`restart_app`).
   */
  restartApp?: () => Promise<void> | void;
};

export type RequestAppQuitOptions = {
  /**
   * Runs only after the user has confirmed discarding unsaved changes (or the document is clean)
   * and after backend sync has been drained.
   */
  beforeQuit?: BeforeQuitHook;
  /**
   * Shown if `beforeQuit` rejects; the quit/restart is aborted.
   */
  beforeQuitErrorToast?: string;
  /**
   * Overrides the default unsaved-changes confirm copy.
   */
  dirtyConfirmMessage?: string;
};

let quitHandlers: AppQuitHandlers | null = null;
let quitInFlight = false;

export function registerAppQuitHandlers(next: AppQuitHandlers | null): void {
  quitHandlers = next;
}

async function requestAppExit(
  exit: (handlers: AppQuitHandlers) => Promise<void> | void,
  options: RequestAppQuitOptions = {},
): Promise<boolean> {
  const handlers = quitHandlers;
  if (!handlers) {
    console.warn("requestAppQuit/requestAppRestart called before quit handlers were registered");
    return false;
  }

  if (quitInFlight) return false;
  quitInFlight = true;

  try {
    if (handlers.runWorkbookBeforeClose) {
      try {
        await handlers.runWorkbookBeforeClose();
      } catch (err) {
        // Don't block quitting on a macro crash; match the existing tray-quit behavior.
        console.warn("Workbook_BeforeClose event macro failed:", err);
      }
    }

    if (handlers.isDirty()) {
      const discard = await nativeDialogs.confirm(
        options.dirtyConfirmMessage ?? t("prompt.unsavedChangesDiscardConfirm"),
      );
      if (!discard) return false;
    }

    // Best-effort flush of any microtask-batched workbook edits before quitting/restarting.
    await new Promise<void>((resolve) => queueMicrotask(resolve));
    await handlers.drainBackendSync();

    if (options.beforeQuit) {
      try {
        await options.beforeQuit();
      } catch (err) {
        const message = options.beforeQuitErrorToast ?? `Failed to quit: ${String(err)}`;
        try {
          showToast(message, "error");
        } catch {
          console.error(message);
        }
        return false;
      }
    }

    await exit(handlers);
    return true;
  } finally {
    // If `quitApp` hard-exits, this `finally` never runs, which is fine.
    quitInFlight = false;
  }
}

export async function requestAppQuit(options: RequestAppQuitOptions = {}): Promise<boolean> {
  return requestAppExit((handlers) => handlers.quitApp(), options);
}

export async function requestAppRestart(options: {
  beforeQuit: BeforeQuitHook;
  beforeQuitErrorToast?: string;
  dirtyConfirmMessage?: string;
}): Promise<boolean> {
  return requestAppExit(
    (handlers) => (handlers.restartApp ?? handlers.quitApp)(),
    {
      beforeQuit: options.beforeQuit,
      beforeQuitErrorToast: options.beforeQuitErrorToast ?? t("updater.restartFailed"),
      dirtyConfirmMessage: options.dirtyConfirmMessage,
    },
  );
}
