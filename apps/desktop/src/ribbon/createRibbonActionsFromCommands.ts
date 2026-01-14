import type { CommandContribution, CommandRegistry } from "../extensions/commandRegistry.js";
import { showToast } from "../extensions/ui.js";
import type { RibbonActions, RibbonFileActions } from "./ribbonSchema.js";

type RibbonCommandOverride = () => void | Promise<void>;
type RibbonToggleOverride = (pressed: boolean) => void | Promise<void>;
type RibbonUnknownToggleResult = boolean | void | Promise<boolean | void>;

export function createRibbonActionsFromCommands(params: {
  commandRegistry: CommandRegistry;
  onCommandError?: (commandId: string, err: unknown) => void;
  /**
   * Optional hook that runs before executing a registered command (e.g. lazy-load
   * extensions before executing extension-contributed commands).
   */
  onBeforeExecuteCommand?: (commandId: string, source: CommandContribution["source"]) => void | Promise<void>;
  /**
   * Special-case handlers for commands that should not (or cannot) be dispatched
   * through the CommandRegistry.
   */
  commandOverrides?: Record<string, RibbonCommandOverride>;
  /**
   * Special-case handlers for toggle commands that should not (or cannot) be
   * dispatched through the CommandRegistry.
   */
  toggleOverrides?: Record<string, RibbonToggleOverride>;
  /**
   * Fallback handler for unregistered commands. If omitted, the bridge will show
   * a toast (best-effort) and log a warning.
   */
  onUnknownCommand?: (commandId: string) => void | Promise<void>;
  /**
   * Fallback handler for unregistered toggle commands.
   *
   * NOTE: Ribbon toggles invoke both `onToggle` and `onCommand`. When this
   * handler returns anything other than `false`, the subsequent `onCommand`
   * callback will be suppressed to avoid double-executing toggle actions.
   *
   * To opt into the normal `onCommand` fallback behavior for unknown toggles
   * (e.g. show the default toast / call `onUnknownCommand`), return `false`
   * synchronously.
   */
  onUnknownToggle?: (commandId: string, pressed: boolean) => RibbonUnknownToggleResult;
}): RibbonActions {
  const {
    commandRegistry,
    onCommandError,
    onBeforeExecuteCommand,
    commandOverrides = {},
    toggleOverrides = {},
    onUnknownCommand,
    onUnknownToggle,
  } = params;

  const safeShowToast = (message: string): void => {
    try {
      showToast(message);
      return;
    } catch {
      // `showToast` depends on DOM globals and a #toast-root. Fall back to a console warning
      // so non-UI contexts (SSR/tests) still surface missing command wiring.
    }
    console.warn(message);
  };

  const reportError = (commandId: string, err: unknown): void => {
    try {
      onCommandError?.(commandId, err);
    } catch {
      // Avoid cascading failures (e.g. toast root missing) causing unhandled rejections.
    }
  };

  /**
   * Ribbon toggle buttons invoke both `onToggle` and `onCommand`. The default
   * bridge behavior is to handle toggle semantics via `onToggle` (so callers can
   * receive the pressed state) and suppress the follow-up `onCommand` callback.
   */
  const pendingToggleSuppress = new Set<string>();
  const scheduleMicrotask =
    typeof queueMicrotask === "function" ? queueMicrotask : (cb: () => void) => Promise.resolve().then(cb);
  const markToggleHandled = (commandId: string): void => {
    pendingToggleSuppress.add(commandId);
    // Ensure we don't leak memory if the host only calls `onToggle` (tests/custom hosts).
    scheduleMicrotask(() => pendingToggleSuppress.delete(commandId));
  };

  const run = (commandId: string, fn: () => void | Promise<void>): void => {
    void (async () => {
      try {
        await fn();
      } catch (err) {
        reportError(commandId, err);
      }
    })();
  };

  return {
    onCommand: (commandId: string) => {
      if (pendingToggleSuppress.has(commandId)) {
        pendingToggleSuppress.delete(commandId);
        return;
      }

      run(commandId, async () => {
        const override = commandOverrides[commandId];
        if (override) {
          await override();
          return;
        }

        const registered = commandRegistry.getCommand(commandId);
        if (registered) {
          if (onBeforeExecuteCommand) {
            await onBeforeExecuteCommand(commandId, registered.source);
          }
          await commandRegistry.executeCommand(commandId);
          return;
        }

        if (onUnknownCommand) {
          await onUnknownCommand(commandId);
          return;
        }

        if (commandId.startsWith("file.")) {
          safeShowToast(`File command not implemented: ${commandId}`);
        } else {
          safeShowToast(`Ribbon: ${commandId}`);
        }
      });
    },
    onToggle: (commandId: string, pressed: boolean) => {
      run(commandId, async () => {
        const override = toggleOverrides[commandId];
        if (override) {
          markToggleHandled(commandId);
          await override(pressed);
          return;
        }

        const registered = commandRegistry.getCommand(commandId);
        if (registered) {
          markToggleHandled(commandId);
          if (onBeforeExecuteCommand) {
            await onBeforeExecuteCommand(commandId, registered.source);
          }
          await commandRegistry.executeCommand(commandId, pressed);
          return;
        }

        if (onUnknownToggle) {
          const handled = onUnknownToggle(commandId, pressed);
          // `Ribbon` calls `onCommand` immediately after `onToggle`, so we must decide
          // whether to suppress synchronously. Returning `false` opts out of suppression.
          if (handled !== false) {
            markToggleHandled(commandId);
          }
          await handled;
        }
        // If there's no handler for this toggle, intentionally do nothing: Ribbon toggles
        // also invoke `onCommand`, which will fall back to `onUnknownCommand` and/or show
        // the default toast for unknown commands.
      });
    },
  };
}

export function createRibbonFileActionsFromCommands(params: {
  commandRegistry: CommandRegistry;
  onCommandError?: (commandId: string, err: unknown) => void;
  commandIds: {
    newWorkbook?: string;
    openWorkbook?: string;
    saveWorkbook?: string;
    saveWorkbookAs?: string;
    toggleAutoSave?: string;
    versionHistory?: string;
    branchManager?: string;
    print?: string;
    printPreview?: string;
    pageSetup?: string;
    closeWindow?: string;
    quit?: string;
  };
}): RibbonFileActions {
  const { commandRegistry, onCommandError, commandIds } = params;

  const reportError = (commandId: string, err: unknown): void => {
    try {
      onCommandError?.(commandId, err);
    } catch {
      // Avoid cascading failures (e.g. toast root missing) causing unhandled rejections.
    }
  };

  const run = (commandId: string, fn: () => void | Promise<void>): void => {
    void (async () => {
      try {
        await fn();
      } catch (err) {
        reportError(commandId, err);
      }
    })();
  };

  const command = (commandId: string | undefined): (() => void) | undefined => {
    if (!commandId) return undefined;
    return () =>
      run(commandId, async () => {
        await commandRegistry.executeCommand(commandId);
      });
  };

  return {
    newWorkbook: command(commandIds.newWorkbook),
    openWorkbook: command(commandIds.openWorkbook),
    saveWorkbook: command(commandIds.saveWorkbook),
    saveWorkbookAs: command(commandIds.saveWorkbookAs),
    toggleAutoSave: commandIds.toggleAutoSave
      ? (enabled) =>
          run(commandIds.toggleAutoSave!, async () => {
            await commandRegistry.executeCommand(commandIds.toggleAutoSave!, enabled);
          })
      : undefined,
    versionHistory: command(commandIds.versionHistory),
    branchManager: command(commandIds.branchManager),
    print: command(commandIds.print),
    printPreview: command(commandIds.printPreview),
    pageSetup: command(commandIds.pageSetup),
    closeWindow: command(commandIds.closeWindow),
    quit: command(commandIds.quit),
  };
}
