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
   * Toggle buttons should invoke `onToggle(id, pressed)` only.
   *
   * For backwards compatibility with legacy hosts that invoke `onCommand(id)` immediately
   * after `onToggle`, the bridge will suppress the follow-up `onCommand` callback to avoid
   * double-executing toggle actions / showing duplicate fallback toasts.
   *
   * To opt into the normal unknown-command fallback behavior for unknown toggles (i.e. call
   * `onUnknownCommand`, or show the default toast when omitted), return `false` synchronously.
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
   * Some hosts historically invoked `onCommand` immediately after `onToggle` for toggle buttons.
   *
   * Keep a short-lived suppression map so we can ignore the follow-up `onCommand` call(s) and avoid
   * double execution.
   */
  const pendingToggleSuppress = new Map<string, { count: number; cleanupToken: number }>();
  let nextToggleSuppressCleanupToken = 0;
  const scheduleMicrotask =
    typeof queueMicrotask === "function"
      ? queueMicrotask
      : (cb: () => void) => {
          void Promise.resolve()
            .then(cb)
            .catch(() => {
              // Best-effort: avoid unhandled rejections from the microtask fallback scheduler.
            });
        };
  // Prefer a macrotask boundary so the suppression survives follow-up `onCommand` calls that are
  // scheduled asynchronously (microtasks or timers). When `setTimeout` is available, we use a
  // nested timer so cleanup always runs after any same-turn `setTimeout(..., 0)` follow-ups,
  // regardless of how many microtask turns the host uses before scheduling its timer.
  const scheduleToggleSuppressCleanup = (cb: () => void): void => {
    if (typeof setTimeout !== "function") {
      // Queue two microtasks so host follow-up `onCommand` microtasks run before cleanup.
      scheduleMicrotask(() => scheduleMicrotask(cb));
      return;
    }
    void setTimeout(() => void setTimeout(cb, 0), 0);
  };
  const markToggleHandled = (commandId: string): void => {
    const token = (nextToggleSuppressCleanupToken += 1);
    const entry = pendingToggleSuppress.get(commandId);
    if (entry) {
      entry.count += 1;
      entry.cleanupToken = token;
    } else {
      pendingToggleSuppress.set(commandId, { count: 1, cleanupToken: token });
    }
    // Ensure we don't leak memory if the host only calls `onToggle` (tests/custom hosts).
    scheduleToggleSuppressCleanup(() => {
      const current = pendingToggleSuppress.get(commandId);
      if (!current) return;
      // Ignore stale cleanup callbacks when additional toggles happen before cleanup runs.
      if (current.cleanupToken !== token) return;
      pendingToggleSuppress.delete(commandId);
    });
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

  const handleUnknownCommand = async (commandId: string): Promise<void> => {
    if (onUnknownCommand) {
      await onUnknownCommand(commandId);
      return;
    }
    if (commandId.startsWith("file.")) {
      safeShowToast(`File command not implemented: ${commandId}`);
    } else {
      safeShowToast(`Ribbon: ${commandId}`);
    }
  };

  return {
    onCommand: (commandId: string) => {
      const pending = pendingToggleSuppress.get(commandId);
      if (pending) {
        pending.count -= 1;
        if (pending.count <= 0) {
          pendingToggleSuppress.delete(commandId);
        }
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

        await handleUnknownCommand(commandId);
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
          if (handled === false) {
            // Fall through to the `onUnknownCommand` behavior directly so unknown toggles are not
            // silent in hosts that only invoke `onToggle` (no follow-up `onCommand`).
            markToggleHandled(commandId);
            await handleUnknownCommand(commandId);
            return;
          }
          markToggleHandled(commandId);
          await handled;
          return;
        }

        // Fallback: in modern Ribbon builds, toggle buttons do *not* invoke `onCommand`.
        // Mirror the unknown-command behavior here so unregistered toggle ids still surface
        // a toast (and so older hosts won't double-toast thanks to `markToggleHandled`).
        markToggleHandled(commandId);
        await handleUnknownCommand(commandId);
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
