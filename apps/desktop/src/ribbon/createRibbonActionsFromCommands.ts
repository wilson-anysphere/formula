import type { CommandContribution, CommandRegistry } from "../extensions/commandRegistry.js";
import { showToast } from "../extensions/ui.js";
import type { RibbonActions } from "./ribbonSchema.js";

type RibbonCommandOverride = () => void | Promise<void>;
type RibbonToggleOverride = (pressed: boolean) => void | Promise<void>;

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
   * NOTE: Ribbon toggles invoke both `onToggle` and `onCommand`. When provided,
   * this handler is treated as "handled" and the subsequent `onCommand` call
   * will be suppressed.
   */
  onUnknownToggle?: (commandId: string, pressed: boolean) => void | Promise<void>;
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
          markToggleHandled(commandId);
          await onUnknownToggle(commandId, pressed);
        }
        // If there's no handler for this toggle, intentionally do nothing: Ribbon toggles
        // also invoke `onCommand`, which will fall back to `onUnknownCommand` and/or show
        // the default toast for unknown commands.
      });
    },
  };
}
