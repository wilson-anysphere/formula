import type { CommandRegistry } from "../extensions/commandRegistry.js";

export type RibbonAutoFilterCommandHandlers = {
  /**
   * Toggle (or explicitly set) the ribbon AutoFilter state for the active sheet.
   *
   * When `pressed` is provided (ribbon toggle), implementations should enable/disable filtering
   * deterministically. When omitted (command palette / keybinding), implementations can apply an
   * Excel-style toggle.
   */
  toggle: (pressed?: boolean) => void | Promise<void>;
  clear: () => void | Promise<void>;
  reapply: () => void | Promise<void>;
};

function registerIfMissing(
  commandRegistry: CommandRegistry,
  commandId: string,
  title: string,
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  run: (...args: any[]) => void | Promise<void>,
  options: {
    category: string;
    icon?: string | null;
    description?: string | null;
    keywords?: string[] | null;
    when?: string | null;
  },
): void {
  // Avoid silent overwrites when multiple registration surfaces call this helper.
  if (commandRegistry.getCommand(commandId)) return;
  commandRegistry.registerBuiltinCommand(commandId, title, run, options);
}

export function registerRibbonAutoFilterCommands(params: {
  commandRegistry: CommandRegistry;
  /**
   * Resolve the host-provided AutoFilter implementation. When null, the commands become no-ops
   * (callers can display a toast from within `getHandlers`).
   */
  getHandlers: () => RibbonAutoFilterCommandHandlers | null;
  isEditing: () => boolean;
  category: string;
}): void {
  const { commandRegistry, getHandlers, isEditing, category } = params;

  const withHandlers = async (fn: (handlers: RibbonAutoFilterCommandHandlers) => void | Promise<void>): Promise<void> => {
    if (isEditing()) return;
    const handlers = getHandlers();
    if (!handlers) return;
    await fn(handlers);
  };

  registerIfMissing(commandRegistry, "data.sortFilter.filter", "Filter", (pressed?: boolean) => withHandlers((handlers) => handlers.toggle(pressed)), {
    category,
    icon: null,
    description: "Toggle AutoFilter for the active sheet",
    keywords: ["filter", "auto filter", "autofilter", "sort & filter"],
  });

  registerIfMissing(commandRegistry, "data.sortFilter.clear", "Clear", () => withHandlers((handlers) => handlers.clear()), {
    category,
    icon: null,
    description: "Clear the current AutoFilter criteria",
    keywords: ["clear", "filter", "auto filter", "autofilter"],
  });

  registerIfMissing(commandRegistry, "data.sortFilter.reapply", "Reapply", () => withHandlers((handlers) => handlers.reapply()), {
    category,
    icon: null,
    description: "Reapply the current AutoFilter",
    keywords: ["reapply", "filter", "auto filter", "autofilter"],
  });

  registerIfMissing(commandRegistry, "data.sortFilter.advanced.clearFilter", "Clear Filter", () => withHandlers((handlers) => handlers.clear()), {
    category,
    icon: null,
    description: "Clear the current AutoFilter criteria",
    keywords: ["clear filter", "filter", "auto filter", "autofilter"],
  });
}
