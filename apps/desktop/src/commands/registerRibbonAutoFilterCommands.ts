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
  /**
   * Clear AutoFilter criteria (show all rows) while keeping AutoFilter enabled.
   */
  clear: () => void | Promise<void>;
  /**
   * Recompute and reapply the current AutoFilter criteria.
   */
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
    description: "Clear AutoFilter criteria (show all rows)",
    keywords: ["clear", "filter", "auto filter", "autofilter"],
    // Hide from the command palette when AutoFilter is not enabled for the active sheet.
    when: "spreadsheet.hasAutoFilter == true",
  });

  registerIfMissing(commandRegistry, "data.sortFilter.reapply", "Reapply", () => withHandlers((handlers) => handlers.reapply()), {
    category,
    icon: null,
    description: "Reapply the current AutoFilter",
    keywords: ["reapply", "filter", "auto filter", "autofilter"],
    // Hide from the command palette when AutoFilter is not enabled for the active sheet.
    when: "spreadsheet.hasAutoFilter == true",
  });

  // Ribbon-only alias for the canonical Clear command. Keep it registered so the ribbon isn't
  // auto-disabled, but hide it from the command palette to avoid duplicate entries.
  registerIfMissing(commandRegistry, "data.sortFilter.advanced.clearFilter", "Clear Filter", () => commandRegistry.executeCommand("data.sortFilter.clear"), {
    category,
    icon: null,
    when: "false",
  });
} 
