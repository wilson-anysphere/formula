import type { CommandRegistry } from "./commandRegistry.js";
import type { ResolvedMenuItem } from "./contextMenus.js";

export type ContextMenuModelItem =
  | { kind: "separator" }
  | { kind: "command"; commandId: string; label: string; enabled: boolean };

function groupName(group: string | null): string {
  if (!group) return "";
  const raw = String(group);
  const [name] = raw.split("@", 2);
  return name ?? raw;
}

function formatMenuLabel(
  commandId: string,
  commandRegistry: Pick<CommandRegistry, "getCommand">,
): string {
  const command = commandRegistry.getCommand(commandId);
  if (!command) return commandId;
  return command.category ? `${command.category}: ${command.title}` : command.title;
}

/**
 * Converts resolved extension menu items into a renderable model.
 *
 * We intentionally preserve `resolveMenuItems(...)` sorting semantics; this function
 * only inserts separators when the menu group changes between consecutive items.
 */
export function buildContextMenuModel(
  items: ResolvedMenuItem[],
  commandRegistry: Pick<CommandRegistry, "getCommand">,
): ContextMenuModelItem[] {
  const model: ContextMenuModelItem[] = [];
  let lastGroup: string | undefined;

  for (const item of items) {
    const nextGroup = groupName(item.group);
    if (lastGroup !== undefined && nextGroup !== lastGroup) {
      model.push({ kind: "separator" });
    }
    lastGroup = nextGroup;

    model.push({
      kind: "command",
      commandId: item.command,
      enabled: item.enabled,
      label: formatMenuLabel(item.command, commandRegistry),
    });
  }

  return model;
}

