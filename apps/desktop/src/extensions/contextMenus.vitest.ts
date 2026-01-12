import { describe, expect, it } from "vitest";

import { resolveMenuItems, type ContributedMenuItem } from "./contextMenus.js";
import type { ResolvedMenuItem } from "./contextMenus.js";
import { buildContextMenuModel } from "./contextMenuModel.js";
import type { ContextMenuModelItem } from "./contextMenuModel.js";
import type { CommandContribution } from "./commandRegistry.js";

describe("context menu resolution", () => {
  it("evaluates when clauses and sorts by group/order", () => {
    const items: ContributedMenuItem[] = [
      { extensionId: "ext", command: "cmd.two", when: "cellHasValue", group: "extensions@2" },
      { extensionId: "ext", command: "cmd.one", when: "cellHasValue", group: "extensions@1" },
      { extensionId: "ext", command: "cmd.hidden", when: "!cellHasValue", group: "extensions@3" },
      { extensionId: "ext", command: "cmd.ungrouped", when: null, group: null },
    ];

    const resolved = resolveMenuItems(items, (key) => ({ cellHasValue: true }[key]));

    expect(resolved.find((i) => i.command === "cmd.one")?.enabled).toBe(true);
    expect(resolved.find((i) => i.command === "cmd.two")?.enabled).toBe(true);
    expect(resolved.find((i) => i.command === "cmd.hidden")?.enabled).toBe(false);
    expect(resolved.find((i) => i.command === "cmd.ungrouped")?.enabled).toBe(true);

    const order = resolved.map((i) => i.command);
    // Ungrouped first (empty group name), then extensions@1..@n
    expect(order).toEqual(["cmd.ungrouped", "cmd.one", "cmd.two", "cmd.hidden"]);
  });

  it("inserts separators when the extension menu group changes", () => {
    const resolved: ResolvedMenuItem[] = [
      { extensionId: "ext", command: "cmd.ungrouped1", when: null, group: null, enabled: true },
      // Empty-string groups should be treated as the same group as `null` (both become the empty group name).
      { extensionId: "ext", command: "cmd.ungrouped2", when: null, group: "", enabled: true },
      { extensionId: "ext", command: "cmd.alpha1", when: null, group: "alpha@1", enabled: true },
      { extensionId: "ext", command: "cmd.alpha2", when: null, group: "alpha@2", enabled: true },
      { extensionId: "ext", command: "cmd.beta", when: null, group: "beta@1", enabled: true },
    ];

    const commandRegistry = {
      getCommand: (id: string): CommandContribution =>
        ({
          commandId: id,
          title: id === "cmd.alpha1" ? "Alpha One" : id,
          category: id === "cmd.alpha1" ? "Category" : null,
          icon: null,
          source: { kind: "builtin" as const },
        }),
    };

    const model = buildContextMenuModel(resolved, commandRegistry);

    expect(model.map((entry) => (entry.kind === "separator" ? "|" : entry.commandId))).toEqual([
      "cmd.ungrouped1",
      "cmd.ungrouped2",
      "|",
      "cmd.alpha1",
      "cmd.alpha2",
      "|",
      "cmd.beta",
    ]);

    const alpha = model.find(
      (e): e is Extract<ContextMenuModelItem, { kind: "command" }> => e.kind === "command" && e.commandId === "cmd.alpha1"
    );
    expect(alpha?.label).toBe("Category: Alpha One");
  });
});
