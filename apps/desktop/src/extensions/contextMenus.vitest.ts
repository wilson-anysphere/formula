import { describe, expect, it } from "vitest";

import { resolveMenuItems, type ContributedMenuItem } from "./contextMenus.js";

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
});

