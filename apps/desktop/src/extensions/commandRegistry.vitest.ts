import { describe, expect, it } from "vitest";

import { CommandRegistry } from "./commandRegistry.js";

describe("CommandRegistry", () => {
  it("registers, lists, and executes builtin commands", async () => {
    const registry = new CommandRegistry();

    registry.registerBuiltinCommand(
      "view.test",
      "Test Command",
      (value: number) => value + 1,
      { category: "View", icon: "test-icon", description: "Adds one", keywords: [" add ", "plus", "   "] },
    );

    const listed = registry.listCommands();
    expect(listed).toHaveLength(1);
    expect(listed[0]).toMatchObject({
      commandId: "view.test",
      title: "Test Command",
      category: "View",
      icon: "test-icon",
      description: "Adds one",
      keywords: ["add", "plus"],
      source: { kind: "builtin" },
    });

    expect(registry.getCommand("view.test")).toMatchObject({ commandId: "view.test", title: "Test Command" });
    await expect(registry.executeCommand("view.test", 41)).resolves.toBe(42);
  });

  it("setExtensionCommands replaces only extension commands", async () => {
    const registry = new CommandRegistry();
    registry.registerBuiltinCommand("builtin.one", "Builtin One", () => "builtin", { category: "View" });

    registry.setExtensionCommands(
      [{ extensionId: "ext1", command: "ext.one", title: "Ext One" }],
      async (commandId) => `ran:${commandId}`,
    );

    expect(registry.listCommands().map((c) => c.commandId).sort()).toEqual(["builtin.one", "ext.one"].sort());
    await expect(registry.executeCommand("ext.one")).resolves.toBe("ran:ext.one");

    registry.setExtensionCommands(
      [{ extensionId: "ext1", command: "ext.two", title: "Ext Two" }],
      async (commandId) => `ran:${commandId}`,
    );

    // Builtins are preserved; previous extension commands are removed.
    expect(registry.listCommands().map((c) => c.commandId).sort()).toEqual(["builtin.one", "ext.two"].sort());
    expect(registry.getCommand("ext.one")).toBeUndefined();
  });

  it("trims extension command keywords before storing them", () => {
    const registry = new CommandRegistry();
    registry.setExtensionCommands(
      [
        {
          extensionId: "ext1",
          command: "ext.one",
          title: "Ext One",
          keywords: ["  foo  ", "", "bar"],
        },
      ],
      async () => null,
    );

    expect(registry.getCommand("ext.one")?.keywords).toEqual(["foo", "bar"]);
  });

  it("handles duplicate command ids deterministically", () => {
    const registry = new CommandRegistry();
    registry.registerBuiltinCommand("dup", "Dup", () => {});

    expect(() =>
      registry.setExtensionCommands([{ extensionId: "ext1", command: "dup", title: "Dup from ext" }], async () => null),
    ).toThrow(/Duplicate command id/);

    expect(registry.listCommands()).toHaveLength(1);
    expect(registry.getCommand("dup")?.source.kind).toBe("builtin");
  });

  it("fires execution events for both builtin and extension commands", async () => {
    const registry = new CommandRegistry();

    const events: Array<{ commandId: string; args: any[]; result?: any; error?: unknown }> = [];
    const dispose = registry.onDidExecuteCommand((evt) => {
      events.push(evt);
    });

    registry.registerBuiltinCommand("builtin.exec", "Builtin Exec", (a: number, b: number) => a + b);
    await expect(registry.executeCommand("builtin.exec", 1, 2)).resolves.toBe(3);

    registry.setExtensionCommands(
      [{ extensionId: "ext1", command: "ext.exec", title: "Ext Exec" }],
      async (_commandId, value: string) => `ok:${value}`,
    );
    await expect(registry.executeCommand("ext.exec", "hi")).resolves.toBe("ok:hi");

    dispose();
    await registry.executeCommand("builtin.exec", 2, 2);

    expect(events.map((e) => e.commandId)).toEqual(["builtin.exec", "ext.exec"]);
    expect(events[0]).toMatchObject({ commandId: "builtin.exec", args: [1, 2], result: 3 });
    expect(events[1]).toMatchObject({ commandId: "ext.exec", args: ["hi"], result: "ok:hi" });
  });

  it("allows pre-execution hooks to block command execution", async () => {
    const registry = new CommandRegistry();
    const didRun = { value: false };

    registry.registerBuiltinCommand("blocked.cmd", "Blocked", () => {
      didRun.value = true;
      return "ok";
    });

    const events: Array<{ commandId: string; args: any[]; result?: any; error?: unknown }> = [];
    const disposeDidExecute = registry.onDidExecuteCommand((evt) => events.push(evt));

    const disposeWillExecute = registry.onWillExecuteCommand(() => {
      throw new Error("blocked");
    });

    await expect(registry.executeCommand("blocked.cmd")).rejects.toThrow(/blocked/);
    expect(didRun.value).toBe(false);
    expect(events).toHaveLength(1);
    expect(events[0]?.commandId).toBe("blocked.cmd");
    expect("error" in events[0]!).toBe(true);

    disposeWillExecute();
    disposeDidExecute();
  });
});
