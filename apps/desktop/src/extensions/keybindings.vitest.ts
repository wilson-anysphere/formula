import { describe, expect, it } from "vitest";

import {
  buildCommandKeybindingDisplayIndex,
  formatKeybindingForDisplay,
  getPrimaryCommandKeybindingDisplay,
  matchesKeybinding,
  parseKeybinding,
} from "./keybindings.js";

function eventForKey(
  key: string,
  opts: Partial<Pick<KeyboardEvent, "ctrlKey" | "shiftKey" | "altKey" | "metaKey">> = {},
): KeyboardEvent {
  return {
    key,
    ctrlKey: false,
    shiftKey: false,
    altKey: false,
    metaKey: false,
    ...opts,
  } as KeyboardEvent;
}

describe("keybindings", () => {
  it("parseKeybinding handles modifier synonyms", () => {
    expect(parseKeybinding("cmd", "ctrl+shift+y")).toMatchObject({
      ctrl: true,
      alt: false,
      shift: true,
      meta: false,
      key: "y",
    });

    expect(parseKeybinding("cmd", "control+option+cmd+p")).toMatchObject({
      ctrl: true,
      alt: true,
      shift: false,
      meta: true,
      key: "p",
    });

    expect(parseKeybinding("cmd", "meta+shift+k")).toMatchObject({
      ctrl: false,
      alt: false,
      shift: true,
      meta: true,
      key: "k",
    });
  });

  it("matchesKeybinding handles special keys (space/escape)", () => {
    const spaceBinding = parseKeybinding("cmd.space", "space");
    expect(spaceBinding).not.toBeNull();
    expect(matchesKeybinding(spaceBinding!, eventForKey(" "))).toBe(true);

    const escapeBinding = parseKeybinding("cmd.escape", "escape");
    expect(escapeBinding).not.toBeNull();
    expect(matchesKeybinding(escapeBinding!, eventForKey("Esc"))).toBe(true);
  });

  it("normalizes common key name aliases (esc/del/arrows)", () => {
    const esc = parseKeybinding("cmd", "ctrl+esc");
    expect(esc).not.toBeNull();
    expect(matchesKeybinding(esc!, eventForKey("Escape", { ctrlKey: true }))).toBe(true);

    const del = parseKeybinding("cmd", "ctrl+del");
    expect(del).not.toBeNull();
    expect(matchesKeybinding(del!, eventForKey("Delete", { ctrlKey: true }))).toBe(true);

    const up = parseKeybinding("cmd", "ctrl+up");
    const left = parseKeybinding("cmd", "ctrl+left");
    expect(up).not.toBeNull();
    expect(left).not.toBeNull();
    expect(matchesKeybinding(up!, eventForKey("ArrowUp", { ctrlKey: true }))).toBe(true);
    expect(matchesKeybinding(left!, eventForKey("ArrowLeft", { ctrlKey: true }))).toBe(true);
  });

  it("formatKeybindingForDisplay renders mac vs other", () => {
    const binding = parseKeybinding("cmd.test", "ctrl+option+shift+cmd+arrowup")!;
    expect(formatKeybindingForDisplay(binding, "mac")).toMatchInlineSnapshot('"⌃⌥⇧⌘↑"');
    expect(formatKeybindingForDisplay(binding, "other")).toMatchInlineSnapshot('"Ctrl+Alt+Shift+Meta+Up"');

    const escape = parseKeybinding("cmd.test", "cmd+escape")!;
    expect(formatKeybindingForDisplay(escape, "mac")).toMatchInlineSnapshot('"⌘⎋"');
    expect(formatKeybindingForDisplay(escape, "other")).toMatchInlineSnapshot('"Meta+Esc"');
  });

  it("buildCommandKeybindingDisplayIndex returns primary binding for a command", () => {
    const index = buildCommandKeybindingDisplayIndex({
      platform: "other",
      builtin: [{ command: "cmd.one", key: "ctrl+b" }],
      contributed: [
        { command: "cmd.one", key: "ctrl+k" },
        { command: "cmd.one", key: "ctrl+k" }, // duplicate should not add a second time
      ],
    });

    expect(index.get("cmd.one")).toEqual(["Ctrl+B", "Ctrl+K"]);
    expect(getPrimaryCommandKeybindingDisplay("cmd.one", index)).toBe("Ctrl+B");
  });
});
