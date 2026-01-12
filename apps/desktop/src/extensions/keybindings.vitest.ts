import { describe, expect, it } from "vitest";

import { matchesKeybinding, parseKeybinding } from "./keybindings.js";

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

describe("extension keybindings", () => {
  it("normalizes common key name aliases (esc)", () => {
    const binding = parseKeybinding("cmd", "ctrl+esc");
    expect(binding).not.toBeNull();
    expect(matchesKeybinding(binding!, eventForKey("Escape", { ctrlKey: true }))).toBe(true);
  });

  it("normalizes common key name aliases (del)", () => {
    const binding = parseKeybinding("cmd", "ctrl+del");
    expect(binding).not.toBeNull();
    expect(matchesKeybinding(binding!, eventForKey("Delete", { ctrlKey: true }))).toBe(true);
  });

  it("normalizes arrow key aliases (up/down/left/right)", () => {
    const up = parseKeybinding("cmd", "ctrl+up");
    const left = parseKeybinding("cmd", "ctrl+left");
    expect(up).not.toBeNull();
    expect(left).not.toBeNull();

    expect(matchesKeybinding(up!, eventForKey("ArrowUp", { ctrlKey: true }))).toBe(true);
    expect(matchesKeybinding(left!, eventForKey("ArrowLeft", { ctrlKey: true }))).toBe(true);
  });
});

