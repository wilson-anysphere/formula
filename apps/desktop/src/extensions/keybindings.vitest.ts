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
  opts: Partial<Pick<KeyboardEvent, "ctrlKey" | "shiftKey" | "altKey" | "metaKey" | "code">> = {},
): KeyboardEvent {
  return {
    key,
    code: "",
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

    expect(parseKeybinding("cmd", "win+shift+k")).toMatchObject({
      ctrl: false,
      alt: false,
      shift: true,
      meta: true,
      key: "k",
    });

    expect(parseKeybinding("cmd", "super+shift+k")).toMatchObject({
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

  it("normalizes page navigation aliases (pgup/pgdn)", () => {
    const pgup = parseKeybinding("cmd", "ctrl+pgup");
    expect(pgup).not.toBeNull();
    expect(matchesKeybinding(pgup!, eventForKey("PageUp", { ctrlKey: true }))).toBe(true);

    const pgdn = parseKeybinding("cmd", "ctrl+pgdn");
    expect(pgdn).not.toBeNull();
    expect(matchesKeybinding(pgdn!, eventForKey("PageDown", { ctrlKey: true }))).toBe(true);
  });

  it("matches shifted punctuation via KeyboardEvent.code fallback (ctrl+shift+;)", () => {
    const binding = parseKeybinding("cmd", "ctrl+shift+;");
    expect(binding).not.toBeNull();
    expect(matchesKeybinding(binding!, eventForKey(":", { ctrlKey: true, shiftKey: true, code: "Semicolon" }))).toBe(true);
  });

  it("matches shifted punctuation via KeyboardEvent.code fallback (ctrl+shift+=)", () => {
    const binding = parseKeybinding("cmd", "ctrl+shift+=");
    expect(binding).not.toBeNull();
    expect(matchesKeybinding(binding!, eventForKey("+", { ctrlKey: true, shiftKey: true, code: "Equal" }))).toBe(true);
  });

  it("matches shifted digits via KeyboardEvent.code fallback (ctrl+shift+1)", () => {
    const binding = parseKeybinding("cmd", "ctrl+shift+1");
    expect(binding).not.toBeNull();
    expect(matchesKeybinding(binding!, eventForKey("!", { ctrlKey: true, shiftKey: true, code: "Digit1" }))).toBe(true);
  });

  it("matches digit keys via KeyboardEvent.code fallback when event.key differs by layout (ctrl+1)", () => {
    const binding = parseKeybinding("cmd", "ctrl+1");
    expect(binding).not.toBeNull();

    // Example: on AZERTY layouts, the physical `Digit1` key can report `event.key === "&"` without Shift.
    expect(matchesKeybinding(binding!, eventForKey("&", { ctrlKey: true, code: "Digit1" }))).toBe(true);
  });

  it("matches shifted brackets via KeyboardEvent.code fallback (ctrl+shift+[)", () => {
    const binding = parseKeybinding("cmd", "ctrl+shift+[");
    expect(binding).not.toBeNull();
    expect(matchesKeybinding(binding!, eventForKey("{", { ctrlKey: true, shiftKey: true, code: "BracketLeft" }))).toBe(true);
  });

  it("matches shifted backslash via KeyboardEvent.code fallback for IntlBackslash keyboards", () => {
    const binding = parseKeybinding("cmd", "ctrl+shift+\\");
    expect(binding).not.toBeNull();
    expect(matchesKeybinding(binding!, eventForKey("|", { ctrlKey: true, shiftKey: true, code: "IntlBackslash" }))).toBe(true);
  });

  it("does not fall back to KeyboardEvent.code for letters (layout-aware matching)", () => {
    const binding = parseKeybinding("cmd", "ctrl+y");
    expect(binding).not.toBeNull();

    // Example: On QWERTZ layouts, the physical `KeyY` produces `event.key === "z"`.
    // We intentionally keep letter matching semantics based on `event.key`.
    expect(matchesKeybinding(binding!, eventForKey("z", { ctrlKey: true, code: "KeyY" }))).toBe(false);
  });

  it("formatKeybindingForDisplay renders mac vs other", () => {
    const binding = parseKeybinding("cmd.test", "ctrl+option+shift+cmd+arrowup")!;
    expect(formatKeybindingForDisplay(binding, "mac")).toMatchInlineSnapshot('"⌃⌥⇧⌘↑"');
    expect(formatKeybindingForDisplay(binding, "other")).toMatchInlineSnapshot('"Ctrl+Alt+Shift+Meta+Up"');

    const escape = parseKeybinding("cmd.test", "cmd+escape")!;
    expect(formatKeybindingForDisplay(escape, "mac")).toMatchInlineSnapshot('"⌘⎋"');
    expect(formatKeybindingForDisplay(escape, "other")).toMatchInlineSnapshot('"Meta+Esc"');
  });

  it("formatKeybindingForDisplay preserves punctuation tokens", () => {
    const semicolon = parseKeybinding("cmd.test", "ctrl+shift+;")!;
    expect(formatKeybindingForDisplay(semicolon, "other")).toBe("Ctrl+Shift+;");
    expect(formatKeybindingForDisplay(semicolon, "mac")).toBe("⌃⇧;");

    const equal = parseKeybinding("cmd.test", "ctrl+shift+=")!;
    expect(formatKeybindingForDisplay(equal, "other")).toBe("Ctrl+Shift+=");
    expect(formatKeybindingForDisplay(equal, "mac")).toBe("⌃⇧=");
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
