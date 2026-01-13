import { describe, expect, it } from "vitest";

import {
  buildCommandKeybindingAriaIndex,
  buildCommandKeybindingDisplayIndex,
  formatKeybindingForAria,
  formatKeybindingForDisplay,
  getPrimaryCommandKeybindingAria,
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

    expect(parseKeybinding("cmd", "ctl+shift+y")).toMatchObject({
      ctrl: true,
      alt: false,
      shift: true,
      meta: false,
      key: "y",
    });

    expect(parseKeybinding("cmd", "opt+shift+k")).toMatchObject({
      ctrl: false,
      alt: true,
      shift: true,
      meta: false,
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

  it("parseKeybinding rejects unsupported chord syntax", () => {
    expect(parseKeybinding("cmd", "ctrl+k ctrl+c")).toBeNull();
  });

  it("parseKeybinding accepts literal space bindings", () => {
    const binding = parseKeybinding("cmd", " ");
    expect(binding).not.toBeNull();
    expect(binding).toMatchObject({ key: "space", ctrl: false, alt: false, shift: false, meta: false });

    const ctrlSpace = parseKeybinding("cmd", "ctrl+ ");
    expect(ctrlSpace).not.toBeNull();
    expect(ctrlSpace).toMatchObject({ key: "space", ctrl: true });

    // Tabs/newlines are not treated as the spacebar.
    expect(parseKeybinding("cmd", "\t")).toBeNull();
    expect(parseKeybinding("cmd", "ctrl+\t")).toBeNull();
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

  it.each([
    { symbol: "$", code: "Digit4", eventKey: "4" },
    { symbol: "%", code: "Digit5", eventKey: "5" },
    { symbol: "#", code: "Digit3", eventKey: "3" },
  ])("matches Excel-style number format shortcut ctrl+shift+$symbol via KeyboardEvent.code fallback", ({ symbol, code, eventKey }) => {
    const binding = parseKeybinding("cmd", `ctrl+shift+${symbol}`);
    expect(binding).not.toBeNull();
    expect(matchesKeybinding(binding!, eventForKey(eventKey, { ctrlKey: true, shiftKey: true, code }))).toBe(true);
  });

  it.each([
    { symbol: "$", code: "Digit4", eventKey: "4" },
    { symbol: "%", code: "Digit5", eventKey: "5" },
    { symbol: "#", code: "Digit3", eventKey: "3" },
  ])("does not match unshifted ctrl+$symbol via Digit code fallback (avoid ctrl+digit ambiguity)", ({ symbol, code, eventKey }) => {
    const binding = parseKeybinding("cmd", `ctrl+${symbol}`);
    expect(binding).not.toBeNull();
    expect(matchesKeybinding(binding!, eventForKey(eventKey, { ctrlKey: true, shiftKey: false, code }))).toBe(false);
  });

  it("matches Ctrl+Shift+8 when the key is '*' (Excel Ctrl+Shift+*)", () => {
    const binding = parseKeybinding("cmd", "ctrl+shift+8");
    expect(binding).not.toBeNull();
    // On US layouts, Shift+8 produces "*" while the physical key remains Digit8.
    expect(matchesKeybinding(binding!, eventForKey("*", { ctrlKey: true, shiftKey: true, code: "Digit8" }))).toBe(true);
  });

  it("matches Ctrl+Shift+* when the key is '*'", () => {
    const binding = parseKeybinding("cmd", "ctrl+shift+*");
    expect(binding).not.toBeNull();
    expect(matchesKeybinding(binding!, eventForKey("*", { ctrlKey: true, shiftKey: true, code: "Digit8" }))).toBe(true);
  });

  it("matches Ctrl+Shift+* via KeyboardEvent.code fallback when event.key differs", () => {
    const binding = parseKeybinding("cmd", "ctrl+shift+*");
    expect(binding).not.toBeNull();
    // Some environments report `event.key === "8"` even with Shift; ensure we still match
    // the physical key via `code === Digit8`.
    expect(matchesKeybinding(binding!, eventForKey("8", { ctrlKey: true, shiftKey: true, code: "Digit8" }))).toBe(true);
  });

  it("does not match ctrl+* when the physical key is Digit8 without Shift (avoid ctrl+8 ambiguity)", () => {
    const binding = parseKeybinding("cmd", "ctrl+*");
    expect(binding).not.toBeNull();
    expect(matchesKeybinding(binding!, eventForKey("8", { ctrlKey: true, shiftKey: false, code: "Digit8" }))).toBe(false);
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

  it("formatKeybindingForAria renders aria-keyshortcuts strings", () => {
    const binding = parseKeybinding("cmd.test", "ctrl+option+shift+cmd+arrowup")!;
    expect(formatKeybindingForAria(binding)).toBe("Control+Alt+Shift+Meta+ArrowUp");

    const escape = parseKeybinding("cmd.test", "cmd+escape")!;
    expect(formatKeybindingForAria(escape)).toBe("Meta+Escape");

    const pageUp = parseKeybinding("cmd.test", "ctrl+pgup")!;
    expect(formatKeybindingForAria(pageUp)).toBe("Control+PageUp");

    const f2 = parseKeybinding("cmd.test", "shift+f2")!;
    expect(formatKeybindingForAria(f2)).toBe("Shift+F2");

    const copy = parseKeybinding("cmd.test", "ctrl+c")!;
    expect(formatKeybindingForAria(copy)).toBe("Control+C");
  });

  it("formatKeybindingForDisplay preserves punctuation tokens", () => {
    const semicolon = parseKeybinding("cmd.test", "ctrl+shift+;")!;
    expect(formatKeybindingForDisplay(semicolon, "other")).toBe("Ctrl+Shift+;");
    expect(formatKeybindingForDisplay(semicolon, "mac")).toBe("⌃⇧;");

    const equal = parseKeybinding("cmd.test", "ctrl+shift+=")!;
    expect(formatKeybindingForDisplay(equal, "other")).toBe("Ctrl+Shift+=");
    expect(formatKeybindingForDisplay(equal, "mac")).toBe("⌃⇧=");

    const currency = parseKeybinding("cmd.test", "ctrl+shift+$")!;
    expect(formatKeybindingForDisplay(currency, "other")).toBe("Ctrl+Shift+$");
    expect(formatKeybindingForDisplay(currency, "mac")).toBe("⌃⇧$");

    const percent = parseKeybinding("cmd.test", "ctrl+shift+%")!;
    expect(formatKeybindingForDisplay(percent, "other")).toBe("Ctrl+Shift+%");
    expect(formatKeybindingForDisplay(percent, "mac")).toBe("⌃⇧%");

    const date = parseKeybinding("cmd.test", "ctrl+shift+#")!;
    expect(formatKeybindingForDisplay(date, "other")).toBe("Ctrl+Shift+#");
    expect(formatKeybindingForDisplay(date, "mac")).toBe("⌃⇧#");
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

  it("buildCommandKeybindingAriaIndex returns primary binding for a command", () => {
    const index = buildCommandKeybindingAriaIndex({
      platform: "other",
      builtin: [{ command: "cmd.one", key: "ctrl+b" }],
      contributed: [
        { command: "cmd.one", key: "ctrl+k" },
        { command: "cmd.one", key: "ctrl+k" }, // duplicate should not add a second time
      ],
    });

    expect(index.get("cmd.one")).toEqual(["Control+B", "Control+K"]);
    expect(getPrimaryCommandKeybindingAria("cmd.one", index)).toBe("Control+B");
  });
});
