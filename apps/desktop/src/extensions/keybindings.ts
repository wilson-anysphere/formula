export type ContributedKeybinding = {
  extensionId: string;
  command: string;
  key: string;
  mac: string | null;
  when: string | null;
};

export type ParsedKeybinding = {
  command: string;
  when: string | null;
  ctrl: boolean;
  alt: boolean;
  shift: boolean;
  meta: boolean;
  key: string;
};

const KEY_ALIASES: Record<string, string> = {
  esc: "escape",
  del: "delete",
  return: "enter",
  spacebar: "space",
  pgup: "pageup",
  pgdn: "pagedown",
  pgdown: "pagedown",
  up: "arrowup",
  down: "arrowdown",
  left: "arrowleft",
  right: "arrowright",
};

function normalizeKeyToken(token: string): string {
  // `KeyboardEvent.key` uses a literal space for the spacebar in modern browsers.
  // Keep our internal string key representation consistent with extension manifests.
  if (token === " ") return "space";
  const normalized = token.trim().toLowerCase();
  return KEY_ALIASES[normalized] ?? normalized;
}

export function parseKeybinding(command: string, binding: string, when: string | null = null): ParsedKeybinding | null {
  const raw = String(binding ?? "");
  if (!raw) return null;

  // Treat a binding that is literally a single space token as the spacebar.
  // (This matches `KeyboardEvent.key === " "` and is occasionally used in manifests.)
  //
  // Important: do not treat other whitespace-only strings (e.g. tabs/newlines) as Space.
  if (raw.trim() === "") {
    if (!/^[ ]+$/.test(raw)) return null;
    return {
      command,
      when: when ?? null,
      ctrl: false,
      alt: false,
      shift: false,
      meta: false,
      key: "space",
    };
  }

  const parts: string[] = [];

  for (const rawPart of raw.split("+")) {
    const trimmed = rawPart.trim();

    // Allow the last token to be a literal space (e.g. `"ctrl+ "` in user-provided strings).
    // Avoid treating missing tokens (e.g. `"ctrl++c"`) as space, and avoid treating other
    // whitespace (tabs/newlines) as space.
    if (trimmed === "") {
      if (/^[ ]+$/.test(rawPart)) {
        parts.push("space");
        continue;
      }
      return null;
    }

    // Defensive: reject unsupported "chord" syntax like `ctrl+k ctrl+c`. We allow
    // whitespace around tokens (trimmed above), but not inside a token.
    if (/\s/.test(trimmed)) return null;

    parts.push(normalizeKeyToken(trimmed));
  }

  if (parts.length === 0) return null;

  const key = parts[parts.length - 1]!;
  const isModifierToken = (token: string): boolean =>
    token === "ctrl" ||
    token === "control" ||
    token === "ctl" ||
    token === "shift" ||
    token === "alt" ||
    token === "option" ||
    token === "opt" ||
    token === "cmd" ||
    token === "command" ||
    token === "meta" ||
    token === "win" ||
    token === "super";

  // A keybinding must specify a non-modifier key.
  if (isModifierToken(key)) return null;

  const out: ParsedKeybinding = {
    command,
    when: when ?? null,
    ctrl: false,
    alt: false,
    shift: false,
    meta: false,
    key,
  };

  for (const part of parts.slice(0, -1)) {
    if (part === "ctrl" || part === "control" || part === "ctl") out.ctrl = true;
    else if (part === "shift") out.shift = true;
    else if (part === "alt" || part === "option" || part === "opt") out.alt = true;
    else if (part === "cmd" || part === "command" || part === "meta" || part === "win" || part === "super") out.meta = true;
    else return null;
  }

  return out;
}

function normalizeEventKey(event: KeyboardEvent): string {
  return normalizeKeyToken(event.key);
}

const keyTokenToCodeFallback: Record<string, KeyboardEvent["code"] | KeyboardEvent["code"][]> = {
  ";": "Semicolon",
  "=": "Equal",
  "-": "Minus",
  "`": "Backquote",
  "[": "BracketLeft",
  "]": "BracketRight",
  // Dedicated "*" keys can appear on the numpad (`KeyboardEvent.code === "NumpadMultiply"`).
  //
  // Note: some Excel-style shortcuts refer to `Ctrl+Shift+*` but are actually bound to the
  // physical `Digit8` key on many layouts. We handle that conditionally in `matchesKeybinding`
  // when Shift is required, so `ctrl+*` does *not* accidentally match `Ctrl+8`.
  "*": "NumpadMultiply",
  // ISO/JIS keyboards may report alternate codes for the physical backslash key.
  "\\": ["Backslash", "IntlBackslash", "IntlYen", "IntlRo"],
  "/": "Slash",
  ",": "Comma",
  ".": "Period",
  "'": "Quote",

  // Digits often produce different characters under Shift (e.g. "1" -> "!"),
  // but the physical key stays stable via `KeyboardEvent.code`.
  //
  // Excel number-format presets use `Ctrl/Cmd+Shift+$`, `Ctrl/Cmd+Shift+%`, and
  // `Ctrl/Cmd+Shift+#`. On many layouts the literal `$`/`%`/`#` character may not be
  // reported in `event.key` for the chord, but the physical Digit keys are stable.
  "$": "Digit4",
  "%": "Digit5",
  "#": "Digit3",
  "0": "Digit0",
  "1": "Digit1",
  "2": "Digit2",
  "3": "Digit3",
  "4": "Digit4",
  "5": "Digit5",
  "6": "Digit6",
  "7": "Digit7",
  "8": "Digit8",
  "9": "Digit9",
};

export function matchesKeybinding(binding: ParsedKeybinding, event: KeyboardEvent): boolean {
  if (binding.ctrl !== event.ctrlKey) return false;
  if (binding.shift !== event.shiftKey) return false;
  if (binding.alt !== event.altKey) return false;
  if (binding.meta !== event.metaKey) return false;

  if (normalizeEventKey(event) === binding.key) return true;

  // Special-case: Excel-style `Ctrl+Shift+*` is frequently entered as `Ctrl+Shift+8`
  // on many layouts. When Shift is required, allow falling back to the physical Digit8 key.
  if (binding.key === "*" && binding.shift) {
    if (event.code === "Digit8") return true;
  }

  // Excel-style number-format shortcuts are spelled with `$`, `%`, and `#` but correspond to the
  // shifted Digit4/Digit5/Digit3 keys on many layouts. Only apply the `code` fallback when Shift
  // is explicitly part of the binding to avoid `ctrl+$` accidentally matching `Ctrl+4`, etc.
  if (!binding.shift && (binding.key === "$" || binding.key === "%" || binding.key === "#")) return false;

  const fallback = keyTokenToCodeFallback[binding.key];
  if (!fallback) return false;

  if (Array.isArray(fallback)) return fallback.includes(event.code);
  return event.code === fallback;
}

export function platformKeybinding(binding: ContributedKeybinding, platform: "mac" | "other"): string {
  if (platform === "mac" && binding.mac) return binding.mac;
  return binding.key;
}

function formatKeyTokenForDisplay(keyToken: string, platform: "mac" | "other"): string {
  const key = normalizeKeyToken(keyToken);

  switch (key) {
    case "arrowup":
      return platform === "mac" ? "↑" : "Up";
    case "arrowdown":
      return platform === "mac" ? "↓" : "Down";
    case "arrowleft":
      return platform === "mac" ? "←" : "Left";
    case "arrowright":
      return platform === "mac" ? "→" : "Right";
    case "escape":
      return platform === "mac" ? "⎋" : "Esc";
    case "enter":
    case "return":
      return platform === "mac" ? "↩" : "Enter";
    case "backspace":
      return platform === "mac" ? "⌫" : "Backspace";
    case "delete":
      // macOS has both delete (backspace) and forward delete; VS Code uses ⌦
      // for the forward delete key.
      return platform === "mac" ? "⌦" : "Delete";
    case "space":
      return "Space";
    case "tab":
      return platform === "mac" ? "⇥" : "Tab";
    case "pageup":
      return platform === "mac" ? "⇞" : "PgUp";
    case "pagedown":
      return platform === "mac" ? "⇟" : "PgDn";
    default:
      break;
  }

  // Preserve punctuation (/, ., etc) as-is.
  if (key.length === 1) return key.toUpperCase();

  // F-keys are common and should be upper-cased (f1 -> F1).
  if (/^f\d{1,2}$/.test(key)) return key.toUpperCase();

  // Default: Title-case only the first character (pageup -> Pageup). Prefer a
  // predictable output over overfitting for every possible `event.key`.
  return key.slice(0, 1).toUpperCase() + key.slice(1);
}

/**
 * Renders a parsed keybinding as a human-readable label for UI.
 *
 * Note: This is display-only. Matching semantics are intentionally unchanged.
 */
export function formatKeybindingForDisplay(binding: ParsedKeybinding, platform: "mac" | "other"): string {
  const key = formatKeyTokenForDisplay(binding.key, platform);

  if (platform === "mac") {
    // Match macOS menu ordering: Control, Option, Shift, Command.
    return `${binding.ctrl ? "⌃" : ""}${binding.alt ? "⌥" : ""}${binding.shift ? "⇧" : ""}${binding.meta ? "⌘" : ""}${key}`;
  }

  const parts: string[] = [];
  // Match the requirement ordering: Ctrl+Alt+Shift+Meta+<Key>.
  if (binding.ctrl) parts.push("Ctrl");
  if (binding.alt) parts.push("Alt");
  if (binding.shift) parts.push("Shift");
  if (binding.meta) parts.push("Meta");
  parts.push(key);
  return parts.join("+");
}

export type KeybindingContribution = {
  command: string;
  key: string;
  mac?: string | null;
  when?: string | null;
};

/**
 * Builds a `commandId -> [displayKeybinding...]` index (primary binding first).
 *
 * Intended for command palette and context menu rendering.
 */
export function buildCommandKeybindingDisplayIndex(params: {
  platform: "mac" | "other";
  contributed: KeybindingContribution[];
  builtin?: KeybindingContribution[];
}): Map<string, string[]> {
  const index = new Map<string, string[]>();

  const add = (kb: KeybindingContribution) => {
    const binding = params.platform === "mac" && kb.mac ? kb.mac : kb.key;
    const parsed = parseKeybinding(kb.command, binding, kb.when ?? null);
    if (!parsed) return;
    const display = formatKeybindingForDisplay(parsed, params.platform);

    const existing = index.get(parsed.command);
    if (!existing) {
      index.set(parsed.command, [display]);
      return;
    }
    if (!existing.includes(display)) existing.push(display);
  };

  for (const kb of params.builtin ?? []) add(kb);
  for (const kb of params.contributed) add(kb);

  return index;
}

export function getPrimaryCommandKeybindingDisplay(commandId: string, index: Map<string, string[]>): string | null {
  return index.get(String(commandId))?.[0] ?? null;
}
