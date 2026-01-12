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
  const raw = String(binding ?? "").trim();
  if (!raw) return null;

  const parts = raw
    .split("+")
    .map((p) => normalizeKeyToken(p))
    .filter((p) => p.length > 0);

  if (parts.length === 0) return null;

  const out: ParsedKeybinding = {
    command,
    when: when ?? null,
    ctrl: false,
    alt: false,
    shift: false,
    meta: false,
    key: parts[parts.length - 1]!,
  };

  for (const part of parts.slice(0, -1)) {
    if (part === "ctrl" || part === "control") out.ctrl = true;
    else if (part === "shift") out.shift = true;
    else if (part === "alt" || part === "option") out.alt = true;
    else if (part === "cmd" || part === "command" || part === "meta") out.meta = true;
  }

  return out;
}

function normalizeEventKey(event: KeyboardEvent): string {
  return normalizeKeyToken(event.key);
}

export function matchesKeybinding(binding: ParsedKeybinding, event: KeyboardEvent): boolean {
  if (binding.ctrl !== event.ctrlKey) return false;
  if (binding.shift !== event.shiftKey) return false;
  if (binding.alt !== event.altKey) return false;
  if (binding.meta !== event.metaKey) return false;

  return normalizeEventKey(event) === binding.key;
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
