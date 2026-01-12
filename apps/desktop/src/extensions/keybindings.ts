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
