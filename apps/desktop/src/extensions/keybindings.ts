export type ContributedKeybinding = {
  extensionId: string;
  command: string;
  key: string;
  mac: string | null;
};

export type ParsedKeybinding = {
  command: string;
  ctrl: boolean;
  alt: boolean;
  shift: boolean;
  meta: boolean;
  key: string;
};

function normalizeKeyToken(token: string): string {
  return token.trim().toLowerCase();
}

export function parseKeybinding(command: string, binding: string): ParsedKeybinding | null {
  const raw = String(binding ?? "").trim();
  if (!raw) return null;

  const parts = raw
    .split("+")
    .map((p) => normalizeKeyToken(p))
    .filter((p) => p.length > 0);

  if (parts.length === 0) return null;

  const out: ParsedKeybinding = {
    command,
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
  if (event.key === " ") return "space";
  const key = event.key.toLowerCase();
  if (key === "esc") return "escape";
  return key;
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

