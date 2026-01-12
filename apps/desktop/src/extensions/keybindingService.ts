import type { CommandRegistry } from "./commandRegistry.js";
import type { ContextKeyService } from "./contextKeys.js";
import { matchesKeybinding, parseKeybinding, type ContributedKeybinding, type ParsedKeybinding } from "./keybindings.js";
import { evaluateWhenClause } from "./whenClause.js";

export type BuiltinKeybinding = {
  command: string;
  key: string;
  mac?: string | null;
  when?: string | null;
  /**
   * Optional tiebreaker within the same priority group. Higher wins.
   * If omitted, all bindings have the same weight.
   */
  weight?: number;
};

export type KeybindingSource = { kind: "builtin" } | { kind: "extension"; extensionId: string };

type StoredKeybinding = {
  source: KeybindingSource;
  binding: ParsedKeybinding;
  weight: number;
  order: number;
};

function detectPlatform(): "mac" | "other" {
  const platform = typeof navigator !== "undefined" ? navigator.platform : "";
  return /Mac|iPhone|iPad|iPod/.test(platform) ? "mac" : "other";
}

function pickPlatformKeybinding(binding: { key: string; mac?: string | null }, platform: "mac" | "other"): string {
  if (platform === "mac" && binding.mac) return binding.mac;
  return binding.key;
}

function shouldIgnoreTarget(target: EventTarget | null): boolean {
  const el = target as HTMLElement | null;
  if (!el) return false;
  const tag = el.tagName;
  return tag === "INPUT" || tag === "TEXTAREA" || el.isContentEditable;
}

// Reserved shortcuts that extensions can never claim (safety net).
const RESERVED_EXTENSION_SHORTCUTS: ParsedKeybinding[] = [
  // Copy/Cut/Paste
  parseKeybinding("__reserved__", "ctrl+c")!,
  parseKeybinding("__reserved__", "cmd+c")!,
  parseKeybinding("__reserved__", "ctrl+x")!,
  parseKeybinding("__reserved__", "cmd+x")!,
  parseKeybinding("__reserved__", "ctrl+v")!,
  parseKeybinding("__reserved__", "cmd+v")!,
  // Paste Special
  parseKeybinding("__reserved__", "ctrl+shift+v")!,
  parseKeybinding("__reserved__", "cmd+shift+v")!,
  // Command palette
  parseKeybinding("__reserved__", "ctrl+shift+p")!,
  parseKeybinding("__reserved__", "cmd+shift+p")!,
];

function isReservedForExtensions(event: KeyboardEvent): boolean {
  // Fast path: most shortcuts won't be primary modifier combos.
  if (!event.ctrlKey && !event.metaKey) return false;
  if (event.altKey) return false;
  return RESERVED_EXTENSION_SHORTCUTS.some((binding) => matchesKeybinding(binding, event));
}

export class KeybindingService {
  private readonly platform: "mac" | "other";
  private builtin: StoredKeybinding[] = [];
  private extensions: StoredKeybinding[] = [];
  private orderCounter = 0;

  private removeListener: (() => void) | null = null;

  constructor(
    private readonly params: {
      commandRegistry: CommandRegistry;
      contextKeys: ContextKeyService;
      /**
       * Hook to ensure commands are available before executing. Most commonly used to
       * lazy-load/sync extension commands before dispatch.
       */
      onBeforeExecuteCommand?: (commandId: string, source: KeybindingSource) => Promise<void>;
      /**
       * Optional error handler for failed command execution.
       */
      onCommandError?: (commandId: string, err: unknown) => void;
      platform?: "mac" | "other";
    },
  ) {
    this.platform = params.platform ?? detectPlatform();
  }

  setBuiltinKeybindings(bindings: BuiltinKeybinding[]): void {
    const next: StoredKeybinding[] = [];
    for (const kb of bindings) {
      const primary = pickPlatformKeybinding(kb, this.platform);
      const candidates = new Set<string>([primary]);

      // On macOS, allow the base keybinding as a fallback when a mac-specific variant exists.
      // Example: Replace is `Cmd+Option+F` on macOS to avoid the system `Cmd+H` shortcut, but
      // we still want `Ctrl+H` to work for users accustomed to the Windows/Linux shortcut.
      if (this.platform === "mac" && kb.mac && kb.key && kb.key !== primary) {
        candidates.add(kb.key);
      }

      for (const raw of candidates) {
        const parsed = parseKeybinding(kb.command, raw, kb.when ?? null);
        if (!parsed) continue;
        next.push({
          source: { kind: "builtin" },
          binding: parsed,
          weight: typeof kb.weight === "number" ? kb.weight : 0,
          order: ++this.orderCounter,
        });
      }
    }
    this.builtin = next;
  }

  setExtensionKeybindings(bindings: ContributedKeybinding[]): void {
    const next: StoredKeybinding[] = [];
    for (const kb of bindings) {
      const raw = pickPlatformKeybinding(kb, this.platform);
      const parsed = parseKeybinding(kb.command, raw, kb.when ?? null);
      if (!parsed) continue;
      next.push({
        source: { kind: "extension", extensionId: kb.extensionId },
        binding: parsed,
        weight: 0,
        order: ++this.orderCounter,
      });
    }
    this.extensions = next;
  }

  /**
   * Install the global keydown listener (bubble phase by default).
   */
  installWindowListener(target: Window = window, opts: { capture?: boolean } = {}): () => void {
    this.dispose();
    const capture = opts.capture ?? false;
    const handler = (e: KeyboardEvent) => {
      void this.dispatchKeydown(e);
    };
    target.addEventListener("keydown", handler, { capture });
    this.removeListener = () => target.removeEventListener("keydown", handler, { capture });
    return () => this.dispose();
  }

  dispose(): void {
    this.removeListener?.();
    this.removeListener = null;
  }

  async dispatchKeydown(event: KeyboardEvent): Promise<boolean> {
    if (event.defaultPrevented) return false;
    if (shouldIgnoreTarget(event.target)) return false;

    const match = this.findMatchingBinding(event);
    if (!match) return false;

    event.preventDefault();
    try {
      await this.params.onBeforeExecuteCommand?.(match.binding.command, match.source);
      await this.params.commandRegistry.executeCommand(match.binding.command);
    } catch (err) {
      this.params.onCommandError?.(match.binding.command, err);
    }
    return true;
  }

  private findMatchingBinding(event: KeyboardEvent): StoredKeybinding | null {
    const lookup = this.params.contextKeys.asLookup();

    // Built-ins always win.
    const builtin = this.findFirstMatch(this.builtin, event, lookup);
    if (builtin) return builtin;

    // Safety net: reserved shortcuts should never be claimed by extensions.
    if (isReservedForExtensions(event)) return null;

    return this.findFirstMatch(this.extensions, event, lookup);
  }

  private findFirstMatch(
    bindings: StoredKeybinding[],
    event: KeyboardEvent,
    lookup: ReturnType<ContextKeyService["asLookup"]>,
  ): StoredKeybinding | null {
    // Deterministic ordering: higher weight wins; otherwise, first registered wins.
    // Avoid sorting allocations on every keydown by computing a stable scan order.
    // For now, do a simple linear scan picking the best match.
    let best: StoredKeybinding | null = null;
    for (const entry of bindings) {
      if (!matchesKeybinding(entry.binding, event)) continue;
      if (!evaluateWhenClause(entry.binding.when, lookup)) continue;
      if (!best) {
        best = entry;
        continue;
      }
      if (entry.weight !== best.weight) {
        if (entry.weight > best.weight) best = entry;
        continue;
      }
      if (entry.order < best.order) best = entry;
    }
    return best;
  }
}
