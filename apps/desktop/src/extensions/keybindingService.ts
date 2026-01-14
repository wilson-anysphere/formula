import type { CommandRegistry } from "./commandRegistry.js";
import type { ContextKeyService } from "./contextKeys.js";
import {
  buildCommandKeybindingAriaIndex,
  buildCommandKeybindingDisplayIndex,
  matchesKeybinding,
  parseKeybinding,
  type ContributedKeybinding,
  type KeybindingContribution,
  type ParsedKeybinding,
} from "./keybindings.js";
import { evaluateWhenClause } from "./whenClause.js";

export type BuiltinKeybinding = KeybindingContribution & {
  /**
   * Optional tiebreaker within the same priority group. Higher wins.
   * If omitted, all bindings have the same weight.
   */
  weight?: number;
  /**
   * If true, the binding is allowed to fire on repeated `keydown` events while
   * the user holds the chord down (e.g. Excel-like sheet navigation).
   *
   * Defaults to false to avoid accidental repeats for toggle commands (palette,
   * AI chat, etc).
   */
  allowRepeat?: boolean;
};

export type KeybindingSource = { kind: "builtin" } | { kind: "extension"; extensionId: string };

type StoredKeybinding = {
  source: KeybindingSource;
  binding: ParsedKeybinding;
  weight: number;
  order: number;
  allowRepeat: boolean;
};

function detectPlatform(): "mac" | "other" {
  const platform = typeof navigator !== "undefined" ? navigator.platform : "";
  return /Mac|iPhone|iPad|iPod/.test(platform) ? "mac" : "other";
}

function pickPlatformKeybinding(binding: { key: string; mac?: string | null }, platform: "mac" | "other"): string {
  if (platform === "mac" && binding.mac) return binding.mac;
  return binding.key;
}

function isInputTarget(target: EventTarget | null): boolean {
  const el = target as HTMLElement | null;
  if (!el) return false;
  const tag = el.tagName;
  // Coerce `isContentEditable` because some DOM shims (jsdom) may not define it.
  return tag === "INPUT" || tag === "TEXTAREA" || Boolean((el as any).isContentEditable);
}

function isInsideKeybindingBarrier(target: EventTarget | null): boolean {
  if (!target || typeof target !== "object") return false;

  // Preferred fast path: Element.closest (works for nested overlay roots and is
  // resilient to focus being on any descendant element).
  const closest = (target as any).closest;
  if (typeof closest === "function") {
    try {
      // NOTE: `closest()` does not cross shadow DOM boundaries. If no match is found
      // and the target isn't inside a shadow root, we can safely return `false`
      // without doing a second manual traversal.
      if (closest.call(target, '[data-keybinding-barrier="true"]')) return true;

      const root = (target as any).getRootNode?.();
      const inShadowRoot = Boolean(root && typeof root === "object" && typeof (root as any).host === "object");
      if (!inShadowRoot) return false;
    } catch {
      // ignore and fall back to manual traversal
    }
  }

  // Fallback for non-Element targets (or test doubles) that don't support `closest`.
  let node: any = target;
  while (node && typeof node === "object") {
    if (node.dataset?.keybindingBarrier === "true") return true;
    if (typeof node.getAttribute === "function" && node.getAttribute("data-keybinding-barrier") === "true") return true;
    // Shadow DOM: climb from a ShadowRoot back to its host element.
    node = node.parentElement ?? node.parentNode ?? node.host ?? null;
  }
  return false;
}

export const DEFAULT_RESERVED_EXTENSION_SHORTCUTS = [
  // Core cancellation key (closing dialogs/menus, canceling interactions, etc).
  // Extensions should never be able to claim this.
  "escape",
  // Core activation/navigation key (e.g. committing edits, dialog default actions).
  // Extensions should never be able to claim this.
  "enter",
  // Copy/Cut/Paste (core text handling should not be overrideable by extensions).
  "ctrl+c",
  "cmd+c",
  "ctrl+cmd+c",
  "ctrl+x",
  "cmd+x",
  "ctrl+cmd+x",
  "ctrl+v",
  "cmd+v",
  "ctrl+cmd+v",
  // Paste Special (built-in chord; extensions should not claim it).
  "ctrl+shift+v",
  "cmd+shift+v",
  "ctrl+cmd+shift+v",
  // Command palette (extensions should not claim it).
  "ctrl+shift+p",
  "cmd+shift+p",
  // Some keyboards emit both ctrl+meta on the same chord.
  "ctrl+cmd+shift+p",
  // Quick Open (Tauri global shortcut; extensions should not claim it).
  "ctrl+shift+o",
  "cmd+shift+o",
  // Some keyboards emit both ctrl+meta on the same chord.
  "ctrl+cmd+shift+o",
  // Inline AI edit (core UX shortcut; extensions should not claim it).
  "ctrl+k",
  "cmd+k",
  // Some keyboards emit both ctrl+meta on the same chord.
  "ctrl+cmd+k",
  // Edit cell (Excel-style).
  "f2",
  // Add Comment (Excel-style).
  "shift+f2",
  // Focus cycling between major UI regions (Excel-style).
  "f6",
  "shift+f6",
  // Open context menu (Excel-style / Windows keyboard context-menu key).
  "shift+f10",
  "contextmenu",
  // Workbook sheet navigation (Excel-style).
  "ctrl+pageup",
  "cmd+pageup",
  // Some keyboards emit both ctrl+meta on the same chord.
  "ctrl+cmd+pageup",
  "ctrl+pagedown",
  "cmd+pagedown",
  // Some keyboards emit both ctrl+meta on the same chord.
  "ctrl+cmd+pagedown",
  // AI Chat toggle (core UX shortcut; extensions should not claim it).
  // - Windows/Linux: Ctrl+Shift+A
  // - macOS: Cmd+I (with Ctrl+Cmd+I fallback for some remote/VM keyboard setups)
  "ctrl+shift+a",
  "cmd+i",
  "ctrl+cmd+i",
  // Comments panel toggle (core UX shortcut; extensions should not claim it).
  "ctrl+shift+m",
  "cmd+shift+m",
  // Some keyboards emit both ctrl+meta on the same chord.
  "ctrl+cmd+shift+m",
  // macOS system shortcut: Hide (Cmd+H). Extensions should never be able to claim it.
  "cmd+h",
  // Some keyboards emit both ctrl+meta on the same chord.
  "ctrl+cmd+h",
  // Core file shortcuts (new/open/save/close/quit). Once these are migrated into the
  // KeybindingService, extensions should never be able to claim them.
  "ctrl+n",
  "cmd+n",
  // Some keyboards emit both ctrl+meta on the same chord.
  "ctrl+cmd+n",
  "ctrl+o",
  "cmd+o",
  // Some keyboards emit both ctrl+meta on the same chord.
  "ctrl+cmd+o",
  "ctrl+s",
  "cmd+s",
  // Some keyboards emit both ctrl+meta on the same chord.
  "ctrl+cmd+s",
  "ctrl+shift+s",
  "cmd+shift+s",
  // Some keyboards emit both ctrl+meta on the same chord.
  "ctrl+cmd+shift+s",
  // Print (Ctrl/Cmd+P).
  "ctrl+p",
  "cmd+p",
  // Some keyboards emit both ctrl+meta on the same chord.
  "ctrl+cmd+p",
  "ctrl+w",
  "cmd+w",
  // Some keyboards emit both ctrl+meta on the same chord.
  "ctrl+cmd+w",
  "ctrl+q",
  "cmd+q",
  // Some keyboards emit both ctrl+meta on the same chord.
  "ctrl+cmd+q",
];

export class KeybindingService {
  private readonly platform: "mac" | "other";
  private readonly reservedExtensionShortcuts: ParsedKeybinding[];
  private readonly ignoreInputTargets: "all" | "extensions" | "none";

  private builtinKeybindings: BuiltinKeybinding[] = [];
  private extensionKeybindings: ContributedKeybinding[] = [];

  private builtin: StoredKeybinding[] = [];
  private extensions: StoredKeybinding[] = [];
  private orderCounter = 0;

  // Shared `commandId -> [displayKeybinding]` index for UI (palette, menus).
  private readonly commandKeybindingDisplayIndex = new Map<string, string[]>();
  // Shared `commandId -> [ariaKeyShortcuts]` index for accessibility metadata.
  private readonly commandKeybindingAriaIndex = new Map<string, string[]>();

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
      /**
       * Keybinding chords that extensions should never be allowed to claim.
       *
       * Defaults to reserving a set of core UX shortcuts (cancel/activation keys,
       * clipboard chords, command palette/quick open, core file shortcuts, etc).
       */
      reservedShortcuts?: string[];
      platform?: "mac" | "other";
      /**
       * Determines whether keydown events originating from text input targets should be ignored.
       *
       * - "all" (default): ignore keybindings completely when the target is an INPUT/TEXTAREA/contenteditable.
       * - "extensions": allow built-in keybindings, but prevent extensions from matching.
       * - "none": allow both built-ins and extensions to match.
       */
      ignoreInputTargets?: "all" | "extensions" | "none";
    },
  ) {
    this.platform = params.platform ?? detectPlatform();
    this.ignoreInputTargets = params.ignoreInputTargets ?? "all";
    this.reservedExtensionShortcuts = (params.reservedShortcuts ?? DEFAULT_RESERVED_EXTENSION_SHORTCUTS)
      .map((binding) => parseKeybinding("__reserved__", binding, null))
      .filter((binding): binding is ParsedKeybinding => binding != null);
  }

  setBuiltinKeybindings(bindings: BuiltinKeybinding[]): void {
    this.builtinKeybindings = [...bindings];
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

      // On non-mac platforms, also accept the macOS binding as an alternate (typically `Meta+...`
      // in the browser). This keeps Cmd-only shortcuts testable in Playwright on Linux/Windows,
      // while still keeping the UI/display binding platform-specific.
      if (this.platform === "other" && kb.mac && kb.mac !== primary) {
        candidates.add(kb.mac);
      }

      for (const raw of candidates) {
        const parsed = parseKeybinding(kb.command, raw, kb.when ?? null);
        if (!parsed) continue;
        next.push({
          source: { kind: "builtin" },
          binding: parsed,
          weight: typeof kb.weight === "number" ? kb.weight : 0,
          order: ++this.orderCounter,
          allowRepeat: kb.allowRepeat === true,
        });
      }
    }
    this.builtin = next;
    this.rebuildCommandKeybindingDisplayIndex();
  }

  setExtensionKeybindings(bindings: ContributedKeybinding[]): void {
    const filtered: ContributedKeybinding[] = [];
    const next: StoredKeybinding[] = [];
    for (const kb of bindings) {
      const raw = pickPlatformKeybinding(kb, this.platform);
      const parsed = parseKeybinding(kb.command, raw, kb.when ?? null);
      if (!parsed) continue;
      // Extensions should not be allowed to claim shortcuts that we reserve for core UX
      // (clipboard chords, command palette, etc). Filter these out early so UI surfaces
      // don't advertise keybindings that will never fire.
      if (this.isReservedExtensionKeybinding(parsed)) continue;
      filtered.push(kb);
      next.push({
        source: { kind: "extension", extensionId: kb.extensionId },
        binding: parsed,
        weight: 0,
        order: ++this.orderCounter,
        allowRepeat: false,
      });
    }
    this.extensionKeybindings = filtered;
    this.extensions = next;
    this.rebuildCommandKeybindingDisplayIndex();
  }

  getCommandKeybindingDisplayIndex(): Map<string, string[]> {
    return this.commandKeybindingDisplayIndex;
  }

  getCommandKeybindingAriaIndex(): Map<string, string[]> {
    return this.commandKeybindingAriaIndex;
  }

  /**
   * Install the global keydown listener (bubble phase by default).
   */
  installWindowListener(
    target: Window = window,
    opts: { capture?: boolean; allowBuiltins?: boolean; allowExtensions?: boolean } = {},
  ): () => void {
    this.dispose();
    const { capture = false, allowBuiltins, allowExtensions } = opts;

    // When installed in capture phase, never run extension-contributed keybindings in capture.
    // Instead, run built-ins in capture and extensions in bubble.
    if (capture) {
      const captureHandler = (e: KeyboardEvent) => {
        if (allowBuiltins === false) return;
        void this.dispatchKeydown(e, { allowBuiltins: true, allowExtensions: false }).catch(() => {});
      };
      const bubbleHandler = (e: KeyboardEvent) => {
        if (allowExtensions === false) return;
        void this.dispatchKeydown(e, { allowBuiltins: false, allowExtensions: true }).catch(() => {});
      };
      target.addEventListener("keydown", captureHandler, { capture: true });
      target.addEventListener("keydown", bubbleHandler, { capture: false });
      this.removeListener = () => {
        target.removeEventListener("keydown", captureHandler, { capture: true });
        target.removeEventListener("keydown", bubbleHandler, { capture: false });
      };
      return () => this.dispose();
    }

    const handler = (e: KeyboardEvent) => {
      void this.dispatchKeydown(e, { allowBuiltins, allowExtensions }).catch(() => {});
    };
    target.addEventListener("keydown", handler, { capture: false });
    this.removeListener = () => target.removeEventListener("keydown", handler, { capture: false });
    return () => this.dispose();
  }

  dispose(): void {
    this.removeListener?.();
    this.removeListener = null;
  }

  /**
   * Synchronous helper for keydown listeners. Dispatches asynchronously.
   *
   * Returns `true` when handled and calls `preventDefault()`.
   */
  handleKeydown(
    event: KeyboardEvent,
    opts: { allowBuiltins?: boolean; allowExtensions?: boolean } = {},
  ): boolean {
    if (event.defaultPrevented) return false;
    if (isInsideKeybindingBarrier(event.target)) return false;
    // Prefer context keys for focus/input state so callers can centralize the logic
    // (e.g. via `installKeyboardContextKeys`) rather than scattering DOM checks.
    // Fall back to `event.target` for environments that don't maintain these keys.
    const focusInTextInput = this.params.contextKeys.get("focus.inTextInput");
    const inputTarget = typeof focusInTextInput === "boolean" ? focusInTextInput : isInputTarget(event.target);
    if (inputTarget && this.ignoreInputTargets === "all") return false;

    const match = this.findMatchingBinding(event, {
      ...opts,
      allowExtensions:
        (opts.allowExtensions ?? true) && !(inputTarget && this.ignoreInputTargets === "extensions"),
    });
    if (!match) return false;
    if (event.repeat && !match.allowRepeat) return false;

    event.preventDefault();
    void (async () => {
      try {
        await this.executeMatch(match);
      } catch (err) {
        try {
          this.params.onCommandError?.(match.binding.command, err);
        } catch {
          // ignore
        }
      }
    })().catch(() => {});
    return true;
  }

  async dispatchKeydown(
    event: KeyboardEvent,
    opts: { allowBuiltins?: boolean; allowExtensions?: boolean } = {},
  ): Promise<boolean> {
    if (event.defaultPrevented) return false;
    if (isInsideKeybindingBarrier(event.target)) return false;
    const focusInTextInput = this.params.contextKeys.get("focus.inTextInput");
    const inputTarget = typeof focusInTextInput === "boolean" ? focusInTextInput : isInputTarget(event.target);
    if (inputTarget && this.ignoreInputTargets === "all") return false;

    const match = this.findMatchingBinding(event, {
      ...opts,
      allowExtensions:
        (opts.allowExtensions ?? true) && !(inputTarget && this.ignoreInputTargets === "extensions"),
    });
    if (!match) return false;
    // Avoid repeatedly firing commands when the user holds a key down (e.g. toggles like
    // command palette / AI chat). Some commands (e.g. sheet navigation) explicitly opt into
    // repeated dispatch.
    if (event.repeat && !match.allowRepeat) return false;

    event.preventDefault();
    try {
      await this.executeMatch(match);
    } catch (err) {
      try {
        this.params.onCommandError?.(match.binding.command, err);
      } catch {
        // ignore
      }
    }
    return true;
  }

  private async executeMatch(match: StoredKeybinding): Promise<void> {
    await this.params.onBeforeExecuteCommand?.(match.binding.command, match.source);
    await this.params.commandRegistry.executeCommand(match.binding.command);
  }

  private findMatchingBinding(
    event: KeyboardEvent,
    opts: { allowBuiltins?: boolean; allowExtensions?: boolean } = {},
  ): StoredKeybinding | null {
    const allowBuiltins = opts.allowBuiltins ?? true;
    const allowExtensions = opts.allowExtensions ?? true;
    const lookup = this.params.contextKeys.asLookup();

    // Built-ins always win when enabled.
    if (allowBuiltins) {
      const builtin = this.findFirstMatch(this.builtin, event, lookup);
      if (builtin) return builtin;
    }

    if (!allowExtensions) return null;

    // Safety net: reserved shortcuts should never be claimed by extensions.
    if (this.isReservedForExtensions(event)) return null;

    return this.findFirstMatch(this.extensions, event, lookup);
  }

  private isReservedForExtensions(event: KeyboardEvent): boolean {
    return this.reservedExtensionShortcuts.some((binding) => matchesKeybinding(binding, event));
  }

  private isReservedExtensionKeybinding(binding: ParsedKeybinding): boolean {
    return this.reservedExtensionShortcuts.some(
      (reserved) =>
        reserved.ctrl === binding.ctrl &&
        reserved.shift === binding.shift &&
        reserved.alt === binding.alt &&
        reserved.meta === binding.meta &&
        reserved.key === binding.key,
    );
  }

  private rebuildCommandKeybindingDisplayIndex(): void {
    const next = buildCommandKeybindingDisplayIndex({
      platform: this.platform,
      builtin: this.builtinKeybindings,
      contributed: this.extensionKeybindings,
    });

    const nextAria = buildCommandKeybindingAriaIndex({
      platform: this.platform,
      builtin: this.builtinKeybindings,
      contributed: this.extensionKeybindings,
    });

    // Preserve identity so UI surfaces can hold onto a stable map reference.
    this.commandKeybindingDisplayIndex.clear();
    for (const [commandId, bindings] of next.entries()) {
      this.commandKeybindingDisplayIndex.set(commandId, bindings);
    }

    this.commandKeybindingAriaIndex.clear();
    for (const [commandId, bindings] of nextAria.entries()) {
      this.commandKeybindingAriaIndex.set(commandId, bindings);
    }
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
      // Avoid repeatedly firing commands when the user holds a key down (e.g. toggles like
      // command palette / AI chat). Bindings must explicitly opt into repeat behavior.
      if (event.repeat && !entry.allowRepeat) continue;
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
