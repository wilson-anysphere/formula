import type { CommandRegistry } from "../extensions/commandRegistry.js";
import type { RibbonActions } from "./ribbonSchema.js";

type FocusMode = "none" | "focus" | "queueFocus" | "after";

type RoutedRibbonCommand =
  | {
      kind: "execute";
      commandId: string;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      args?: any[];
      focus?: FocusMode;
    }
  | { kind: "ignore" };

function routeRibbonToggle(commandId: string, pressed: boolean): RoutedRibbonCommand | null {
  switch (commandId) {
    // Bold/Italic/Underline toggles are routed via `onCommand` (matching main.ts), but
    // still trigger `onToggle` from the ribbon UI. Treat them as handled so fallbacks
    // don't execute twice.
    case "format.toggleBold":
    case "format.toggleItalic":
    case "format.toggleUnderline":
      return { kind: "ignore" };
    case "format.toggleWrapText":
      return { kind: "execute", commandId: "format.toggleWrapText", args: [pressed] };
    default:
      return null;
  }
}

function routeRibbonCommand(commandId: string): RoutedRibbonCommand | null {
  // Toggle buttons invoke both `onToggle` and `onCommand`. Some toggles are routed via
  // `onToggle` (wrap text), while others are routed via `onCommand` (bold/italic/underline).
  if (commandId === "format.toggleWrapText") return { kind: "ignore" };

  // Prefer canonical CommandRegistry ids when available (matches ribbon schema convention).
  if (commandId.startsWith("clipboard.")) {
    return { kind: "execute", commandId };
  }

  if (commandId.startsWith("format.")) {
    return { kind: "execute", commandId };
  }

  switch (commandId) {
    case "home.font.fontSize":
      return { kind: "execute", commandId: "format.fontSize.set" };
    case "home.font.fontColor":
      return { kind: "execute", commandId: "format.fontColor" };
    case "home.font.fillColor":
      return { kind: "execute", commandId: "format.fillColor" };

    case "edit.find":
    case "edit.replace":
    case "navigation.goTo":
      return { kind: "execute", commandId };
  }

  const fillColorPrefix = "home.font.fillColor.";
  if (commandId.startsWith(fillColorPrefix)) {
    const preset = commandId.slice(fillColorPrefix.length);
    if (preset === "moreColors") return { kind: "execute", commandId: "format.fillColor" };
    if (preset === "none" || preset === "noFill") return { kind: "execute", commandId: "format.fillColor", args: [null] };
    const argb = (() => {
      switch (preset) {
        case "lightGray":
          return "#FFD9D9D9";
        case "yellow":
          return "#FFFFFF00";
        case "blue":
          return "#FF0000FF";
        case "green":
          return "#FF00FF00";
        case "red":
          return "#FFFF0000";
        default:
          return null;
      }
    })();
    if (!argb) return null;
    return { kind: "execute", commandId: "format.fillColor", args: [argb] };
  }

  const fontColorPrefix = "home.font.fontColor.";
  if (commandId.startsWith(fontColorPrefix)) {
    const preset = commandId.slice(fontColorPrefix.length);
    if (preset === "moreColors") return { kind: "execute", commandId: "format.fontColor" };
    if (preset === "automatic") return { kind: "execute", commandId: "format.fontColor", args: [null] };
    const argb = (() => {
      switch (preset) {
        case "black":
          return "#FF000000";
        case "blue":
          return "#FF0000FF";
        case "green":
          return "#FF00FF00";
        case "red":
          return "#FFFF0000";
        default:
          return null;
      }
    })();
    if (!argb) return null;
    return { kind: "execute", commandId: "format.fontColor", args: [argb] };
  }

  const numberFormatPrefix = "home.number.numberFormat.";
  if (commandId.startsWith(numberFormatPrefix)) {
    const kind = commandId.slice(numberFormatPrefix.length);
    if (kind === "currency" || kind === "accounting") return { kind: "execute", commandId: "format.numberFormat.currency" };
    if (kind === "percentage") return { kind: "execute", commandId: "format.numberFormat.percent" };
    if (kind === "shortDate" || kind === "longDate") return { kind: "execute", commandId: "format.numberFormat.date" };
  }

  return null;
}

export function createRibbonActions(params: {
  commandRegistry: CommandRegistry;
  onError?: (err: unknown) => void;
  focusGrid?: (() => void) | null;
  queueFocusGrid?: (() => void) | null;
  onCommandFallback?: ((commandId: string) => void) | null;
  onToggleFallback?: ((commandId: string, pressed: boolean) => void) | null;
}): RibbonActions {
  const {
    commandRegistry,
    onError,
    focusGrid = null,
    queueFocusGrid = null,
    onCommandFallback = null,
    onToggleFallback = null,
  } = params;

  const execute = (id: string, args: any[] | undefined, focus: FocusMode | undefined = "none") => {
    const promise = commandRegistry.executeCommand(id, ...(args ?? []));
    promise.catch((err) => onError?.(err));

    if (focus === "focus") {
      focusGrid?.();
      return;
    }
    if (focus === "queueFocus") {
      queueFocusGrid?.();
      return;
    }
    if (focus === "after") {
      promise.finally(() => focusGrid?.());
    }
  };

  return {
    onToggle: (commandId, pressed) => {
      const routed = routeRibbonToggle(commandId, pressed);
      if (!routed) return onToggleFallback?.(commandId, pressed);
      if (routed.kind === "ignore") return;
      execute(routed.commandId, routed.args, routed.focus);
    },
    onCommand: (commandId) => {
      const routed = routeRibbonCommand(commandId);
      if (!routed) return onCommandFallback?.(commandId);
      if (routed.kind === "ignore") return;
      execute(routed.commandId, routed.args, routed.focus);
    },
  };
}
