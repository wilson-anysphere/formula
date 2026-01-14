import type { DocumentController } from "../document/documentController.js";
import { mergeAcross, mergeCells, mergeCenter, unmergeCells } from "../document/mergedCells.js";
import {
  applyAllBorders,
  applyNumberFormatPreset,
  setFillColor,
  setFontColor,
  setFontSize,
  setHorizontalAlign,
  toggleBold,
  toggleItalic,
  toggleStrikethrough,
  toggleUnderline,
  toggleWrap,
  type CellRange,
} from "../formatting/toolbar.js";
import { DEFAULT_FORMATTING_APPLY_CELL_LIMIT } from "../formatting/selectionSizeGuard.js";
import { getStyleNumberFormat } from "../formatting/styleFieldAccess.js";

export type RibbonFormattingApplyFn = (
  doc: DocumentController,
  sheetId: string,
  ranges: CellRange[],
) => void | boolean;

export type RibbonFormattingApplyToSelection = (
  label: string,
  fn: RibbonFormattingApplyFn,
  options?: { forceBatch?: boolean },
) => void;

export type RibbonQuickPick = <T>(
  items: Array<{ label: string; value: T }>,
  options?: { placeHolder?: string },
) => Promise<T | null>;

export type RibbonOpenColorPicker = (
  input: HTMLInputElement,
  label: string,
  apply: (sheetId: string, ranges: CellRange[], argb: string) => void,
) => void;

export type RibbonCommandHandlerContext = {
  app: {
    getDocument: () => DocumentController;
    getCurrentSheetId: () => string;
    getActiveCell: () => { row: number; col: number };
    getSelectionRanges?: () => Array<{ startRow: number; endRow: number; startCol: number; endCol: number }>;
    focus: () => void;
  };
  /**
   * `main.ts` uses `isSpreadsheetEditing()` which also accounts for split-view secondary editor state.
   * Pass that in here so command handlers can match app-shell behavior.
   */
  isEditing?: () => boolean;
  applyFormattingToSelection: RibbonFormattingApplyToSelection;
  /**
   * Wrapper around the desktop CommandRegistry (or equivalent) so command handlers can
   * delegate to existing builtin commands without importing app-shell code.
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  executeCommand?: (commandId: string, ...args: any[]) => void;
  sortSelection?: (options: { order: "ascending" | "descending" }) => void;
  openCustomSort?: (commandId: string) => void;
  promptCustomNumberFormat?: () => void;
  toggleAutoFilter?: () => void;
  clearAutoFilter?: () => void;
  reapplyAutoFilter?: () => void;
  openFormatCells?: () => void;
  /**
   * Pass-through for app shell concerns. Not currently used by formatting handlers, but
   * included to keep the context extensible without coupling this module to `main.ts`.
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  commandRegistry?: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  layoutController?: any;
  showToast?: (message: string, type?: "info" | "warning" | "error") => void;
};

const FONT_SIZE_STEPS = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];

function activeCellFontSizePt(ctx: RibbonCommandHandlerContext): number {
  const sheetId = ctx.app.getCurrentSheetId();
  const cell = ctx.app.getActiveCell();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const docAny = ctx.app.getDocument() as any;
  const effectiveSize = docAny.getCellFormat?.(sheetId, cell)?.font?.size;
  const state = docAny.getCell?.(sheetId, cell);
  const style = docAny.styleTable?.get?.(state?.styleId ?? 0) ?? {};
  const size = typeof effectiveSize === "number" ? effectiveSize : style.font?.size;
  return typeof size === "number" && Number.isFinite(size) && size > 0 ? size : 11;
}

function activeCellNumberFormat(ctx: RibbonCommandHandlerContext): string | null {
  const sheetId = ctx.app.getCurrentSheetId();
  const cell = ctx.app.getActiveCell();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const docAny = ctx.app.getDocument() as any;
  const style = docAny.getCellFormat?.(sheetId, cell);
  return getStyleNumberFormat(style);
}

function activeCellIndentLevel(ctx: RibbonCommandHandlerContext): number {
  const sheetId = ctx.app.getCurrentSheetId();
  const cell = ctx.app.getActiveCell();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const docAny = ctx.app.getDocument() as any;
  const raw = docAny.getCellFormat?.(sheetId, cell)?.alignment?.indent;
  const value = typeof raw === "number" ? raw : typeof raw === "string" && raw.trim() !== "" ? Number(raw) : 0;
  return Number.isFinite(value) ? Math.max(0, Math.trunc(value)) : 0;
}

function stepFontSize(current: number, direction: "increase" | "decrease"): number {
  const value = Number(current);
  const resolved = Number.isFinite(value) && value > 0 ? value : 11;
  if (direction === "increase") {
    for (const step of FONT_SIZE_STEPS) {
      if (step > resolved + 1e-6) return step;
    }
    return resolved;
  }

  for (let i = FONT_SIZE_STEPS.length - 1; i >= 0; i -= 1) {
    const step = FONT_SIZE_STEPS[i]!;
    if (step < resolved - 1e-6) return step;
  }
  return resolved;
}

function parseDecimalPlaces(format: string): number {
  const dot = format.indexOf(".");
  if (dot === -1) return 0;
  let count = 0;
  for (let i = dot + 1; i < format.length; i++) {
    const ch = format[i];
    if (ch === "0" || ch === "#") count += 1;
    else break;
  }
  return count;
}

function stepDecimalPlacesInNumberFormat(format: string | null, direction: "increase" | "decrease"): string | null {
  const raw = (format ?? "").trim();
  const section = (raw.split(";")[0] ?? "").trim();
  const lower = section.toLowerCase();
  const compact = lower.replace(/\s+/g, "");
  // Avoid trying to manipulate date/time format codes.
  if (compact.includes("m/d/yyyy") || compact.includes("yyyy-mm-dd")) return null;
  if (/^h{1,2}:m{1,2}(:s{1,2})?$/.test(compact)) return null;
  // Avoid mutating explicit text number formats.
  if (compact === "@") return null;

  // Preserve scientific notation when possible (e.g. `0.00E+00`).
  const scientificMatch = /E([+-])([0]+)/i.exec(section);
  if (scientificMatch) {
    const base = section.slice(0, scientificMatch.index);
    const decimals = parseDecimalPlaces(base);
    const nextDecimals = direction === "increase" ? Math.min(10, decimals + 1) : Math.max(0, decimals - 1);
    if (nextDecimals === decimals) return null;

    const expSign = scientificMatch[1] ?? "+";
    const expDigits = scientificMatch[2]?.length ?? 0;
    if (expDigits <= 0) return null;

    const fraction = nextDecimals > 0 ? `.${"0".repeat(nextDecimals)}` : "";
    return `0${fraction}E${expSign}${"0".repeat(expDigits)}`;
  }

  // Preserve classic fraction formats (e.g. `# ?/?`, `# ??/??`) by adjusting the number of
  // `?` placeholders (instead of converting to a decimal format).
  const trimmed = section.trim();
  if (/^#\s+\?+\/\?+$/.test(trimmed)) {
    const slash = trimmed.indexOf("/");
    if (slash === -1) return null;
    const denom = trimmed.slice(slash + 1).trim();
    const digits = denom.length;
    const nextDigits = direction === "increase" ? Math.min(10, digits + 1) : Math.max(1, digits - 1);
    if (nextDigits === digits) return null;
    const qs = "?".repeat(nextDigits);
    return `# ${qs}/${qs}`;
  }

  const currencyMatch = /[$€£¥]/.exec(section);
  const prefix = currencyMatch?.[0] ?? "";
  const suffix = section.includes("%") ? "%" : "";
  const useThousands = section.includes(",");
  const decimals = parseDecimalPlaces(section);

  const nextDecimals = direction === "increase" ? Math.min(10, decimals + 1) : Math.max(0, decimals - 1);
  if (nextDecimals === decimals) return null;

  const integer = useThousands ? "#,##0" : "0";
  const fraction = nextDecimals > 0 ? `.${"0".repeat(nextDecimals)}` : "";
  return `${prefix}${integer}${fraction}${suffix}`;
}

export function handleRibbonToggle(ctx: RibbonCommandHandlerContext, commandId: string, pressed: boolean): boolean {
  switch (commandId) {
    case "home.font.bold":
    case "format.toggleBold":
      ctx.applyFormattingToSelection("Bold", (doc, sheetId, ranges) => toggleBold(doc, sheetId, ranges, { next: pressed }));
      return true;
    case "home.font.italic":
    case "format.toggleItalic":
      ctx.applyFormattingToSelection("Italic", (doc, sheetId, ranges) => toggleItalic(doc, sheetId, ranges, { next: pressed }));
      return true;
    case "home.font.underline":
    case "format.toggleUnderline":
      ctx.applyFormattingToSelection("Underline", (doc, sheetId, ranges) => toggleUnderline(doc, sheetId, ranges, { next: pressed }));
      return true;
    case "home.font.strikethrough":
    case "format.toggleStrikethrough":
      ctx.applyFormattingToSelection("Strikethrough", (doc, sheetId, ranges) =>
        toggleStrikethrough(doc, sheetId, ranges, { next: pressed }),
      );
      return true;
    case "home.alignment.wrapText":
    case "format.toggleWrapText":
      ctx.applyFormattingToSelection("Wrap", (doc, sheetId, ranges) => toggleWrap(doc, sheetId, ranges, { next: pressed }));
      return true;
    default:
      return false;
  }
}

export function handleRibbonCommand(ctx: RibbonCommandHandlerContext, commandId: string): boolean {
  const doc = ctx.app.getDocument();

  // Allow invoking formatting toggles as plain commands (useful for keyboard shortcuts/tests).
  // Note: ribbon toggle buttons also trigger `onToggle`, so `main.ts` must avoid calling this
  // for toggle-originated `onCommand` events to prevent double-toggling.
  switch (commandId) {
    case "home.font.bold":
    case "format.toggleBold":
      ctx.applyFormattingToSelection("Bold", (doc, sheetId, ranges) => toggleBold(doc, sheetId, ranges));
      return true;
    case "home.font.italic":
    case "format.toggleItalic":
      ctx.applyFormattingToSelection("Italic", (doc, sheetId, ranges) => toggleItalic(doc, sheetId, ranges));
      return true;
    case "home.font.underline":
    case "format.toggleUnderline":
      ctx.applyFormattingToSelection("Underline", (doc, sheetId, ranges) => toggleUnderline(doc, sheetId, ranges));
      return true;
    case "format.toggleStrikethrough":
      ctx.applyFormattingToSelection("Strikethrough", (doc, sheetId, ranges) => toggleStrikethrough(doc, sheetId, ranges));
      return true;
    case "home.alignment.wrapText":
    case "format.toggleWrapText":
      ctx.applyFormattingToSelection("Wrap", (doc, sheetId, ranges) => toggleWrap(doc, sheetId, ranges));
      return true;
    case "format.clearFormats":
      ctx.applyFormattingToSelection("Clear formats", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, null, { label: "Clear formats" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    case "format.clearContents":
      ctx.applyFormattingToSelection("Clear contents", (doc, sheetId, ranges) => {
        for (const range of ranges) {
          doc.clearRange(sheetId, range, { label: "Clear contents" });
        }
      });
      return true;
    case "format.clearAll":
      ctx.applyFormattingToSelection(
        "Clear all",
        (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            doc.clearRange(sheetId, range, { label: "Clear all" });
            const ok = doc.setRangeFormat(sheetId, range, null, { label: "Clear all" });
            if (ok === false) applied = false;
          }
          return applied;
        },
        { forceBatch: true },
      );
      return true;
    default:
      break;
  }

  const fontNamePrefix = "format.fontName.";
  if (commandId.startsWith(fontNamePrefix)) {
    const preset = commandId.slice(fontNamePrefix.length);
    const fontName = (() => {
      switch (preset) {
        case "calibri":
          return "Calibri";
        case "arial":
          return "Arial";
        case "times":
          return "Times New Roman";
        case "courier":
          return "Courier New";
        default:
          return null;
      }
    })();
    if (!fontName) return true;
    ctx.applyFormattingToSelection("Font", (doc, sheetId, ranges) => {
      let applied = true;
      for (const range of ranges) {
        const ok = doc.setRangeFormat(sheetId, range, { font: { name: fontName } }, { label: "Font" });
        if (ok === false) applied = false;
      }
      return applied;
    });
    return true;
  }

  const fontSizePrefix = "format.fontSize.";
  if (commandId.startsWith(fontSizePrefix)) {
    const size = Number(commandId.slice(fontSizePrefix.length));
    if (!Number.isFinite(size) || size <= 0) return true;
    ctx.applyFormattingToSelection("Font size", (_doc, sheetId, ranges) => setFontSize(doc, sheetId, ranges, size));
    return true;
  }

  if (commandId === "format.increaseFontSize" || commandId === "format.decreaseFontSize") {
    const direction = commandId === "format.increaseFontSize" ? "increase" : "decrease";
    const current = activeCellFontSizePt(ctx);
    const next = stepFontSize(current, direction);
    if (next !== current) {
      ctx.applyFormattingToSelection("Font size", (_doc, sheetId, ranges) => setFontSize(doc, sheetId, ranges, next));
    }
    return true;
  }

  const fillColorPrefix = "format.fillColor.";
  if (commandId.startsWith(fillColorPrefix)) {
    const preset = commandId.slice(fillColorPrefix.length);
    if (preset === "moreColors") {
      // Prefer delegating to the builtin command (which opens the picker UI).
      ctx.executeCommand?.("format.fillColor.moreColors");
      return true;
    }
    const argb = (() => {
      switch (preset) {
        case "lightGray":
          return ["#", "FF", "D9D9D9"].join("");
        case "yellow":
          return ["#", "FF", "FFFF00"].join("");
        case "blue":
          return ["#", "FF", "0000FF"].join("");
        case "green":
          return ["#", "FF", "00FF00"].join("");
        case "red":
          return ["#", "FF", "FF0000"].join("");
        default:
          return null;
      }
    })();

    if (preset === "none" || preset === "noFill") {
      ctx.applyFormattingToSelection("Fill color", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { fill: null }, { label: "Fill color" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    }

    if (argb) {
      ctx.applyFormattingToSelection("Fill color", (_doc, sheetId, ranges) => setFillColor(doc, sheetId, ranges, argb));
    }
    return true;
  }

  const fontColorPrefix = "format.fontColor.";
  if (commandId.startsWith(fontColorPrefix)) {
    const preset = commandId.slice(fontColorPrefix.length);
    if (preset === "moreColors") {
      // Prefer delegating to the builtin command (which opens the picker UI).
      ctx.executeCommand?.("format.fontColor.moreColors");
      return true;
    }
    const argb = (() => {
      switch (preset) {
        case "black":
          return ["#", "FF", "000000"].join("");
        case "blue":
          return ["#", "FF", "0000FF"].join("");
        case "green":
          return ["#", "FF", "00FF00"].join("");
        case "red":
          return ["#", "FF", "FF0000"].join("");
        default:
          return null;
      }
    })();

    if (preset === "automatic") {
      ctx.applyFormattingToSelection("Font color", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { font: { color: null } }, { label: "Font color" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    }

    if (argb) {
      ctx.applyFormattingToSelection("Font color", (_doc, sheetId, ranges) => setFontColor(doc, sheetId, ranges, argb));
    }
    return true;
  }

  const bordersPrefix = "format.borders.";
  if (commandId.startsWith(bordersPrefix)) {
    const kind = commandId.slice(bordersPrefix.length);
    const defaultBorderColor = ["#", "FF", "000000"].join("");
    if (kind === "none") {
      ctx.applyFormattingToSelection("Borders", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { border: null }, { label: "Borders" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    }

    if (kind === "all") {
      ctx.applyFormattingToSelection("Borders", (_doc, sheetId, ranges) => applyAllBorders(doc, sheetId, ranges));
      return true;
    }

    if (kind === "outside" || kind === "thickBox") {
      const edgeStyle = kind === "thickBox" ? "thick" : "thin";
      const edge = { style: edgeStyle, color: defaultBorderColor };
      ctx.applyFormattingToSelection(
        "Borders",
        (doc, sheetId, ranges) => {
          let applied = true;
          const applyBorder = (targetRange: CellRange, patch: Record<string, any> | null) => {
            const ok = doc.setRangeFormat(sheetId, targetRange, patch, { label: "Borders" });
            if (ok === false) applied = false;
          };
          for (const range of ranges) {
            const startRow = range.start.row;
            const endRow = range.end.row;
            const startCol = range.start.col;
            const endCol = range.end.col;

            // Top edge.
            applyBorder({ start: { row: startRow, col: startCol }, end: { row: startRow, col: endCol } }, { border: { top: edge } });

            // Bottom edge.
            applyBorder({ start: { row: endRow, col: startCol }, end: { row: endRow, col: endCol } }, { border: { bottom: edge } });

            // Left edge.
            applyBorder({ start: { row: startRow, col: startCol }, end: { row: endRow, col: startCol } }, { border: { left: edge } });

            // Right edge.
            applyBorder({ start: { row: startRow, col: endCol }, end: { row: endRow, col: endCol } }, { border: { right: edge } });
          }
          return applied;
        },
        { forceBatch: true },
      );
      return true;
    }

    const edge = { style: "thin", color: defaultBorderColor };
    const borderPatch = (() => {
      switch (kind) {
        case "bottom":
          return { border: { bottom: edge } };
        case "top":
          return { border: { top: edge } };
        case "left":
          return { border: { left: edge } };
        case "right":
          return { border: { right: edge } };
        default:
          return null;
      }
    })();

    if (borderPatch) {
      ctx.applyFormattingToSelection(
        "Borders",
        (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const startRow = range.start.row;
            const endRow = range.end.row;
            const startCol = range.start.col;
            const endCol = range.end.col;

            const targetRange = (() => {
              switch (kind) {
                case "bottom":
                  return { start: { row: endRow, col: startCol }, end: { row: endRow, col: endCol } };
                case "top":
                  return { start: { row: startRow, col: startCol }, end: { row: startRow, col: endCol } };
                case "left":
                  return { start: { row: startRow, col: startCol }, end: { row: endRow, col: startCol } };
                case "right":
                  return { start: { row: startRow, col: endCol }, end: { row: endRow, col: endCol } };
                default:
                  return range;
              }
            })();

            const ok = doc.setRangeFormat(sheetId, targetRange, borderPatch, { label: "Borders" });
            if (ok === false) applied = false;
          }
          return applied;
        },
        { forceBatch: true },
      );
    }
    return true;
  }

  const numberFormatPrefix = "format.numberFormat.";
  if (commandId.startsWith(numberFormatPrefix)) {
    const kind = commandId.slice(numberFormatPrefix.length);
    if (kind === "general") {
      ctx.applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { numberFormat: null }, { label: "Number format" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    }
    if (kind === "number") {
      ctx.applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { numberFormat: "0.00" }, { label: "Number format" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    }
    if (kind === "currency" || kind === "accounting") {
      ctx.applyFormattingToSelection("Number format", (_doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "currency"));
      return true;
    }
    if (kind === "percent" || kind === "percentage") {
      ctx.applyFormattingToSelection("Number format", (_doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "percent"));
      return true;
    }
    if (kind === "date" || kind === "shortDate") {
      ctx.applyFormattingToSelection("Number format", (_doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "date"));
      return true;
    }
    if (kind === "longDate") {
      ctx.applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { numberFormat: "yyyy-mm-dd" }, { label: "Number format" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    }
    if (kind === "time") {
      ctx.applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { numberFormat: "h:mm:ss" }, { label: "Number format" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    }
    if (kind === "commaStyle") {
      ctx.applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { numberFormat: "#,##0.00" }, { label: "Number format" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    }
    if (kind === "increaseDecimal" || kind === "decreaseDecimal") {
      const next = stepDecimalPlacesInNumberFormat(activeCellNumberFormat(ctx), kind === "increaseDecimal" ? "increase" : "decrease");
      if (!next) return true;
      ctx.applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { numberFormat: next }, { label: "Number format" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    }
    if (kind === "fraction") {
      ctx.applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { numberFormat: "# ?/?" }, { label: "Number format" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    }
    if (kind === "scientific") {
      ctx.applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { numberFormat: "0.00E+00" }, { label: "Number format" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    }
    if (kind === "text") {
      ctx.applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { numberFormat: "@" }, { label: "Number format" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    }
    return true;
  }

  const accountingPrefix = "format.numberFormat.accounting.";
  if (commandId.startsWith(accountingPrefix)) {
    const currency = commandId.slice(accountingPrefix.length);
    const symbol = (() => {
      switch (currency) {
        case "eur":
          return "€";
        case "gbp":
          return "£";
        case "jpy":
          return "¥";
        case "usd":
        default:
          return "$";
      }
    })();

    ctx.applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
      let applied = true;
      for (const range of ranges) {
        const ok = doc.setRangeFormat(sheetId, range, { numberFormat: `${symbol}#,##0.00` }, { label: "Number format" });
        if (ok === false) applied = false;
      }
      return applied;
    });
    return true;
  }

  switch (commandId) {
    case "home.font.borders":
      // This command is a dropdown with menu items; the top-level command is not expected
      // to fire when the menu is present. Keep this as a fallback.
      ctx.applyFormattingToSelection("Borders", (_doc, sheetId, ranges) => applyAllBorders(doc, sheetId, ranges));
      return true;
    case "home.font.fontColor":
      ctx.executeCommand?.("format.fontColor");
      return true;
    case "home.font.fillColor":
      ctx.executeCommand?.("format.fillColor");
      return true;
    case "home.font.fontSize":
      ctx.executeCommand?.("format.fontSize.set");
      return true;
    case "home.alignment.alignLeft":
      ctx.applyFormattingToSelection("Align left", (doc, sheetId, ranges) => setHorizontalAlign(doc, sheetId, ranges, "left"));
      return true;
    case "home.alignment.topAlign":
      ctx.applyFormattingToSelection("Vertical align", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { alignment: { vertical: "top" } }, { label: "Vertical align" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    case "home.alignment.middleAlign":
      ctx.applyFormattingToSelection("Vertical align", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          // Spreadsheet vertical alignment uses "center" (Excel/OOXML); the grid maps this to CSS middle.
          const ok = doc.setRangeFormat(sheetId, range, { alignment: { vertical: "center" } }, { label: "Vertical align" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    case "home.alignment.bottomAlign":
      ctx.applyFormattingToSelection("Vertical align", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { alignment: { vertical: "bottom" } }, { label: "Vertical align" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    case "home.alignment.center":
      ctx.applyFormattingToSelection("Align center", (doc, sheetId, ranges) => setHorizontalAlign(doc, sheetId, ranges, "center"));
      return true;
    case "home.alignment.alignRight":
      ctx.applyFormattingToSelection("Align right", (doc, sheetId, ranges) => setHorizontalAlign(doc, sheetId, ranges, "right"));
      return true;
    case "home.alignment.increaseIndent": {
      const current = activeCellIndentLevel(ctx);
      const next = Math.min(250, current + 1);
      if (next === current) return true;
      ctx.applyFormattingToSelection("Indent", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { alignment: { indent: next } }, { label: "Indent" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    }
    case "home.alignment.decreaseIndent": {
      const current = activeCellIndentLevel(ctx);
      const next = Math.max(0, current - 1);
      if (next === current) return true;
      ctx.applyFormattingToSelection("Indent", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { alignment: { indent: next } }, { label: "Indent" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    }
    case "home.alignment.orientation.angleCounterclockwise":
      ctx.applyFormattingToSelection("Text orientation", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { alignment: { textRotation: 45 } }, { label: "Text orientation" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    case "home.alignment.orientation.angleClockwise":
      ctx.applyFormattingToSelection("Text orientation", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { alignment: { textRotation: -45 } }, { label: "Text orientation" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    case "home.alignment.orientation.verticalText":
      ctx.applyFormattingToSelection("Text orientation", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          // Excel/OOXML uses 255 as a sentinel for vertical text (stacked).
          const ok = doc.setRangeFormat(sheetId, range, { alignment: { textRotation: 255 } }, { label: "Text orientation" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    case "home.alignment.orientation.rotateUp":
      ctx.applyFormattingToSelection("Text orientation", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { alignment: { textRotation: 90 } }, { label: "Text orientation" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    case "home.alignment.orientation.rotateDown":
      ctx.applyFormattingToSelection("Text orientation", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { alignment: { textRotation: -90 } }, { label: "Text orientation" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    case "home.alignment.orientation.formatCellAlignment":
      ctx.executeCommand?.("format.openFormatCells");
      ctx.openFormatCells?.();
      return true;

    case "home.alignment.mergeCenter":
      // Dropdown container id; some ribbon interactions can surface this in `onCommand`.
      // Treat it as a no-op fallback (menu items trigger the real commands).
      return true;

    case "home.alignment.mergeCenter.mergeCenter":
    case "home.alignment.mergeCenter.mergeCells":
    case "home.alignment.mergeCenter.mergeAcross": {
      if (ctx.isEditing?.()) return true;

      const selection = ctx.app.getSelectionRanges?.() ?? [];
      if (selection.length > 1) {
        ctx.showToast?.("Merge commands only support a single selection range.", "warning");
        ctx.app.focus();
        return true;
      }

      const normalized = (() => {
        if (selection.length === 0) {
          const cell = ctx.app.getActiveCell();
          return { startRow: cell.row, endRow: cell.row, startCol: cell.col, endCol: cell.col };
        }
        const r = selection[0]!;
        const startRow = Math.min(r.startRow, r.endRow);
        const endRow = Math.max(r.startRow, r.endRow);
        const startCol = Math.min(r.startCol, r.endCol);
        const endCol = Math.max(r.startCol, r.endCol);
        return { startRow, endRow, startCol, endCol };
      })();

      const rows = normalized.endRow - normalized.startRow + 1;
      const cols = normalized.endCol - normalized.startCol + 1;
      const totalCells = rows * cols;
      const maxCells = DEFAULT_FORMATTING_APPLY_CELL_LIMIT;
      if (totalCells > maxCells) {
        ctx.showToast?.(
          `Selection too large to merge (>${maxCells.toLocaleString()} cells). Select fewer cells and try again.`,
          "warning",
        );
        ctx.app.focus();
        return true;
      }

      const sheetId = ctx.app.getCurrentSheetId();
      const label =
        commandId === "home.alignment.mergeCenter.mergeCenter"
          ? "Merge & Center"
          : commandId === "home.alignment.mergeCenter.mergeAcross"
            ? "Merge Across"
            : "Merge Cells";

      // Merge Across is only meaningful for multi-column selections.
      if (commandId === "home.alignment.mergeCenter.mergeAcross" && cols <= 1) {
        ctx.app.focus();
        return true;
      }

      doc.beginBatch({ label });
      let committed = false;
      try {
        if (commandId === "home.alignment.mergeCenter.mergeCenter") {
          mergeCenter(doc, sheetId, normalized, { label });
        } else if (commandId === "home.alignment.mergeCenter.mergeAcross") {
          mergeAcross(doc, sheetId, normalized, { label });
        } else {
          mergeCells(doc, sheetId, normalized, { label });
        }
        committed = true;
      } finally {
        if (committed) doc.endBatch();
        else doc.cancelBatch();
      }

      ctx.app.focus();
      return true;
    }

    case "home.alignment.mergeCenter.unmergeCells": {
      if (ctx.isEditing?.()) return true;

      const selection = ctx.app.getSelectionRanges?.() ?? [];
      if (selection.length > 1) {
        ctx.showToast?.("Unmerge Cells only supports a single selection range.", "warning");
        ctx.app.focus();
        return true;
      }

      const normalized = (() => {
        if (selection.length === 0) {
          const cell = ctx.app.getActiveCell();
          return { startRow: cell.row, endRow: cell.row, startCol: cell.col, endCol: cell.col };
        }
        const r = selection[0]!;
        const startRow = Math.min(r.startRow, r.endRow);
        const endRow = Math.max(r.startRow, r.endRow);
        const startCol = Math.min(r.startCol, r.endCol);
        const endCol = Math.max(r.startCol, r.endCol);
        return { startRow, endRow, startCol, endCol };
      })();

      const sheetId = ctx.app.getCurrentSheetId();
      unmergeCells(doc, sheetId, normalized, { label: "Unmerge Cells" });
      ctx.app.focus();
      return true;
    }
    case "home.number.percent":
      ctx.applyFormattingToSelection("Number format", (doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "percent"));
      return true;
    case "home.number.accounting":
      ctx.applyFormattingToSelection("Number format", (doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "currency"));
      return true;
    case "home.number.date":
      ctx.applyFormattingToSelection("Number format", (doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "date"));
      return true;
    case "home.number.comma":
      ctx.applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { numberFormat: "#,##0.00" }, { label: "Number format" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    case "home.number.moreFormats.custom":
      if (ctx.promptCustomNumberFormat) {
        ctx.promptCustomNumberFormat();
      } else {
        ctx.showToast?.("Custom number formats are not available.", "error");
        ctx.app.focus();
      }
      return true;
    case "home.number.increaseDecimal": {
      const next = stepDecimalPlacesInNumberFormat(activeCellNumberFormat(ctx), "increase");
      if (!next) return true;
      ctx.applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { numberFormat: next }, { label: "Number format" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    }
    case "home.number.decreaseDecimal": {
      const next = stepDecimalPlacesInNumberFormat(activeCellNumberFormat(ctx), "decrease");
      if (!next) return true;
      ctx.applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { numberFormat: next }, { label: "Number format" });
          if (ok === false) applied = false;
        }
        return applied;
      });
      return true;
    }
    case "home.number.formatCells":
    case "home.number.moreFormats.formatCells":
    case "home.cells.format.formatCells":
    case "format.openFormatCells":
      ctx.executeCommand?.("format.openFormatCells");
      ctx.openFormatCells?.();
      return true;

    case "home.editing.sortFilter.sortAtoZ":
    case "data.sortFilter.sortAtoZ":
    case "data.sortFilter.sort.sortAtoZ":
      ctx.sortSelection?.({ order: "ascending" });
      return true;
    case "home.editing.sortFilter.sortZtoA":
    case "data.sortFilter.sortZtoA":
    case "data.sortFilter.sort.sortZtoA":
      ctx.sortSelection?.({ order: "descending" });
      return true;
    case "home.editing.sortFilter.customSort":
    case "data.sortFilter.sort.customSort":
      if (ctx.openCustomSort) {
        ctx.openCustomSort(commandId);
      } else {
        ctx.showToast?.("Custom sort is not available.", "error");
        ctx.app.focus();
      }
      return true;
    case "home.editing.sortFilter.filter":
    case "data.sortFilter.filter":
      if (ctx.toggleAutoFilter) {
        ctx.toggleAutoFilter();
      } else {
        ctx.showToast?.("Filtering is not available.", "error");
        ctx.app.focus();
      }
      return true;
    case "home.editing.sortFilter.clear":
    case "data.sortFilter.clear":
    case "data.sortFilter.advanced.clearFilter":
      if (ctx.clearAutoFilter) {
        ctx.clearAutoFilter();
      } else {
        ctx.showToast?.("Filtering is not available.", "error");
        ctx.app.focus();
      }
      return true;
    case "home.editing.sortFilter.reapply":
    case "data.sortFilter.reapply":
      if (ctx.reapplyAutoFilter) {
        ctx.reapplyAutoFilter();
      } else {
        ctx.showToast?.("Filtering is not available.", "error");
        ctx.app.focus();
      }
      return true;
    default:
      return false;
  }
}
