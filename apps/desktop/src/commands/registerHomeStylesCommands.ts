import type { SpreadsheetApp } from "../app/spreadsheetApp";
import { showCollabEditRejectedToast } from "../collab/editRejectionToast";
import type { CommandRegistry } from "../extensions/commandRegistry.js";
import type { QuickPickItem } from "../extensions/ui.js";
import { showToast } from "../extensions/ui.js";
import { DEFAULT_FORMATTING_APPLY_CELL_LIMIT, evaluateFormattingSelectionSize } from "../formatting/selectionSizeGuard.js";
import {
  applyGoodBadNeutralCellStyle,
  getGoodBadNeutralCellStyleQuickPickItems,
  GOOD_BAD_NEUTRAL_CELL_STYLE_PRESETS,
  type GoodBadNeutralCellStyleId,
} from "../formatting/cellStyles.js";
import {
  applyFormatAsTablePreset,
  estimateFormatAsTableBandedRowOps,
  FORMAT_AS_TABLE_MAX_BANDED_ROW_OPS,
  type FormatAsTablePresetId,
} from "../formatting/formatAsTablePresets.js";
import type { GridLimits, Range } from "../selection/types";
import { DEFAULT_DESKTOP_LOAD_MAX_COLS, DEFAULT_DESKTOP_LOAD_MAX_ROWS } from "../workbook/load/clampUsedRange.js";

import type { ApplyFormattingToSelection } from "./registerDesktopCommands.js";

export const HOME_STYLES_COMMAND_IDS = {
  cellStyles: {
    goodBadNeutral: "home.styles.cellStyles.goodBadNeutral",
  },
  formatAsTable: {
    light: "home.styles.formatAsTable.light",
    medium: "home.styles.formatAsTable.medium",
    dark: "home.styles.formatAsTable.dark",
  },
} as const;

export function registerHomeStylesCommands(params: {
  commandRegistry: CommandRegistry;
  app: SpreadsheetApp;
  category: string | null;
  applyFormattingToSelection: ApplyFormattingToSelection;
  showQuickPick: <T>(items: QuickPickItem<T>[], options?: { placeHolder?: string }) => Promise<T | null>;
  /**
   * Optional spreadsheet edit-state predicate. When omitted, falls back to `app.isEditing()`.
   *
   * The desktop shell passes a custom predicate that includes split-view secondary editing state.
   */
  isEditing?: (() => boolean) | null;
}): void {
  const { commandRegistry, app, category, applyFormattingToSelection, showQuickPick, isEditing: isEditingParam = null } = params;

  const isEditingActive = (): boolean => {
    if (typeof isEditingParam === "function") return isEditingParam();
    if (typeof (app as any)?.isEditing === "function") return Boolean((app as any).isEditing());
    return false;
  };

  const focusGrid = (): void => {
    try {
      (app as any).focus?.();
    } catch {
      // ignore (tests/headless)
    }
  };

  const safeShowToast = (message: string, type: Parameters<typeof showToast>[1] = "info"): void => {
    try {
      showToast(message, type);
    } catch {
      // ignore (toast root missing in tests/headless)
    }
  };

  const getGridLimitsForFormatting = (): GridLimits => {
    const raw = typeof (app as any)?.getGridLimits === "function" ? (app as any).getGridLimits() : null;
    const maxRows = Number.isInteger(raw?.maxRows) && raw.maxRows > 0 ? raw.maxRows : DEFAULT_DESKTOP_LOAD_MAX_ROWS;
    const maxCols = Number.isInteger(raw?.maxCols) && raw.maxCols > 0 ? raw.maxCols : DEFAULT_DESKTOP_LOAD_MAX_COLS;
    return { maxRows, maxCols };
  };

  commandRegistry.registerBuiltinCommand(
    HOME_STYLES_COMMAND_IDS.cellStyles.goodBadNeutral,
    "Cell Styles: Good, Bad, and Neutral",
    async () => {
      // Formatting actions should never run while the user is editing (primary or split-view secondary editor).
      if (isEditingActive()) return;

      // Guard before prompting so users don't pick a style only to hit the size cap on apply.
      const selectionRaw = typeof (app as any)?.getSelectionRanges === "function" ? (app as any).getSelectionRanges() : [];
      const selection = Array.isArray(selectionRaw) ? (selectionRaw as Range[]) : [];
      const limits = getGridLimitsForFormatting();
      const decision = evaluateFormattingSelectionSize(selection, limits, { maxCells: DEFAULT_FORMATTING_APPLY_CELL_LIMIT });
      if (!decision.allowed) {
        safeShowToast("Selection is too large to format. Try selecting fewer cells or an entire row/column.", "warning");
        focusGrid();
        return;
      }

      // `applyFormattingToSelection` enforces the same read-only band-selection restrictions. Guard
      // early here so users don't pick a style only to be blocked on apply.
      if (typeof (app as any)?.isReadOnly === "function" && (app as any).isReadOnly() === true && !decision.allRangesBand) {
        showCollabEditRejectedToast([{ rejectionKind: "formatDefaults", rejectionReason: "permission" }]);
        focusGrid();
        return;
      }

      const presetId = await showQuickPick<GoodBadNeutralCellStyleId>(getGoodBadNeutralCellStyleQuickPickItems(), {
        placeHolder: "Good, Bad, and Neutral",
      });
      if (!presetId) {
        focusGrid();
        return;
      }

      const presetLabel = GOOD_BAD_NEUTRAL_CELL_STYLE_PRESETS[presetId]?.label ?? "Cell style";
      applyFormattingToSelection(`Cell style: ${presetLabel}`, (doc, sheetId, ranges) =>
        applyGoodBadNeutralCellStyle(doc, sheetId, ranges, presetId),
      );
    },
    {
      category,
      icon: null,
      keywords: ["cell styles", "cell style", "good", "bad", "neutral", "formatting"],
    },
  );

  const registerFormatAsTableCommand = (presetId: FormatAsTablePresetId): void => {
    const title = presetId[0]?.toUpperCase() ? `${presetId[0].toUpperCase()}${presetId.slice(1)}` : presetId;
    commandRegistry.registerBuiltinCommand(
      HOME_STYLES_COMMAND_IDS.formatAsTable[presetId],
      `Format as Table: ${title}`,
      () => {
        // Formatting actions should never run while the user is editing (including split-view secondary editor).
        if (isEditingActive()) return;

        applyFormattingToSelection(
          "Format as Table",
          (doc, sheetId, ranges) => {
            if (ranges.length !== 1) {
              safeShowToast("Format as Table currently supports a single rectangular selection.", "warning");
              return true;
            }

            // `applyFormattingToSelection` allows full row/column band selections (Excel-scale) because
            // many formatting operations are scalable via layered formats. Format-as-table banding
            // requires per-row formatting and would be O(rows), so impose a stricter cap here.
            const range = ranges[0];
            const rowCount = range.end.row - range.start.row + 1;
            const colCount = range.end.col - range.start.col + 1;
            const cellCount = rowCount * colCount;
            const bandedRowOps = estimateFormatAsTableBandedRowOps(rowCount);
            if (cellCount > DEFAULT_FORMATTING_APPLY_CELL_LIMIT || bandedRowOps > FORMAT_AS_TABLE_MAX_BANDED_ROW_OPS) {
              safeShowToast("Format as Table selection is too large. Try selecting fewer rows/columns.", "warning");
              return true;
            }

            return applyFormatAsTablePreset(doc, sheetId, range, presetId);
          },
          { forceBatch: true },
        );
      },
      {
        category,
        icon: null,
        keywords: ["format as table", "table", presetId],
      },
    );
  };

  registerFormatAsTableCommand("light");
  registerFormatAsTableCommand("medium");
  registerFormatAsTableCommand("dark");
}
