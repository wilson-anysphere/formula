import type { CommandRegistry } from "../extensions/commandRegistry.js";
import { t } from "../i18n/index.js";

export const FORMAT_PAINTER_COMMAND_ID = "format.toggleFormatPainter";

export function registerFormatPainterCommand(params: {
  commandRegistry: CommandRegistry;
  isArmed: () => boolean;
  arm: () => void;
  disarm: () => void;
  onCancel?: (() => void) | null;
  /**
   * Optional spreadsheet edit-state predicate. When omitted, Format Painter is assumed runnable.
   *
   * The desktop shell passes a custom predicate (`isSpreadsheetEditing`) that includes split-view
   * secondary editor state so command palette/keybindings cannot bypass ribbon disabling.
   */
  isEditing?: (() => boolean) | null;
  /**
   * Optional spreadsheet read-only predicate. When omitted, Format Painter is assumed runnable.
   *
   * The desktop ribbon disables Format Painter in read-only collab roles; guard execution so
   * command palette/keybindings cannot bypass that state.
   */
  isReadOnly?: (() => boolean) | null;
}): void {
  const {
    commandRegistry,
    isArmed,
    arm,
    disarm,
    onCancel = null,
    isEditing = null,
    isReadOnly = null,
  } = params;
  const isEditingFn = isEditing ?? (() => false);
  const isReadOnlyFn = isReadOnly ?? (() => false);

  commandRegistry.registerBuiltinCommand(
    FORMAT_PAINTER_COMMAND_ID,
    t("command.format.toggleFormatPainter"),
    () => {
      if (isArmed()) {
        disarm();
        onCancel?.();
        return;
      }
      if (isEditingFn()) return;
      if (isReadOnlyFn()) return;
      arm();
    },
    {
      category: t("commandCategory.format"),
      // Make the command palette search friendlier.
      keywords: ["format painter", "format", "painter", "paint format"],
    },
  );
}
