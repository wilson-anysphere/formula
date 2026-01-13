import type { CommandRegistry } from "../extensions/commandRegistry.js";
import { t } from "../i18n/index.js";

export const FORMAT_PAINTER_COMMAND_ID = "format.toggleFormatPainter";

export function registerFormatPainterCommand(params: {
  commandRegistry: CommandRegistry;
  isArmed: () => boolean;
  arm: () => void;
  disarm: () => void;
  onCancel?: (() => void) | null;
}): void {
  const { commandRegistry, isArmed, arm, disarm, onCancel = null } = params;

  commandRegistry.registerBuiltinCommand(
    FORMAT_PAINTER_COMMAND_ID,
    t("command.format.toggleFormatPainter"),
    () => {
      if (isArmed()) {
        disarm();
        onCancel?.();
        return;
      }
      arm();
    },
    {
      category: t("commandCategory.format"),
      // Make the command palette search friendlier.
      keywords: ["format painter", "format", "painter", "paint format"],
    },
  );
}

