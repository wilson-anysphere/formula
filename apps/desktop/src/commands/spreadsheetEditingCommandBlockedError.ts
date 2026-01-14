export const SPREADSHEET_EDITING_COMMAND_BLOCKED_ERROR_NAME = "SpreadsheetEditingCommandBlockedError";

/**
 * Error thrown when a command is blocked because the spreadsheet is in "edit mode"
 * (cell editor / formula bar / inline edit / split-view secondary editor).
 *
 * This is used to coordinate UX across entry points (ribbon, command palette,
 * keybindings) without double-toasting.
 */
export class SpreadsheetEditingCommandBlockedError extends Error {
  readonly commandId: string;

  constructor(commandId: string, message = "Finish editing to use this command.") {
    super(message);
    this.name = SPREADSHEET_EDITING_COMMAND_BLOCKED_ERROR_NAME;
    this.commandId = String(commandId);
  }
}

export function isSpreadsheetEditingCommandBlockedError(err: unknown): err is SpreadsheetEditingCommandBlockedError {
  return (
    Boolean(err) &&
    typeof err === "object" &&
    (err as any).name === SPREADSHEET_EDITING_COMMAND_BLOCKED_ERROR_NAME &&
    typeof (err as any).commandId === "string"
  );
}

