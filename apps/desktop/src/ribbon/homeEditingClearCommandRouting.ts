export type FormatClearCommandId = "format.clearAll" | "format.clearFormats" | "format.clearContents";

/**
 * Home → Editing → Clear dropdown uses ribbon-specific command ids that are not part of the
 * canonical CommandRegistry surface area.
 *
 * We still want these menu items to invoke the already-implemented clear logic (registered as
 * `format.clear*` commands) so the ribbon does not fall back to the default "Ribbon: …" toast.
 */
export function resolveHomeEditingClearCommandTarget(commandId: string): FormatClearCommandId | null {
  switch (commandId) {
    case "home.editing.clear.clearAll":
      return "format.clearAll";
    case "home.editing.clear.clearFormats":
      return "format.clearFormats";
    case "home.editing.clear.clearContents":
      return "format.clearContents";
    default:
      return null;
  }
}

