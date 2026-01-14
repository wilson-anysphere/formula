export function resolveHomeEditingClearCommandTarget(commandId: string): string | null {
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

