export const DEFAULT_MAX_CLIPBOARD_PARSE_CELLS = 200_000;
export const DEFAULT_MAX_CLIPBOARD_HTML_CHARS = 10_000_000;

export class ClipboardParseLimitError extends Error {
  constructor(message) {
    super(message);
    this.name = "ClipboardParseLimitError";
  }
}

