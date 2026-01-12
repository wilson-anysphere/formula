type ClipboardPayload = { text?: string; html?: string };

export type ClipboardCopyContextLike = { payload: ClipboardPayload };

export type ClipboardContentLike = {
  text?: unknown;
  html?: unknown;
  rtf?: unknown;
  imagePng?: unknown;
};

function normalizeClipboardText(text: string): string {
  return (
    text
      .replace(/\r\n/g, "\n")
      .replace(/\r/g, "\n")
      // Some clipboard implementations add a trailing newline; ignore it when
      // detecting "internal" pastes for formula shifting.
      .replace(/\n+$/g, "")
  );
}

function hasUsableClipboardContent(content: ClipboardContentLike): boolean {
  return (
    typeof content.text === "string" ||
    typeof content.html === "string" ||
    typeof content.rtf === "string" ||
    content.imagePng != null
  );
}

export function detectInternalPaste(ctx: ClipboardCopyContextLike | null, content: ClipboardContentLike): boolean {
  if (!ctx) return false;

  const textMatches =
    typeof content.text === "string" &&
    typeof ctx.payload.text === "string" &&
    normalizeClipboardText(content.text) === normalizeClipboardText(ctx.payload.text);

  const htmlMatches =
    typeof content.html === "string" && typeof ctx.payload.html === "string" && content.html === ctx.payload.html;

  return textMatches || htmlMatches;
}

/**
 * Reconcile internal clipboard context against the latest clipboard contents.
 *
 * If we have internal copy context but the clipboard has changed (i.e. the paste
 * is not internal), clear the stale context so later clipboard values cannot
 * accidentally match the old snapshot.
 */
export function reconcileClipboardCopyContextForPaste<T extends ClipboardCopyContextLike>(
  ctx: T | null,
  content: ClipboardContentLike
): { isInternalPaste: boolean; nextContext: T | null } {
  const isInternalPaste = detectInternalPaste(ctx, content);

  if (ctx && !isInternalPaste && hasUsableClipboardContent(content)) {
    return { isInternalPaste, nextContext: null };
  }

  return { isInternalPaste, nextContext: ctx };
}

