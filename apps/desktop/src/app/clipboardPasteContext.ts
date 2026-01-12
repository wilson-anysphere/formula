import { extractPlainTextFromRtf } from "../clipboard/index.js";

type ClipboardPayload = { text?: string; html?: string; rtf?: string };

export type ClipboardCopyContextLike = { payload: ClipboardPayload };

export type ClipboardContentLike = {
  text?: unknown;
  html?: unknown;
  rtf?: unknown;
  imagePng?: unknown;
  pngBase64?: unknown;
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

function normalizeClipboardRtf(rtf: string): string {
  // Clipboard backends may:
  // - normalize newlines differently across platforms (CRLF vs LF)
  // - append whitespace/newlines after the payload
  // - include NUL terminators (`\0`) at the end of the string
  //
  // Normalize in a similar spirit to `normalizeClipboardText`: normalize newlines and ignore
  // trailing whitespace so "internal paste" detection remains stable.
  let normalized = rtf.replace(/\r\n/g, "\n").replace(/\r/g, "\n").trimEnd();
  // Some backends may include a NUL terminator after the RTF payload. Strip it and trim again in
  // case the terminator appeared after whitespace/newlines.
  normalized = normalized.replace(/\u0000+$/g, "").trimEnd();
  return normalized;
}

function hasUsableClipboardContent(content: ClipboardContentLike): boolean {
  return (
    typeof content.text === "string" ||
    typeof content.html === "string" ||
    typeof content.rtf === "string" ||
    content.imagePng != null ||
    typeof content.pngBase64 === "string"
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

  const rtfMatches = (() => {
    if (typeof content.rtf !== "string") return false;

    if (typeof ctx.payload.rtf === "string" && normalizeClipboardRtf(content.rtf) === normalizeClipboardRtf(ctx.payload.rtf)) {
      return true;
    }

    // Some platforms expose only `text/rtf` on read. If the RTF payload was rewritten (whitespace,
    // font tables, etc) but the extracted TSV matches our internal copy payload, still treat it as
    // an internal paste so formula shifting applies.
    if (typeof ctx.payload.text === "string") {
      const extracted = extractPlainTextFromRtf(content.rtf);
      if (extracted && normalizeClipboardText(extracted) === normalizeClipboardText(ctx.payload.text)) {
        return true;
      }
    }

    return false;
  })();

  return textMatches || htmlMatches || rtfMatches;
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
