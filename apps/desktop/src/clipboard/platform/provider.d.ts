import type { ClipboardContent, ClipboardWritePayload } from "../types.js";

export type ClipboardProvider = {
  read: () => Promise<ClipboardContent>;
  write: (payload: ClipboardWritePayload) => Promise<void>;
};

export const CLIPBOARD_LIMITS: {
  maxImageBytes: number;
  maxRichTextBytes: number;
  maxPlainTextWriteBytes: number;
};

export function createClipboardProvider(): Promise<ClipboardProvider>;
