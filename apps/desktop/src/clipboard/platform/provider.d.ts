import type { ClipboardContent, ClipboardWritePayload } from "../types.js";

export type ClipboardProvider = {
  read: () => Promise<ClipboardContent>;
  write: (payload: ClipboardWritePayload) => Promise<void>;
};

export function createClipboardProvider(): Promise<ClipboardProvider>;

