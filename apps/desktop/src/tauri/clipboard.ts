import { createClipboardProvider } from "../clipboard/platform/provider.js";
import type { ClipboardContent, ClipboardWritePayload } from "../clipboard/index.js";

export type { ClipboardContent, ClipboardWritePayload };

/**
 * @deprecated Prefer calling `createClipboardProvider()` directly from `clipboard/platform/provider.js`.
 *
 * This module remains as a thin compatibility wrapper for older callsites/tests.
 */
export async function readClipboard(): Promise<ClipboardContent> {
  const provider = await createClipboardProvider();
  return provider.read();
}

/**
 * @deprecated Prefer calling `createClipboardProvider()` directly from `clipboard/platform/provider.js`.
 *
 * This module remains as a thin compatibility wrapper for older callsites/tests.
 */
export async function writeClipboard(payload: ClipboardWritePayload): Promise<void> {
  const provider = await createClipboardProvider();
  await provider.write(payload);
}
