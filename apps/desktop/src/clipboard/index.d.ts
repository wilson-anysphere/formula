/**
 * Clipboard content returned by the platform provider.
 *
 * Prefer `imagePng` (raw bytes). `pngBase64` is a legacy/internal field kept only for
 * backwards compatibility and is used as a wire format for Tauri IPC.
 */
export type ClipboardContent = { text?: string; html?: string; rtf?: string; imagePng?: Uint8Array; pngBase64?: string };

/**
 * Payload written by the platform provider.
 *
 * Prefer `imagePng` (raw bytes). `pngBase64` is a legacy/internal field kept only for
 * backwards compatibility and is used as a wire format for Tauri IPC.
 */
export type ClipboardWritePayload = {
  text: string;
  html?: string;
  rtf?: string;
  imagePng?: Uint8Array | Blob;
  pngBase64?: string;
};
export type PasteSpecialMode = "all" | "values" | "formulas" | "formats";

export type ClipboardProvider = {
  read(): Promise<ClipboardContent>;
  write(payload: ClipboardWritePayload): Promise<void>;
};

export function createClipboardProvider(): Promise<ClipboardProvider>;

export function parseClipboardContentToCellGrid(content: ClipboardContent): any[] | null;

export function clipboardFormatToDocStyle(format: any): any | null;

export function getCellGridFromRange(doc: any, sheetId: string, range: any): any[][];

export function serializeCellGridToClipboardPayload(grid: any[][]): ClipboardWritePayload;

export function serializeCellGridToRtf(grid: any[][]): string;

export function pasteClipboardContent(
  doc: any,
  sheetId: string,
  start: any,
  content: ClipboardContent,
  options?: { mode?: PasteSpecialMode },
): boolean;

export function copyRangeToClipboardPayload(
  doc: any,
  sheetId: string,
  range: any,
  options?: { dlp?: { documentId: string; classificationStore: any; policy: any } }
): ClipboardWritePayload;
