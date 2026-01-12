export type ClipboardContent = {
  text?: string;
  html?: string;
  rtf?: string;
  /**
   * Raw PNG bytes (JS-facing API).
   */
  imagePng?: Uint8Array;
  /**
   * @deprecated Legacy/internal field. Prefer `imagePng`.
   * Base64 is only used as a wire format for Tauri IPC.
   */
  pngBase64?: string;
};

export type ClipboardWritePayload = {
  text: string;
  html?: string;
  rtf?: string;
  /**
   * Preferred API for images: raw PNG bytes (JS-facing).
   *
   * Accepts common byte containers for convenience; callers should provide a
   * `Uint8Array` when possible.
   */
  imagePng?: Uint8Array | ArrayBuffer | ArrayBufferView | Blob;
  /**
   * @deprecated Legacy/internal field. Prefer `imagePng`.
   * Base64 is only used as a wire format for Tauri IPC.
   */
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

export function extractPlainTextFromRtf(rtf: string): string;

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
