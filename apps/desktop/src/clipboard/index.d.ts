export type ClipboardContent = { text?: string; html?: string; rtf?: string; imagePng?: Uint8Array; pngBase64?: string };
export type ClipboardWritePayload = { text: string; html?: string; rtf?: string; pngBase64?: string };

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

export function copyRangeToClipboardPayload(
  doc: any,
  sheetId: string,
  range: any,
  options?: { dlp?: { documentId: string; classificationStore: any; policy: any } }
): ClipboardWritePayload;
