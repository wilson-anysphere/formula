export type ClipboardContent = { text?: string; html?: string; rtf?: string; imagePng?: Uint8Array };
export type ClipboardWritePayload = { text: string; html?: string; rtf?: string; imagePng?: Uint8Array };

export type ClipboardProvider = {
  read(): Promise<ClipboardContent>;
  write(payload: ClipboardWritePayload): Promise<void>;
};

export function createClipboardProvider(): Promise<ClipboardProvider>;

export function parseClipboardContentToCellGrid(content: ClipboardContent): any[] | null;

export function copyRangeToClipboardPayload(
  doc: any,
  sheetId: string,
  range: any,
  options?: { dlp?: { documentId: string; classificationStore: any; policy: any } }
): ClipboardWritePayload;
