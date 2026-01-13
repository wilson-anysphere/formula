import { CLIPBOARD_LIMITS } from "../clipboard/platform/provider.js";

export type ClipboardContent = {
  text?: string;
  html?: string;
  rtf?: string;
  /**
   * Raw PNG bytes (JS-facing API).
   */
  imagePng?: Uint8Array;
  /**
   * @deprecated Legacy/internal field.
   *
   * Base64 is only used as a wire format for Tauri IPC (`__TAURI__.core.invoke`).
   * Prefer `imagePng`.
   */
  pngBase64?: string;
};

export type ClipboardWritePayload = {
  text?: string;
  html?: string;
  rtf?: string;
  /**
   * Raw PNG bytes (JS-facing API).
   *
   * Accepts common byte containers for convenience; callers should provide a
   * `Uint8Array` when possible.
   */
  imagePng?: Uint8Array | ArrayBuffer | ArrayBufferView | Blob;
  /**
   * @deprecated Legacy/internal field.
   *
   * Base64 is only used as a wire format for Tauri IPC (`__TAURI__.core.invoke`).
   * Prefer `imagePng`.
   */
  pngBase64?: string;
};

type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;

// NOTE: Keep this in sync with the Rust backend (`apps/desktop/src-tauri/src/clipboard/mod.rs`).
// Prefer reusing the platform provider's exported caps so we don't accidentally drift.
const MAX_IMAGE_BYTES = CLIPBOARD_LIMITS.maxImageBytes;

const isTrimChar = (code: number) => code === 0x20 || code === 0x09 || code === 0x0a || code === 0x0d; // space, tab, lf, cr

function hasDataUrlPrefixAt(str: string, start: number): boolean {
  if (start + 5 > str.length) return false;
  // ASCII case-insensitive match for "data:" without allocating.
  return (
    ((str.charCodeAt(start) | 32) === 0x64) && // d
    ((str.charCodeAt(start + 1) | 32) === 0x61) && // a
    ((str.charCodeAt(start + 2) | 32) === 0x74) && // t
    ((str.charCodeAt(start + 3) | 32) === 0x61) && // a
    str.charCodeAt(start + 4) === 0x3a // :
  );
}

function base64Bounds(base64: string): { start: number; end: number } {
  let start = 0;
  while (start < base64.length && isTrimChar(base64.charCodeAt(start))) start += 1;

  if (hasDataUrlPrefixAt(base64, start)) {
    // Scan only a small prefix for the comma separator so malformed inputs like
    // `data:AAAAA...` don't force an O(n) search over huge payload strings.
    let comma = -1;
    const maxHeaderScan = Math.min(base64.length, start + 1024);
    for (let i = start; i < maxHeaderScan; i += 1) {
      if (base64.charCodeAt(i) === 0x2c) {
        comma = i;
        break;
      }
    }
    if (comma >= 0) {
      start = comma + 1;
    } else {
      // Malformed data URL (missing comma separator). Treat as empty so we don't accidentally
      // decode `data:...` as base64.
      return { start: base64.length, end: base64.length };
    }
    while (start < base64.length && isTrimChar(base64.charCodeAt(start))) start += 1;
  }

  let end = base64.length;
  while (end > start && isTrimChar(base64.charCodeAt(end - 1))) end -= 1;
  return { start, end };
}

function getTauriInvoke(): TauriInvoke {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  if (!invoke) {
    throw new Error("Tauri invoke API not available");
  }
  return invoke;
}

function normalizeBase64String(base64: string): string {
  if (!base64) return "";
  const { start, end } = base64Bounds(base64);
  if (end <= start) return "";
  return base64.slice(start, end);
}

function estimateBase64Bytes(base64: string): number {
  if (!base64) return 0;
  const { start, end } = base64Bounds(base64);
  const len = end - start;
  if (len <= 0) return 0;

  let padding = 0;
  if (base64.charCodeAt(end - 1) === 0x3d) {
    padding = 1;
    if (end - 2 >= start && base64.charCodeAt(end - 2) === 0x3d) padding = 2;
  }

  const bytes = Math.floor((len * 3) / 4) - padding;
  return bytes > 0 ? bytes : 0;
}

function decodeBase64ToBytes(val: string): Uint8Array | undefined {
  if (!val) return undefined;
  if (estimateBase64Bytes(val) > MAX_IMAGE_BYTES) return undefined;

  const base64 = normalizeBase64String(val);
  if (!base64) return undefined;

  try {
    if (typeof atob === "function") {
      const bin = atob(base64);
      const out = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
      return out;
    }

    if (typeof Buffer !== "undefined") {
      const buf = Buffer.from(base64, "base64");
      if (buf.byteLength > MAX_IMAGE_BYTES) return undefined;
      return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
    }
  } catch {
    return undefined;
  }

  return undefined;
}

function encodeBytesToBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(bytes).toString("base64");
  }

  if (typeof btoa !== "function") {
    throw new Error("base64 encoding is unavailable in this environment");
  }

  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    // eslint-disable-next-line unicorn/prefer-code-point
    binary += String.fromCharCode(...chunk);
  }
  return btoa(binary);
}

async function normalizeImagePngBytes(val: ClipboardWritePayload["imagePng"]): Promise<Uint8Array | undefined> {
  if (!val) return undefined;

  if (typeof Blob !== "undefined" && val instanceof Blob) {
    if (typeof val.size === "number" && val.size > MAX_IMAGE_BYTES) return undefined;
    try {
      const buf = await val.arrayBuffer();
      if (buf.byteLength > MAX_IMAGE_BYTES) return undefined;
      return new Uint8Array(buf);
    } catch {
      return undefined;
    }
  }

  if (val instanceof Uint8Array) {
    if (val.byteLength > MAX_IMAGE_BYTES) return undefined;
    return val;
  }

  if (val instanceof ArrayBuffer) {
    if (val.byteLength > MAX_IMAGE_BYTES) return undefined;
    return new Uint8Array(val);
  }

  if (ArrayBuffer.isView(val) && val.buffer instanceof ArrayBuffer) {
    if (val.byteLength > MAX_IMAGE_BYTES) return undefined;
    return new Uint8Array(val.buffer, val.byteOffset, val.byteLength);
  }

  return undefined;
}

function readPngBase64(source: any): string | undefined {
  if (!source || typeof source !== "object") return undefined;
  if (typeof source.pngBase64 === "string") return source.pngBase64;
  if (typeof source.png_base64 === "string") return source.png_base64;
  if (typeof source.image_png_base64 === "string") return source.image_png_base64;
  return undefined;
}

export async function readClipboard(): Promise<ClipboardContent> {
  const invoke = getTauriInvoke();
  /** @type {any} */
  let payload: any = null;
  try {
    payload = await invoke("clipboard_read");
  } catch {
    // Older desktop builds exposed `read_clipboard`.
    payload = await invoke("read_clipboard");
  }

  const out: ClipboardContent = {};
  if (payload && typeof payload === "object") {
    if (typeof payload.text === "string") out.text = payload.text;
    if (typeof payload.html === "string") out.html = payload.html;
    if (typeof payload.rtf === "string") out.rtf = payload.rtf;

    const pngBase64 = readPngBase64(payload);
    if (typeof pngBase64 === "string") {
      const estimate = estimateBase64Bytes(pngBase64);
      if (estimate > MAX_IMAGE_BYTES) return out;

      const bytes = decodeBase64ToBytes(pngBase64);
      if (bytes) {
        out.imagePng = bytes;
      } else if (estimate <= MAX_IMAGE_BYTES) {
        // Preserve base64 only when decoding fails (legacy/internal).
        const normalized = normalizeBase64String(pngBase64);
        if (normalized) out.pngBase64 = normalized;
      }
    }
  }

  return out;
}

export async function writeClipboard(payload: ClipboardWritePayload): Promise<void> {
  const invoke = getTauriInvoke();
  const imageBytes = await normalizeImagePngBytes(payload.imagePng);

  const pngBase64FromImage = imageBytes ? encodeBytesToBase64(imageBytes) : undefined;
  const legacyPngBase64Raw =
    typeof payload.pngBase64 === "string" && estimateBase64Bytes(payload.pngBase64) <= MAX_IMAGE_BYTES
      ? normalizeBase64String(payload.pngBase64)
      : undefined;
  const legacyPngBase64 = legacyPngBase64Raw ? legacyPngBase64Raw : undefined;

  const pngBase64 = pngBase64FromImage ?? legacyPngBase64;

  /** @type {Record<string, unknown>} */
  const out: Record<string, unknown> = {};
  if (typeof payload.text === "string") out.text = payload.text;
  if (typeof payload.html === "string") out.html = payload.html;
  if (typeof payload.rtf === "string") out.rtf = payload.rtf;
  if (typeof pngBase64 === "string" && pngBase64) out.pngBase64 = pngBase64;

  try {
    await invoke("clipboard_write", { payload: out });
  } catch {
    // Older desktop builds exposed `write_clipboard` with positional args.
    await invoke("write_clipboard", {
      text: typeof payload.text === "string" ? payload.text : "",
      html: typeof payload.html === "string" ? payload.html : undefined,
      rtf: typeof payload.rtf === "string" ? payload.rtf : undefined,
      image_png_base64: typeof pngBase64 === "string" && pngBase64 ? pngBase64 : undefined,
    });
  }
}
