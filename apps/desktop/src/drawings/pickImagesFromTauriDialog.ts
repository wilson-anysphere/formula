import { MAX_INSERT_IMAGE_BYTES } from "./insertImageLimits.js";
import { getTauriDialogOpenOrNull, getTauriInvokeOrThrow } from "../tauri/api";

export const IMAGE_FILE_EXTENSIONS = ["png", "jpg", "jpeg", "gif", "bmp", "webp", "svg"] as const;

type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;

export class FileTooLargeError extends Error {
  readonly fileSize: number;
  readonly maxSize: number;

  constructor(fileSize: number, maxSize: number) {
    super(`File is too large (${fileSize} bytes, max ${maxSize}).`);
    this.name = "FileTooLargeError";
    this.fileSize = fileSize;
    this.maxSize = maxSize;
  }
}
export type PickImagesFromTauriDialogOptions = {
  multiple?: boolean;
};

function normalizeOpenPaths(payload: unknown): string[] {
  if (payload == null) return [];
  if (Array.isArray(payload)) {
    return payload
      .map((v) => (typeof v === "string" ? v.trim() : ""))
      .filter((v): v is string => v !== "");
  }
  if (typeof payload === "string") {
    const trimmed = payload.trim();
    if (trimmed) return [trimmed];
  }
  return [];
}

function normalizeBinaryPayload(payload: unknown): Uint8Array {
  if (typeof payload === "string") {
    if (typeof Buffer !== "undefined") {
      // Node (and some bundlers) provide Buffer.
      // eslint-disable-next-line no-undef
      return new Uint8Array(Buffer.from(payload, "base64"));
    }
    if (typeof atob === "function") {
      const binary = atob(payload);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      return bytes;
    }
    throw new Error("Base64 decoding is not available in this environment");
  }
  if (payload instanceof Uint8Array) return payload;
  // Some APIs return plain number arrays.
  if (Array.isArray(payload)) return new Uint8Array(payload);
  // Node Buffer (Uint8Array subclass) or ArrayBuffer.
  if (payload && typeof (payload as any).byteLength === "number") {
    return payload instanceof ArrayBuffer ? new Uint8Array(payload) : new Uint8Array(payload as any);
  }
  throw new Error("Unexpected binary payload returned from filesystem API");
}

function normalizeFileSize(payload: unknown): number {
  if (payload == null) {
    throw new Error("Unexpected stat payload returned from filesystem API");
  }
  if (typeof payload === "number") {
    if (!Number.isFinite(payload) || payload < 0) {
      throw new Error("Unexpected file size returned from filesystem API");
    }
    return payload;
  }
  if (typeof payload === "string") {
    const numeric = Number(payload);
    if (Number.isFinite(numeric) && numeric >= 0) {
      return numeric;
    }
  }
  if (payload && typeof payload === "object") {
    const obj = payload as any;
    const candidate =
      obj.sizeBytes ??
      obj.size_bytes ??
      obj.size ??
      obj.length ??
      obj.len ??
      obj.fileSize ??
      obj.file_size ??
      obj.bytes ??
      null;
    if (candidate != null) return normalizeFileSize(candidate);
  }
  throw new Error("Unexpected stat payload returned from filesystem API (missing file size)");
}

/**
 * Opens a native file dialog (Tauri) for selecting images.
 *
 * Returns absolute paths on disk.
 */
export async function pickImagesFromTauriDialog(options: PickImagesFromTauriDialogOptions = {}): Promise<string[]> {
  const open = getTauriDialogOpenOrNull();
  if (!open) {
    throw new Error("Tauri dialog.open API not available");
  }

  const multiple = options.multiple ?? true;
  const payload = await open({
    multiple,
    filters: [
      {
        name: "Images",
        extensions: [...IMAGE_FILE_EXTENSIONS],
      },
    ],
  });

  return normalizeOpenPaths(payload);
}

/**
 * Read an entire file into memory using Formula's hardened Tauri filesystem commands.
 *
 * Uses:
 * - `stat_file` for size guardrails
 * - `read_binary_file` for small payloads
 * - `read_binary_file_range` for larger payloads
 */
export async function readBinaryFile(path: string): Promise<Uint8Array> {
  const invoke = getTauriInvokeOrThrow();
  const trimmed = String(path ?? "").trim();
  if (!trimmed) throw new Error("File path cannot be empty");

  const statPayload = await invoke("stat_file", { path: trimmed });
  const fileSize = normalizeFileSize(statPayload);
  if (fileSize <= 0) return new Uint8Array(0);
  if (fileSize > MAX_INSERT_IMAGE_BYTES) {
    throw new FileTooLargeError(fileSize, MAX_INSERT_IMAGE_BYTES);
  }

  // Keep single-call reads for small payloads, but avoid `read_binary_file` for larger
  // files to reduce base64 overhead and keep parity with Power Query's adapter.
  const chunkSize = 1024 * 1024; // 1MiB (must be <= backend MAX_READ_RANGE_BYTES)
  const smallFileThreshold = 4 * chunkSize;
  if (fileSize <= smallFileThreshold) {
    const payload = await invoke("read_binary_file", { path: trimmed });
    const bytes = normalizeBinaryPayload(payload);
    if (bytes.length > MAX_INSERT_IMAGE_BYTES) {
      throw new FileTooLargeError(bytes.length, MAX_INSERT_IMAGE_BYTES);
    }
    return bytes;
  }

  const chunks: Uint8Array[] = [];
  let offset = 0;

  while (offset < fileSize) {
    const nextLength = Math.min(chunkSize, fileSize - offset);
    const payload = await invoke("read_binary_file_range", { path: trimmed, offset, length: nextLength });
    const bytes = normalizeBinaryPayload(payload);
    if (bytes.length === 0) break;
    chunks.push(bytes);
    offset += bytes.length;
    if (bytes.length < nextLength) break;
  }

  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  if (totalLength > MAX_INSERT_IMAGE_BYTES) {
    throw new FileTooLargeError(totalLength, MAX_INSERT_IMAGE_BYTES);
  }

  const out = new Uint8Array(totalLength);
  let pos = 0;
  for (const chunk of chunks) {
    out.set(chunk, pos);
    pos += chunk.length;
  }
  return out;
}
