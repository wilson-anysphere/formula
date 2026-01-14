/**
 * Platform clipboard integration.
 *
 * Desktop (Tauri): prefers the Tauri clipboard API when available.
 * Web: uses the browser Clipboard API (permission-gated).
 *
 * Note: Clipboard access generally requires a user gesture in browsers. Callers
 * should invoke these APIs from explicit copy/paste handlers.
 *
 * @typedef {import("../types.js").ClipboardContent} ClipboardContent
 * @typedef {import("../types.js").ClipboardWritePayload} ClipboardWritePayload
 *
 * @typedef {{
 *   read: () => Promise<ClipboardContent>,
 *   write: (payload: ClipboardWritePayload) => Promise<void>
 * }} ClipboardProvider
 */

// NOTE: Clipboard items can contain extremely large rich payloads (especially images).
// Guard against unbounded memory usage by skipping oversized formats.
//
// Keep this in sync with the Rust backend clipboard guards (`MAX_PNG_BYTES`).
const MAX_IMAGE_BYTES = 5 * 1024 * 1024; // 5MB (raw PNG bytes)
const MAX_RICH_TEXT_BYTES = 2 * 1024 * 1024; // 2MB (HTML / RTF / text)

// Export the effective numeric clipboard caps so tests (and other JS helpers) can assert
// they stay in sync with the Rust backend (`apps/desktop/src-tauri/src/clipboard/mod.rs`).
export const CLIPBOARD_LIMITS = {
  maxImageBytes: MAX_IMAGE_BYTES,
  maxRichTextBytes: MAX_RICH_TEXT_BYTES,
};

// Internal marker used to communicate "we detected an oversized text/plain blob" between provider
// layers (e.g. Web Clipboard API read -> Tauri provider fallback logic) without surfacing it to
// callers.
const SKIPPED_OVERSIZED_PLAINTEXT = Symbol("skippedOversizedPlainText");
// Non-enumerable marker for callers that explicitly handle image paste/insert UX.
// (Non-enumerable so existing deep-equality tests and `Object.keys` consumers are unaffected.)
const SKIPPED_OVERSIZED_IMAGE_PNG = "skippedOversizedImagePng";

function hasTauri() {
  return Boolean(globalThis.__TAURI__);
}

/**
 * Parse a debug flag from common representations.
 *
 * Accepts:
 * - boolean
 * - numbers (0/1)
 * - strings ("0"/"1"/"false"/"true")
 *
 * @param {unknown} value
 * @returns {boolean}
 */
function parseDebugFlag(value) {
  if (value === true) return true;
  if (value === false || value == null) return false;
  if (typeof value === "number") return value !== 0;
  if (typeof value === "string") {
    const v = value.trim().toLowerCase();
    return !(v === "" || v === "0" || v === "false");
  }
  return false;
}

/**
 * Best-effort check for the clipboard debug flag.
 *
 * Notes:
 * - In production Vite builds, `import.meta.env` values are build-time constants.
 * - For field diagnostics without rebuilding, callers can set
 *   `globalThis.FORMULA_DEBUG_CLIPBOARD = true` via devtools.
 *
 * @returns {boolean}
 */
function isClipboardDebugEnabled() {
  // Runtime override (devtools / injected global).
  if (parseDebugFlag(globalThis.FORMULA_DEBUG_CLIPBOARD ?? globalThis.__FORMULA_DEBUG_CLIPBOARD__)) {
    return true;
  }

  // Build-time Vite env.
  try {
    const env = /** @type {any} */ (import.meta).env;
    if (!env || typeof env !== "object") return false;
    return parseDebugFlag(env.VITE_FORMULA_DEBUG_CLIPBOARD ?? env.FORMULA_DEBUG_CLIPBOARD);
  } catch {
    return false;
  }
}

/**
 * @param {boolean} enabled
 * @param  {...any} args
 */
function clipboardDebug(enabled, ...args) {
  if (!enabled) return;
  try {
    const fn = console?.debug ?? console?.log;
    if (typeof fn === "function") fn.call(console, "[clipboard]", ...args);
  } catch {
    // ignore
  }
}

const isTrimChar = (code) => code === 0x20 || code === 0x09 || code === 0x0a || code === 0x0d; // space, tab, lf, cr

/**
 * Returns `true` if `text` is within `maxBytes` when encoded as UTF-8.
 *
 * This is a guardrail used to avoid allocating large buffers or sending huge IPC payloads.
 * It intentionally avoids `TextEncoder` so we don't materialize the full encoded byte array.
 *
 * @param {string} text
 * @param {number} maxBytes
 * @returns {boolean}
 */
function utf8WithinLimit(text, maxBytes) {
  if (text.length > maxBytes) return false;

  let bytes = 0;
  for (let i = 0; i < text.length; i += 1) {
    const code = text.charCodeAt(i);
    if (code < 0x80) {
      bytes += 1;
    } else if (code < 0x800) {
      bytes += 2;
    } else if (code >= 0xd800 && code <= 0xdbff) {
      // High surrogate (potential start of a surrogate pair).
      const next = text.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        // Valid surrogate pair -> 4 bytes.
        bytes += 4;
        i += 1;
      } else {
        // Unpaired surrogate -> U+FFFD replacement (3 bytes).
        bytes += 3;
      }
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      // Unpaired low surrogate -> U+FFFD replacement (3 bytes).
      bytes += 3;
    } else {
      bytes += 3;
    }

    if (bytes > maxBytes) return false;
  }

  return true;
}

/**
 * @param {unknown} value
 * @param {number} maxBytes
 * @returns {value is string}
 */
function isStringWithinUtf8Limit(value, maxBytes) {
  return typeof value === "string" && utf8WithinLimit(value, maxBytes);
}

function hasDataUrlPrefixAt(str, start) {
  if (start + 5 > str.length) return false;
  // ASCII case-insensitive match for "data:" without allocating.
  return (
    (str.charCodeAt(start) | 32) === 0x64 && // d
    (str.charCodeAt(start + 1) | 32) === 0x61 && // a
    (str.charCodeAt(start + 2) | 32) === 0x74 && // t
    (str.charCodeAt(start + 3) | 32) === 0x61 && // a
    str.charCodeAt(start + 4) === 0x3a // :
  );
}

/**
 * Compute the (start,end) bounds for the base64 payload inside an input string.
 *
 * - Trims leading/trailing whitespace without allocating.
 * - If a `data:*;base64,` prefix is present (case-insensitive, ignoring leading whitespace),
 *   skips everything up to the first comma.
 *
 * @param {string} base64
 * @returns {{ start: number; end: number }}
 */
function base64Bounds(base64) {
  let start = 0;
  while (start < base64.length && isTrimChar(base64.charCodeAt(start))) start += 1;

  if (hasDataUrlPrefixAt(base64, start)) {
    // Scan only a small prefix for the comma separator so malformed inputs like
    // `data:AAAAA...` don't force an O(n) search over huge payloads.
    let commaIndex = -1;
    const maxHeaderScan = Math.min(base64.length, start + 1024);
    for (let i = start; i < maxHeaderScan; i += 1) {
      if (base64.charCodeAt(i) === 0x2c) {
        // ','
        commaIndex = i;
        break;
      }
    }
    if (commaIndex >= 0) {
      start = commaIndex + 1;
    } else {
      // Malformed data URL (missing comma separator). Treat as empty so callers don't
      // accidentally decode `data:...` as base64.
      return { start: base64.length, end: base64.length };
    }
    while (start < base64.length && isTrimChar(base64.charCodeAt(start))) start += 1;
  }

  let end = base64.length;
  while (end > start && isTrimChar(base64.charCodeAt(end - 1))) end -= 1;
  return { start, end };
}

/**
 * Strip any `data:*;base64,` prefix and trim whitespace.
 *
 * @param {string} base64
 * @returns {string}
 */
function normalizeBase64String(base64) {
  if (typeof base64 !== "string") return "";
  const { start, end } = base64Bounds(base64);
  if (end <= start) return "";
  return base64.slice(start, end);
}

/**
 * Rough estimate of bytes represented by a base64 string.
 *
 * @param {string} base64
 * @returns {number}
 */
function estimateBase64Bytes(base64) {
  const { start, end } = base64Bounds(base64);
  const len = end - start;
  if (len <= 0) return 0;

  let padding = 0;
  if (base64.charCodeAt(end - 1) === 0x3d) {
    // '=' padding
    padding = 1;
    if (end - 2 >= start && base64.charCodeAt(end - 2) === 0x3d) padding = 2;
  }

  const bytes = Math.floor((len * 3) / 4) - padding;
  return bytes > 0 ? bytes : 0;
}

/**
 * @param {any} source
 * @returns {string | undefined}
 */
function readPngBase64(source) {
  if (!source || typeof source !== "object") return undefined;
  const raw =
    typeof source.pngBase64 === "string"
      ? source.pngBase64
      : typeof source.png_base64 === "string"
        ? source.png_base64
        : typeof source.image_png_base64 === "string"
          ? source.image_png_base64
          : undefined;

  // Return the raw wire format string here. Callers are responsible for applying
  // size checks before stripping `data:` prefixes / trimming to avoid allocating
  // huge intermediate strings for oversized payloads.
  return typeof raw === "string" ? raw : undefined;
}

/**
 * @param {unknown} val
 * @returns {Uint8Array | undefined}
 */
function coerceUint8Array(val) {
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

  if (Array.isArray(val)) {
    // Avoid iterating extremely large arrays; reject based on length first.
    if (val.length > MAX_IMAGE_BYTES) return undefined;
    if (val.every((b) => typeof b === "number")) return new Uint8Array(val);
  }

  // Some native bridges may return base64.
  if (typeof val === "string") {
    // Estimate size before decoding and (importantly) before slicing data URLs into a second large string.
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
      // Ignore.
    }
  }

  return undefined;
}

/**
 * @param {any} val
 * @returns {Blob | undefined}
 */
function normalizeImagePngBlob(val) {
  if (!val) return undefined;

  if (typeof Blob !== "undefined" && val instanceof Blob) {
    if (typeof val.size === "number" && val.size > MAX_IMAGE_BYTES) return undefined;
    if (val.type === "image/png") return val;
    return new Blob([val], { type: "image/png" });
  }

  const bytes = coerceUint8Array(val);
  if (!bytes) return undefined;
  return new Blob([bytes], { type: "image/png" });
}

/**
 * Normalize an image payload into raw bytes (for Tauri IPC).
 *
 * @param {any} val
 * @returns {Promise<Uint8Array | undefined>}
 */
async function normalizeImagePngBytes(val) {
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

  return coerceUint8Array(val);
}

/**
 * Encode bytes as base64 (no `data:` prefix).
 *
 * @param {Uint8Array} bytes
 * @returns {string}
 */
function encodeBase64(bytes) {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(bytes).toString("base64");
  }

  if (typeof btoa !== "function") {
    throw new Error("base64 encoding is unavailable in this environment");
  }

  // Avoid stack overflows by chunking `fromCharCode` calls.
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    // eslint-disable-next-line unicorn/prefer-code-point
    binary += String.fromCharCode(...chunk);
  }
  return btoa(binary);
}

/**
 * Merge clipboard fields from `source` into `target`, only filling missing values.
 *
 * @param {any} target
 * @param {any} source
 */
function mergeClipboardContent(target, source) {
  if (!source || typeof source !== "object") return;
  if (source[SKIPPED_OVERSIZED_IMAGE_PNG]) {
    Object.defineProperty(target, SKIPPED_OVERSIZED_IMAGE_PNG, { value: true });
  }

  if (
    typeof target.text !== "string" &&
    isStringWithinUtf8Limit(source.text, MAX_RICH_TEXT_BYTES)
  ) {
    target.text = source.text;
  }

  if (typeof target.html !== "string" && isStringWithinUtf8Limit(source.html, MAX_RICH_TEXT_BYTES)) {
    target.html = source.html;
  }

  if (typeof target.rtf !== "string" && isStringWithinUtf8Limit(source.rtf, MAX_RICH_TEXT_BYTES)) {
    target.rtf = source.rtf;
  }

  if (!(target.imagePng instanceof Uint8Array)) {
    const pngBase64 = readPngBase64(source);
    const oversizedBase64 = typeof pngBase64 === "string" && estimateBase64Bytes(pngBase64) > MAX_IMAGE_BYTES;
    const oversizedBytes = (() => {
      const candidate = source.imagePng ?? source.image_png;
      if (candidate instanceof Uint8Array) return candidate.byteLength > MAX_IMAGE_BYTES;
      if (candidate instanceof ArrayBuffer) return candidate.byteLength > MAX_IMAGE_BYTES;
      if (ArrayBuffer.isView(candidate) && candidate.buffer instanceof ArrayBuffer) {
        return candidate.byteLength > MAX_IMAGE_BYTES;
      }
      if (Array.isArray(candidate)) return candidate.length > MAX_IMAGE_BYTES;
      return false;
    })();

    const image =
      coerceUint8Array(source.imagePng) ??
      // Support snake_case bridges.
      coerceUint8Array(source.image_png) ??
      // Base64 (Tauri IPC / legacy bridges).
      coerceUint8Array(pngBase64);
    if (image) target.imagePng = image;
    if (image && typeof target.pngBase64 === "string") {
      // Maintain the invariant that callers see `imagePng` as the primary image format.
      // (Keep base64 only on decode failures / legacy callsites.)
      delete target.pngBase64;
    } else if (!image && typeof target.pngBase64 !== "string" && typeof pngBase64 === "string") {
      // Only preserve base64 when we couldn't decode it into bytes.
      if (estimateBase64Bytes(pngBase64) <= MAX_IMAGE_BYTES) {
        const normalized = normalizeBase64String(pngBase64);
        if (normalized) target.pngBase64 = normalized;
      } else if (oversizedBase64 || oversizedBytes) {
        Object.defineProperty(target, SKIPPED_OVERSIZED_IMAGE_PNG, { value: true });
      }
    } else if (!image && (oversizedBase64 || oversizedBytes)) {
      Object.defineProperty(target, SKIPPED_OVERSIZED_IMAGE_PNG, { value: true });
    }
  }
}

/**
 * @param {any} item
 * @param {string} type
 * @param {number} maxBytes
 * @returns {Promise<string | undefined>}
 */
async function readClipboardItemText(item, type, maxBytes) {
  try {
    const blob = await item.getType(type);
    if (!blob || typeof blob.size !== "number") return undefined;
    if (blob.size > maxBytes) return undefined;
    return await blob.text();
  } catch {
    return undefined;
  }
}

/**
 * @param {any} item
 * @param {string} type
 * @param {number} maxBytes
 * @returns {Promise<{ bytes?: Uint8Array, skippedOversized?: boolean }>}
 */
async function readClipboardItemPng(item, type, maxBytes) {
  try {
    const blob = await item.getType(type);
    if (!blob || typeof blob.size !== "number") return { bytes: undefined, skippedOversized: false };
    if (blob.size > maxBytes) return { bytes: undefined, skippedOversized: true };
    const buf = await blob.arrayBuffer();
    return { bytes: new Uint8Array(buf), skippedOversized: false };
  } catch {
    return { bytes: undefined, skippedOversized: false };
  }
}

/**
 * Best-effort rich clipboard read via `navigator.clipboard.read()` without falling back to
 * `navigator.clipboard.readText()`.
 *
 * This is used by the Tauri provider when native IPC already returned `text/html` and we only
 * want to merge in missing rich formats (e.g. `text/rtf`, `image/png`) without triggering
 * permission-gated plaintext reads.
 *
 * @param {{ wantRtf?: boolean, wantImagePng?: boolean }} wants
 * @returns {Promise<ClipboardContent | undefined>}
 */
async function readWebClipboardRichOnly(wants = {}) {
  const wantRtf = Boolean(wants && typeof wants === "object" && wants.wantRtf);
  const wantImagePng = Boolean(wants && typeof wants === "object" && wants.wantImagePng);
  if (!wantRtf && !wantImagePng) return undefined;

  const clipboard = globalThis.navigator?.clipboard;
  if (typeof clipboard?.read !== "function") return undefined;

  try {
    const items = await clipboard.read();
    /** @type {any} */
    const out = {};

    const matchMime = (value, exact) => {
      if (typeof value !== "string") return false;
      const normalized = value.trim().toLowerCase();
      return normalized === exact || normalized.startsWith(`${exact};`);
    };

    for (const item of items) {
      if (!item || typeof item !== "object" || !Array.isArray(item.types)) continue;

      if (wantRtf && typeof out.rtf !== "string") {
        const rtfType = item.types.find(
          (t) =>
            matchMime(t, "text/rtf") || matchMime(t, "application/rtf") || matchMime(t, "application/x-rtf")
        );
        if (rtfType) {
          const rtf = await readClipboardItemText(item, rtfType, MAX_RICH_TEXT_BYTES);
          if (typeof rtf === "string") out.rtf = rtf;
        }
      }

      if (wantImagePng && !(out.imagePng instanceof Uint8Array)) {
        const imagePngType = item.types.find((t) => matchMime(t, "image/png"));
        if (imagePngType) {
          const imagePng = await readClipboardItemPng(item, imagePngType, MAX_IMAGE_BYTES);
          if (imagePng instanceof Uint8Array) out.imagePng = imagePng;
        }
      }

      if ((!wantRtf || typeof out.rtf === "string") && (!wantImagePng || out.imagePng instanceof Uint8Array)) {
        break;
      }
    }

    if (typeof out.rtf === "string" || out.imagePng instanceof Uint8Array) return out;
    return undefined;
  } catch {
    return undefined;
  }
}

/**
 * @returns {Promise<ClipboardProvider>}
 */
export async function createClipboardProvider() {
  const debug = isClipboardDebugEnabled();
  if (hasTauri()) {
    clipboardDebug(debug, "provider=tauri");
    return createTauriClipboardProvider();
  }
  clipboardDebug(debug, "provider=web");
  return createWebClipboardProvider();
}

/**
 * @returns {ClipboardProvider}
 */
function createTauriClipboardProvider() {
  const tauriInvoke = globalThis.__TAURI__?.core?.invoke;
  const tauriClipboard = globalThis.__TAURI__?.clipboard;

  return {
    async read() {
      const debug = isClipboardDebugEnabled();
      /** @type {string[] | null} */
      const path = debug ? [] : null;
      /** @type {any | undefined} */
      let native;
      let clipboardReadErrored = false;
      let skippedOversizedPlainText = false;

      // 1) Prefer rich reads via the native clipboard command when available (Tauri IPC).
      if (typeof tauriInvoke === "function") {
        try {
          const result = await tauriInvoke("clipboard_read");
          if (result && typeof result === "object") {
            if (path) path.push("native-ipc:clipboard_read");
            /** @type {any} */
            const r = result;
            if (typeof r.text === "string" && !utf8WithinLimit(r.text, MAX_RICH_TEXT_BYTES)) {
              skippedOversizedPlainText = true;
            }
            native = {};
            if (isStringWithinUtf8Limit(r.text, MAX_RICH_TEXT_BYTES)) native.text = r.text;
            if (isStringWithinUtf8Limit(r.html, MAX_RICH_TEXT_BYTES)) native.html = r.html;
            if (isStringWithinUtf8Limit(r.rtf, MAX_RICH_TEXT_BYTES)) native.rtf = r.rtf;

            const pngBase64 = readPngBase64(r);
            if (pngBase64) {
              const imagePng = coerceUint8Array(pngBase64);
              if (imagePng) {
                native.imagePng = imagePng;
              } else if (estimateBase64Bytes(pngBase64) <= MAX_IMAGE_BYTES) {
                // Preserve base64 only when decoding fails (legacy/internal).
                const normalized = normalizeBase64String(pngBase64);
                if (normalized) native.pngBase64 = normalized;
              } else {
                Object.defineProperty(native, SKIPPED_OVERSIZED_IMAGE_PNG, { value: true });
              }
            }

            // If we successfully read HTML from the native clipboard, we can often return
            // immediately (it already includes rich spreadsheet formats on supported platforms).
            //
            // However: native HTML reads can miss other rich formats (e.g. RTF / image/png) that
            // may be present via `navigator.clipboard.read()`. Only attempt a rich merge when:
            // - native HTML is present
            // - we're missing at least one useful rich format
            // - `navigator.clipboard.read` exists (no `readText()` fallback / permission prompt)
            if (typeof native.html === "string") {
              const missingRtf = typeof native.rtf !== "string";
              const missingImagePng = !(native.imagePng instanceof Uint8Array);

              if ((missingRtf || missingImagePng) && typeof globalThis.navigator?.clipboard?.read === "function") {
                if (path) path.push("web-clipboard:clipboard.read(rich-only)");
                const webRich = await readWebClipboardRichOnly({
                  wantRtf: missingRtf,
                  wantImagePng: missingImagePng,
                });
                if (webRich) mergeClipboardContent(native, webRich);
              }

              if (path) clipboardDebug(true, "read path:", path.join(" -> "));
              return native;
            }

            // Keep `native` around to merge into web reads below (e.g. rtf-only/image-only reads).
            if (Object.keys(native).length === 0) native = undefined;
          }
        } catch {
          // Ignore; command may not exist on older builds.
          clipboardReadErrored = true;
        }
      }

      // 2) Fall back to rich reads via the WebView Clipboard API when available so we can
      // ingest HTML tables + formats from external spreadsheets.
      if (path) path.push("web-clipboard");
      const web = await createWebClipboardProvider({ suppressDebug: true }).read();
      // If the web clipboard path observed an oversized `text/plain` payload, do not fall back to
      // other plain-text clipboard APIs (they may allocate the same huge string).
      if (web && typeof web === "object" && web[SKIPPED_OVERSIZED_PLAINTEXT]) {
        skippedOversizedPlainText = true;
      }

      const skippedOversizedImagePng = Boolean(web && typeof web === "object" && web[SKIPPED_OVERSIZED_IMAGE_PNG]);

      /** @type {any} */
      const merged = { ...web };
      if (merged.imagePng != null) merged.imagePng = coerceUint8Array(merged.imagePng);
      if (skippedOversizedImagePng) {
        Object.defineProperty(merged, SKIPPED_OVERSIZED_IMAGE_PNG, { value: true });
      }
      // Ensure we don't leak internal sentinel properties to consumers.
      delete merged[SKIPPED_OVERSIZED_PLAINTEXT];

      if (native) {
        mergeClipboardContent(merged, native);
      }

      // Older desktop builds exposed rich clipboard reads via `read_clipboard` instead of
      // `clipboard_read`. If the newer command is missing, try the legacy name as a
      // best-effort merge (never clobbering WebView values).
      if (clipboardReadErrored && typeof tauriInvoke === "function") {
        try {
          if (path) path.push("native-ipc:read_clipboard");
          const legacy = await tauriInvoke("read_clipboard");
          if (legacy && typeof legacy === "object" && typeof legacy.text === "string") {
            if (!utf8WithinLimit(legacy.text, MAX_RICH_TEXT_BYTES)) {
              skippedOversizedPlainText = true;
            }
          }
          mergeClipboardContent(merged, legacy);
        } catch {
          // Ignore.
        }
      }

      // 3) Fill missing plain text via the legacy Tauri plain text clipboard API.
      if (!skippedOversizedPlainText && typeof merged.text !== "string" && tauriClipboard?.readText) {
        try {
          const text = await tauriClipboard.readText();
          if (isStringWithinUtf8Limit(text, MAX_RICH_TEXT_BYTES)) merged.text = text;
          if (typeof merged.text === "string" && path) path.push("legacy-plaintext:tauriClipboard.readText");
        } catch {
          // Ignore.
        }
      }

      if (path) clipboardDebug(true, "read path:", path.join(" -> "));
      return merged;
    },

    async write(payload) {
      const debug = isClipboardDebugEnabled();
      /** @type {string[] | null} */
      const path = debug ? [] : null;
      const hasText = typeof payload.text === "string";
      // Preserve whether callers actually provided `text`. Historically we coerced missing
      // `payload.text` to the empty string, which could unintentionally clobber the user's
      // clipboard when other formats were dropped by size guards (e.g. oversized images).
      const text = hasText ? payload.text : "";
      const textWithinLimit = !hasText || utf8WithinLimit(text, MAX_RICH_TEXT_BYTES);

      // Avoid sending extremely large plain-text payloads over the rich clipboard IPC path.
      // For oversized payloads, best-effort write plain text only.
      if (hasText && !textWithinLimit) {
        if (tauriClipboard?.writeText) {
          try {
            await tauriClipboard.writeText(text);
            if (path) {
              path.push("legacy-plaintext:tauriClipboard.writeText");
              clipboardDebug(true, "write path:", path.join(" -> "));
            }
            return;
          } catch {
            // Fall through to web clipboard fallback below.
          }
        }
        if (path) path.push("web-clipboard");
        await createWebClipboardProvider({ suppressDebug: true }).write({ text });
        if (path) clipboardDebug(true, "write path:", path.join(" -> "));
        return;
      }

      const html = isStringWithinUtf8Limit(payload.html, MAX_RICH_TEXT_BYTES) ? payload.html : undefined;
      const rtf = isStringWithinUtf8Limit(payload.rtf, MAX_RICH_TEXT_BYTES) ? payload.rtf : undefined;
      const imageBytes = await normalizeImagePngBytes(payload.imagePng);

      const pngBase64FromImage = imageBytes ? encodeBase64(imageBytes) : undefined;
      const legacyPngBase64 =
        typeof payload.pngBase64 === "string" && estimateBase64Bytes(payload.pngBase64) <= MAX_IMAGE_BYTES
          ? normalizeBase64String(payload.pngBase64)
          : undefined;

      const pngBase64 = pngBase64FromImage ?? legacyPngBase64;

      // If the caller didn't provide text and *all* other formats are absent (or were dropped by
      // size guards), do nothing rather than writing an empty `text/plain` payload.
      if (!hasText && !html && !rtf && !pngBase64) {
        return;
      }

      // 1) Prefer rich writes via the native clipboard command when available (Tauri IPC).
      let wrote = false;
      if (typeof tauriInvoke === "function") {
        try {
          /** @type {any} */
          const richPayload = {};
          if (hasText) richPayload.text = text;
          if (html) richPayload.html = html;
          if (rtf) richPayload.rtf = rtf;
          if (pngBase64) richPayload.pngBase64 = pngBase64;

          await tauriInvoke("clipboard_write", { payload: richPayload });
          if (path) path.push("native-ipc:clipboard_write");
          wrote = true;
        } catch {
          wrote = false;
        }
      }

      // 1b) Older desktop builds used a `write_clipboard` command for rich formats.
      if (!wrote && hasText && typeof tauriInvoke === "function") {
        try {
          await tauriInvoke("write_clipboard", {
            text,
            html,
            rtf,
            image_png_base64: pngBase64,
          });
          if (path) path.push("native-ipc:write_clipboard");
          wrote = true;
        } catch {
          wrote = false;
        }
      }

      if (!wrote) {
        // Only fall back to plain-text clipboard writes when the caller actually supplied text.
        // If text was omitted, preserve rich-only writes and avoid clobbering the clipboard with
        // an empty `text/plain` payload.
        if (hasText) {
          if (tauriClipboard?.writeText) {
            try {
              await tauriClipboard.writeText(text);
              if (path) path.push("legacy-plaintext:tauriClipboard.writeText");
            } catch {
              if (path) path.push("web-clipboard");
              await createWebClipboardProvider({ suppressDebug: true }).write({ text });
            }
          } else {
            if (path) path.push("web-clipboard");
            await createWebClipboardProvider({ suppressDebug: true }).write({ text });
          }
        }
      }

      // 2) Secondary path: Best-effort HTML write via ClipboardItem when available (WebView-dependent).
      // Only do this when we *didn't* successfully write rich formats via the native Tauri command.
      // The Web Clipboard API write can replace the entire clipboard item, dropping formats like
      // RTF/image that were written natively.
      //
      // ClipboardItem writes are best-effort and intentionally omit RTF so we don't regress HTML
      // clipboard writes on platforms that reject unsupported types.
      const clipboard = globalThis.navigator?.clipboard;
      if (!wrote && html && typeof ClipboardItem !== "undefined" && clipboard?.write) {
        try {
          /** @type {Record<string, Blob>} */
          const itemPayload = {
            "text/html": new Blob([html], { type: "text/html" }),
          };
          if (hasText) itemPayload["text/plain"] = new Blob([text], { type: "text/plain" });
          await clipboard.write([new ClipboardItem(itemPayload)]);
          if (path) path.push("web-clipboard:clipboard.write");
        } catch {
          // Ignore; some platforms deny rich clipboard writes.
        }
      }

      if (path) clipboardDebug(true, "write path:", path.join(" -> "));
    },
  };
}

/**
 * @returns {ClipboardProvider}
 */
function createWebClipboardProvider(options = {}) {
  const suppressDebug = options && typeof options === "object" ? Boolean(options.suppressDebug) : false;
  return {
    async read() {
      const debug = !suppressDebug && isClipboardDebugEnabled();
      const clipboard = globalThis.navigator?.clipboard;

      // If we skip an oversized plain-text payload from `clipboard.read()`, do not fall back to
      // `clipboard.readText()` (which may allocate the same oversized string anyway).
      let skipped_oversized_plain_text = false;
      let skipped_oversized_image_png = false;

      // Prefer rich read if available.
      if (clipboard?.read) {
        try {
          const items = await clipboard.read();
          clipboardDebug(debug, "read provider=web-clipboard.read()");
          for (const item of items) {
            const matchMime = (value, exact) => {
              if (typeof value !== "string") return false;
              const normalized = value.trim().toLowerCase();
              return normalized === exact || normalized.startsWith(`${exact};`);
            };

            const htmlType = item.types.find((t) => matchMime(t, "text/html"));
            const textType = item.types.find((t) => matchMime(t, "text/plain"));
            const rtfType = item.types.find(
              (t) => matchMime(t, "text/rtf") || matchMime(t, "application/rtf") || matchMime(t, "application/x-rtf")
            );
            const imagePngType = item.types.find((t) => matchMime(t, "image/png"));

            const html = htmlType ? await readClipboardItemText(item, htmlType, MAX_RICH_TEXT_BYTES) : undefined;
            /** @type {string | undefined} */
            let text;
            if (textType) {
              try {
                const blob = await item.getType(textType);
                if (blob && typeof blob.size === "number") {
                  if (blob.size > MAX_RICH_TEXT_BYTES) {
                    skipped_oversized_plain_text = true;
                  } else {
                    text = await blob.text();
                  }
                }
              } catch {
                text = undefined;
              }
            }
            const rtf = rtfType ? await readClipboardItemText(item, rtfType, MAX_RICH_TEXT_BYTES) : undefined;
            const imagePngResult = imagePngType ? await readClipboardItemPng(item, imagePngType, MAX_IMAGE_BYTES) : undefined;
            const imagePng = imagePngResult?.bytes;
            if (imagePngResult?.skippedOversized) skipped_oversized_image_png = true;

            if (
              typeof html === "string" ||
              typeof text === "string" ||
              typeof rtf === "string" ||
               imagePng instanceof Uint8Array
             ) {
              /** @type {any} */
              const out = {};
              if (typeof html === "string") out.html = html;
              if (typeof text === "string") out.text = text;
              if (typeof rtf === "string") out.rtf = rtf;
              if (imagePng instanceof Uint8Array) out.imagePng = imagePng;
              if (skipped_oversized_plain_text) {
                Object.defineProperty(out, SKIPPED_OVERSIZED_PLAINTEXT, { value: true });
              }
              if (skipped_oversized_image_png) {
                Object.defineProperty(out, SKIPPED_OVERSIZED_IMAGE_PNG, { value: true });
              }
              return out;
            }
          }
        } catch {
          // Permission denied or unsupported – fall back to plain text below.
        }
      }

      if (skipped_oversized_plain_text || skipped_oversized_image_png) {
        /** @type {any} */
        const out = {};
        // Non-enumerable so consumers' `Object.keys` / JSON serialization ignore it.
        if (skipped_oversized_plain_text) {
          Object.defineProperty(out, SKIPPED_OVERSIZED_PLAINTEXT, { value: true });
        }
        if (skipped_oversized_image_png) {
          Object.defineProperty(out, SKIPPED_OVERSIZED_IMAGE_PNG, { value: true });
        }
        return out;
      }

      let text;
      try {
        text = clipboard?.readText ? await clipboard.readText() : undefined;
        if (typeof text === "string") clipboardDebug(debug, "read provider=web-clipboard.readText()");
      } catch {
        text = undefined;
      }
      if (typeof text === "string" && !utf8WithinLimit(text, MAX_RICH_TEXT_BYTES)) {
        // Best-effort guardrail: `readText()` provides no size metadata, so we can only cap after
        // the fact. Still avoid returning huge strings downstream.
        skipped_oversized_plain_text = true;
        text = undefined;
      }

      if (skipped_oversized_plain_text) {
        /** @type {any} */
        const out = {};
        Object.defineProperty(out, SKIPPED_OVERSIZED_PLAINTEXT, { value: true });
        return out;
      }

      return { text };
    },

    async write(payload) {
      const debug = !suppressDebug && isClipboardDebugEnabled();
      const clipboard = globalThis.navigator?.clipboard;

      const hasText = typeof payload.text === "string";
      const text = hasText ? payload.text : "";
      const textWithinLimit = !hasText || utf8WithinLimit(text, MAX_RICH_TEXT_BYTES);

      // For very large text payloads, prefer `writeText()` and skip rich ClipboardItem writes to
      // avoid allocating large intermediate Blobs.
      if (hasText && !textWithinLimit) {
        if (clipboard?.writeText) {
          try {
            await clipboard.writeText(text);
            clipboardDebug(debug, "write provider=web-clipboard.writeText() (oversized text)");
          } catch {
            // Ignore; clipboard write requires user gesture/permissions in browsers.
          }
          return;
        }

        // Best-effort: `writeText` may be unavailable, but `clipboard.write` might still work.
        if (typeof ClipboardItem !== "undefined" && clipboard?.write) {
          try {
            await clipboard.write([
              new ClipboardItem({
                "text/plain": new Blob([text], { type: "text/plain" }),
              }),
            ]);
            clipboardDebug(debug, "write provider=web-clipboard.write() (plain text fallback)");
          } catch {
            // Ignore.
          }
        }
        return;
      }

      const html = isStringWithinUtf8Limit(payload.html, MAX_RICH_TEXT_BYTES) ? payload.html : undefined;
      const rtf = isStringWithinUtf8Limit(payload.rtf, MAX_RICH_TEXT_BYTES) ? payload.rtf : undefined;
      const imagePngBlob = normalizeImagePngBlob(payload.imagePng) ?? normalizeImagePngBlob(payload.pngBase64);

      // If the caller didn't provide text and *all* other formats are absent (or were dropped by
      // size guards), do nothing rather than overwriting the clipboard with empty text.
      if (!hasText && !html && !rtf && !imagePngBlob) {
        return;
      }

      // Prefer writing rich formats when possible.
      const hasNonPlainTextFormats = Boolean(html || rtf || imagePngBlob);
      if (hasNonPlainTextFormats && typeof ClipboardItem !== "undefined" && clipboard?.write) {
        /** @type {Record<string, Blob>} */
        const itemPayload = {};
        if (hasText) itemPayload["text/plain"] = new Blob([text], { type: "text/plain" });
        if (html) itemPayload["text/html"] = new Blob([html], { type: "text/html" });
        if (rtf) itemPayload["text/rtf"] = new Blob([rtf], { type: "text/rtf" });
      if (imagePngBlob) itemPayload["image/png"] = imagePngBlob;
        try {
          await clipboard.write([new ClipboardItem(itemPayload)]);
          clipboardDebug(debug, "write provider=web-clipboard.write() (ClipboardItem)");
          return;
        } catch {
          // Some platforms reject unknown/unsupported types (e.g. image/png, text/rtf).
          // Retry with the formats we used before introducing optional rich types so
          // we don't regress HTML clipboard writes.
          if (html && (rtf || imagePngBlob)) {
            try {
              /** @type {Record<string, Blob>} */
              const fallbackPayload = {
                "text/html": new Blob([html], { type: "text/html" }),
              };
              if (hasText) fallbackPayload["text/plain"] = new Blob([text], { type: "text/plain" });
              await clipboard.write([new ClipboardItem(fallbackPayload)]);
              clipboardDebug(debug, "write provider=web-clipboard.write() (html fallback)");
              return;
            } catch {
              // Fall back to plain text below.
            }
          }
          // Fall back to plain text below.
        }
      }

      if (hasText && clipboard?.writeText) {
        try {
          await clipboard.writeText(text);
          clipboardDebug(debug, "write provider=web-clipboard.writeText()");
        } catch {
          // Ignore; clipboard write requires user gesture/permissions in browsers.
        }
      }
    },
  };
}
