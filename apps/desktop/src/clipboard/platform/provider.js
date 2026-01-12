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
const MAX_IMAGE_BYTES = 10 * 1024 * 1024; // 10MB
const MAX_RICH_TEXT_BYTES = 2 * 1024 * 1024; // 2MB (HTML / RTF)
const MAX_RICH_TEXT_CHARS = 2 * 1024 * 1024; // writing: approximate guard for JS strings

function hasTauri() {
  return Boolean(globalThis.__TAURI__);
}

/**
 * Rough estimate of bytes represented by a base64 string.
 *
 * @param {string} base64
 * @returns {number}
 */
function estimateBase64Bytes(base64) {
  const trimmed = base64.startsWith("data:") ? base64.slice(base64.indexOf(",") + 1) : base64;
  const normalized = trimmed.trim();
  if (!normalized) return 0;

  const padding = normalized.endsWith("==") ? 2 : normalized.endsWith("=") ? 1 : 0;
  return Math.floor((normalized.length * 3) / 4) - padding;
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

  if (Array.isArray(val) && val.every((b) => typeof b === "number")) {
    if (val.length > MAX_IMAGE_BYTES) return undefined;
    return new Uint8Array(val);
  }

  // Some native bridges may return base64.
  if (typeof val === "string") {
    const base64 = val.startsWith("data:") ? val.slice(val.indexOf(",") + 1) : val;
    if (estimateBase64Bytes(base64) > MAX_IMAGE_BYTES) return undefined;

    try {
      const bin = atob(base64);
      const out = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
      return out;
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
 * Merge clipboard fields from `source` into `target`, only filling missing values.
 *
 * @param {any} target
 * @param {any} source
 */
function mergeClipboardContent(target, source) {
  if (!source || typeof source !== "object") return;

  if (typeof target.text !== "string" && typeof source.text === "string") target.text = source.text;

  if (typeof target.html !== "string" && typeof source.html === "string" && source.html.length <= MAX_RICH_TEXT_CHARS) {
    target.html = source.html;
  }

  if (typeof target.rtf !== "string" && typeof source.rtf === "string" && source.rtf.length <= MAX_RICH_TEXT_CHARS) {
    target.rtf = source.rtf;
  }

  if (!(target.imagePng instanceof Uint8Array)) {
    const image =
      coerceUint8Array(source.imagePng) ??
      // Support snake_case bridges.
      coerceUint8Array(source.image_png);
    if (image) target.imagePng = image;
  }

  if (typeof target.pngBase64 !== "string") {
    const pngBase64 =
      typeof source.pngBase64 === "string"
        ? source.pngBase64
        : typeof source.png_base64 === "string"
          ? source.png_base64
          : typeof source.image_png_base64 === "string"
            ? source.image_png_base64
            : undefined;

    if (typeof pngBase64 === "string" && estimateBase64Bytes(pngBase64) <= MAX_IMAGE_BYTES) {
      target.pngBase64 = pngBase64;
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
    if (blob && typeof blob.size === "number" && blob.size > maxBytes) return undefined;
    return await blob.text();
  } catch {
    return undefined;
  }
}

/**
 * @param {any} item
 * @param {string} type
 * @param {number} maxBytes
 * @returns {Promise<Uint8Array | undefined>}
 */
async function readClipboardItemPng(item, type, maxBytes) {
  try {
    const blob = await item.getType(type);
    if (blob && typeof blob.size === "number" && blob.size > maxBytes) return undefined;
    const buf = await blob.arrayBuffer();
    return new Uint8Array(buf);
  } catch {
    return undefined;
  }
}

/**
 * @returns {Promise<ClipboardProvider>}
 */
export async function createClipboardProvider() {
  if (hasTauri()) return createTauriClipboardProvider();
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
      /** @type {any | undefined} */
      let native;

      // 1) Prefer rich reads via the native clipboard command when available (Tauri IPC).
      if (typeof tauriInvoke === "function") {
        try {
          const result = await tauriInvoke("clipboard_read");
          if (result && typeof result === "object") {
            /** @type {any} */
            const r = result;
            native = {};
            if (typeof r.text === "string") native.text = r.text;
            if (typeof r.html === "string" && r.html.length <= MAX_RICH_TEXT_CHARS) native.html = r.html;
            if (typeof r.rtf === "string" && r.rtf.length <= MAX_RICH_TEXT_CHARS) native.rtf = r.rtf;

            const pngBase64 =
              typeof r.pngBase64 === "string"
                ? r.pngBase64
                : typeof r.png_base64 === "string"
                  ? r.png_base64
                  : typeof r.image_png_base64 === "string"
                    ? r.image_png_base64
                    : undefined;
            if (typeof pngBase64 === "string" && estimateBase64Bytes(pngBase64) <= MAX_IMAGE_BYTES) {
              native.pngBase64 = pngBase64;
            }

            if (typeof native.html === "string" || typeof native.text === "string") {
              return native;
            }

            // Keep `native` around to merge into web reads below (e.g. rtf-only/image-only reads).
            if (Object.keys(native).length === 0) native = undefined;
          }
        } catch {
          // Ignore; command may not exist on older builds.
        }
      }

      // 2) Fall back to rich reads via the WebView Clipboard API when available so we can
      // ingest HTML tables + formats from external spreadsheets.
      const web = await createWebClipboardProvider().read();

      /** @type {any} */
      const merged = { ...web };
      if (merged.imagePng != null) merged.imagePng = coerceUint8Array(merged.imagePng);

      if (native) {
        mergeClipboardContent(merged, native);
      }

      // 3) Fill missing plain text via the legacy Tauri plain text clipboard API.
      if (typeof merged.text !== "string" && tauriClipboard?.readText) {
        try {
          const text = await tauriClipboard.readText();
          if (typeof text === "string") merged.text = text;
        } catch {
          // Ignore.
        }
      }

      return merged;
    },

    async write(payload) {
      const html = typeof payload.html === "string" && payload.html.length <= MAX_RICH_TEXT_CHARS ? payload.html : undefined;
      const rtf = typeof payload.rtf === "string" && payload.rtf.length <= MAX_RICH_TEXT_CHARS ? payload.rtf : undefined;
      const pngBase64 =
        typeof payload.pngBase64 === "string" && estimateBase64Bytes(payload.pngBase64) <= MAX_IMAGE_BYTES
          ? payload.pngBase64
          : undefined;

      // 1) Prefer rich writes via the native clipboard command when available (Tauri IPC).
      let wrote = false;
      if (typeof tauriInvoke === "function") {
        try {
          /** @type {any} */
          const richPayload = { text: payload.text };
          if (html) richPayload.html = html;
          if (rtf) richPayload.rtf = rtf;
          if (pngBase64) richPayload.pngBase64 = pngBase64;

          await tauriInvoke("clipboard_write", { payload: richPayload });
          wrote = true;
        } catch {
          wrote = false;
        }
      }

      if (!wrote) {
        if (tauriClipboard?.writeText) {
          try {
            await tauriClipboard.writeText(payload.text);
          } catch {
            await createWebClipboardProvider().write({ text: payload.text });
          }
        } else {
          await createWebClipboardProvider().write({ text: payload.text });
        }
      }

      // 2) Secondary path: Best-effort HTML write via ClipboardItem when available (WebView-dependent).
      const clipboard = globalThis.navigator?.clipboard;
      if (html && typeof ClipboardItem !== "undefined" && clipboard?.write) {
        try {
          const itemPayload = {
            "text/plain": new Blob([payload.text], { type: "text/plain" }),
            "text/html": new Blob([html], { type: "text/html" }),
          };

          await clipboard.write([new ClipboardItem(itemPayload)]);
        } catch {
          // Ignore; some platforms deny HTML clipboard writes.
        }
      }
    },
  };
}

/**
 * @returns {ClipboardProvider}
 */
function createWebClipboardProvider() {
  return {
    async read() {
      const clipboard = globalThis.navigator?.clipboard;

      // Prefer rich read if available.
      if (clipboard?.read) {
        try {
          const items = await clipboard.read();
          for (const item of items) {
            const htmlType = item.types.find((t) => t === "text/html");
            const textType = item.types.find((t) => t === "text/plain");
            const rtfType = item.types.find((t) => t === "text/rtf" || t === "application/rtf");
            const imagePngType = item.types.find((t) => t === "image/png");

            const html = htmlType ? await readClipboardItemText(item, htmlType, MAX_RICH_TEXT_BYTES) : undefined;
            const text = textType ? await readClipboardItemText(item, textType, Number.POSITIVE_INFINITY) : undefined;
            const rtf = rtfType ? await readClipboardItemText(item, rtfType, MAX_RICH_TEXT_BYTES) : undefined;
            const imagePng = imagePngType ? await readClipboardItemPng(item, imagePngType, MAX_IMAGE_BYTES) : undefined;

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
              return out;
            }
          }
        } catch {
          // Permission denied or unsupported – fall back to plain text below.
        }
      }

      let text;
      try {
        text = clipboard?.readText ? await clipboard.readText() : undefined;
      } catch {
        text = undefined;
      }
      return { text };
    },

    async write(payload) {
      const clipboard = globalThis.navigator?.clipboard;

      const html = typeof payload.html === "string" && payload.html.length <= MAX_RICH_TEXT_CHARS ? payload.html : undefined;
      const rtf = typeof payload.rtf === "string" && payload.rtf.length <= MAX_RICH_TEXT_CHARS ? payload.rtf : undefined;
      // @ts-ignore - payload may optionally include image bytes.
      const imagePngBlob = normalizeImagePngBlob(payload.imagePng);

      // Prefer writing rich formats when possible.
      if (typeof ClipboardItem !== "undefined" && clipboard?.write) {
        /** @type {Record<string, Blob>} */
        const itemPayload = {
          "text/plain": new Blob([payload.text], { type: "text/plain" }),
        };
        if (html) itemPayload["text/html"] = new Blob([html], { type: "text/html" });
        if (rtf) itemPayload["text/rtf"] = new Blob([rtf], { type: "text/rtf" });
        if (imagePngBlob) itemPayload["image/png"] = imagePngBlob;

        const hasRichFormats = Object.keys(itemPayload).length > 1;

        if (hasRichFormats) {
          try {
            await clipboard.write([new ClipboardItem(itemPayload)]);
            return;
          } catch {
            // Some platforms reject unknown/unsupported types (e.g. image/png, text/rtf).
            // Retry with the formats we used before introducing optional rich types so
            // we don't regress HTML clipboard writes.
            if (html && (rtf || imagePngBlob)) {
              try {
                const fallback = new ClipboardItem({
                  "text/plain": new Blob([payload.text], { type: "text/plain" }),
                  "text/html": new Blob([html], { type: "text/html" }),
                });
                await clipboard.write([fallback]);
                return;
              } catch {
                // Fall back to plain text below.
              }
            }
            // Fall back to plain text below.
          }
        }
      }

      if (clipboard?.writeText) {
        try {
          await clipboard.writeText(payload.text);
        } catch {
          // Ignore; clipboard write requires user gesture/permissions in browsers.
        }
      }
    },
  };
}
