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

function hasTauri() {
  return Boolean(globalThis.__TAURI__);
}

/**
 * @param {unknown} val
 * @returns {Uint8Array | undefined}
 */
function coerceUint8Array(val) {
  if (val instanceof Uint8Array) return val;
  if (val instanceof ArrayBuffer) return new Uint8Array(val);
  if (ArrayBuffer.isView(val) && val.buffer instanceof ArrayBuffer) {
    return new Uint8Array(val.buffer, val.byteOffset, val.byteLength);
  }
  if (Array.isArray(val) && val.every((b) => typeof b === "number")) return new Uint8Array(val);

  // Some native bridges may return base64.
  if (typeof val === "string") {
    try {
      const base64 = val.startsWith("data:") ? val.slice(val.indexOf(",") + 1) : val;
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
 * Merge clipboard fields from `source` into `target`, only filling missing values.
 *
 * @param {any} target
 * @param {any} source
 */
function mergeClipboardContent(target, source) {
  if (!source || typeof source !== "object") return;

  if (typeof target.text !== "string" && typeof source.text === "string") target.text = source.text;
  if (typeof target.html !== "string" && typeof source.html === "string") target.html = source.html;
  if (typeof target.rtf !== "string" && typeof source.rtf === "string") target.rtf = source.rtf;

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
          : undefined;
    if (typeof pngBase64 === "string") target.pngBase64 = pngBase64;
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
            if (typeof r.html === "string") native.html = r.html;
            if (typeof r.text === "string") native.text = r.text;
            if (typeof r.rtf === "string") native.rtf = r.rtf;

            const pngBase64 =
              typeof r.pngBase64 === "string"
                ? r.pngBase64
                : typeof r.png_base64 === "string"
                  ? r.png_base64
                  : undefined;
            if (typeof pngBase64 === "string") native.pngBase64 = pngBase64;

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
      // 1) Prefer rich writes via the native clipboard command when available (Tauri IPC).
      let wrote = false;
        if (typeof tauriInvoke === "function") {
          try {
            await tauriInvoke("clipboard_write", {
            payload: {
              text: payload.text,
              html: payload.html,
              rtf: payload.rtf,
              pngBase64: payload.pngBase64,
            },
            });
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
      if (payload.html && typeof ClipboardItem !== "undefined" && clipboard?.write) {
        try {
          const item = new ClipboardItem({
            "text/plain": new Blob([payload.text], { type: "text/plain" }),
            "text/html": new Blob([payload.html], { type: "text/html" }),
          });
          await clipboard.write([item]);
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
            const rtfType = item.types.find((t) => t === "text/rtf");
            const imagePngType = item.types.find((t) => t === "image/png");

            const html =
              htmlType &&
              (await item.getType(htmlType).then((b) => b.text()).catch(() => undefined));
            const text =
              textType &&
              (await item.getType(textType).then((b) => b.text()).catch(() => undefined));
            const rtf =
              rtfType &&
              (await item.getType(rtfType).then((b) => b.text()).catch(() => undefined));
            const imagePng =
              imagePngType &&
              (await item
                .getType(imagePngType)
                .then((b) => b.arrayBuffer())
                .then((buf) => new Uint8Array(buf))
                .catch(() => undefined));

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

      // Prefer writing both formats when possible.
      if (payload.html && typeof ClipboardItem !== "undefined" && clipboard?.write) {
        try {
          const item = new ClipboardItem({
            "text/plain": new Blob([payload.text], { type: "text/plain" }),
            "text/html": new Blob([payload.html], { type: "text/html" }),
          });
          await clipboard.write([item]);
          return;
        } catch {
          // Fall back to plain text.
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
