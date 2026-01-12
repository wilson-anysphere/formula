const BLOCKED_PROTOCOLS = new Set(["javascript", "data"]);

function parseUrlOrThrow(url: string): URL {
  try {
    return new URL(url);
  } catch {
    throw new Error(`Invalid URL: ${url}`);
  }
}

/**
 * Open a URL in the host OS browser (Tauri) when available.
 *
 * - Desktop/Tauri: uses the shell plugin (`__TAURI__.plugin.shell.open`).
 * - Web builds: falls back to `window.open(..., "noopener,noreferrer")`.
 *
 * Security: blocks `javascript:` and `data:` URLs regardless of environment.
 */
export async function shellOpen(url: string): Promise<void> {
  const parsed = parseUrlOrThrow(url);
  const protocol = parsed.protocol.replace(":", "").toLowerCase();
  if (BLOCKED_PROTOCOLS.has(protocol)) {
    throw new Error(`Refusing to open URL with blocked protocol "${protocol}:"`);
  }

  const tauri = (globalThis as any).__TAURI__;
  const tauriOpen = tauri?.plugin?.shell?.open ?? tauri?.shell?.open;
  if (typeof tauriOpen === "function") {
    await tauriOpen(url);
    return;
  }

  if (typeof window !== "undefined" && typeof window.open === "function") {
    window.open(url, "_blank", "noopener,noreferrer");
    return;
  }

  throw new Error("No shellOpen implementation available in this environment");
}

