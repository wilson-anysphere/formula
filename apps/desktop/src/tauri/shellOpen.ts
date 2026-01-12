const BLOCKED_PROTOCOLS = new Set(["javascript", "data", "file"]);

function parseUrlOrThrow(url: string): URL {
  try {
    return new URL(url);
  } catch {
    throw new Error(`Invalid URL: ${url}`);
  }
}

type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;

function getTauriInvokeOrNull(): TauriInvoke | null {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  return typeof invoke === "function" ? invoke : null;
}

type TauriShellOpen = (url: string, options?: Record<string, unknown>) => Promise<unknown> | unknown;

function getTauriShellOpenOrNull(): ((url: string) => Promise<void>) | null {
  const tauri = (globalThis as any).__TAURI__;
  const candidates: unknown[] = [
    // Tauri v1 style
    tauri?.shell?.open,
    // Tauri v2 plugin style
    tauri?.plugin?.shell?.open,
    // Alternate namespaces seen in some builds/tests.
    tauri?.plugins?.shell?.open,
  ];
  for (const candidate of candidates) {
    if (typeof candidate === "function") {
      const fn = candidate as TauriShellOpen;
      return async (url: string) => {
        await fn(url);
      };
    }
  }
  return null;
}

/**
 * Open a URL in the host OS browser (Tauri) when available.
 *
 * - Desktop/Tauri: uses the Rust command `open_external_url` via `__TAURI__.core.invoke(...)`.
 * - Web builds: falls back to `window.open(..., "noopener,noreferrer")`.
 *
 * Security: blocks `javascript:`, `data:`, and `file:` URLs regardless of environment.
 */
export async function shellOpen(url: string): Promise<void> {
  const parsed = parseUrlOrThrow(url);
  const protocol = parsed.protocol.replace(":", "").toLowerCase();
  if (BLOCKED_PROTOCOLS.has(protocol)) {
    throw new Error(`Refusing to open URL with blocked protocol "${protocol}:"`);
  }

  const tauri = (globalThis as any).__TAURI__;
  const invoke = getTauriInvokeOrNull();

  if (invoke) {
    await invoke("open_external_url", { url });
    return;
  }

  const shellOpenFn = getTauriShellOpenOrNull();
  // Fallback for runtimes/tests that expose only the shell plugin API.
  if (shellOpenFn) {
    await shellOpenFn(url);
    return;
  }

  // In web builds (no Tauri runtime) fall back to a browser navigation.
  if (!tauri && typeof window !== "undefined" && typeof window.open === "function") {
    window.open(url, "_blank", "noopener,noreferrer");
    return;
  }

  if (tauri) {
    // If we're running under Tauri we should *never* silently fall back to `window.open`,
    // which would navigate inside the webview instead of the system browser.
    throw new Error("Tauri invoke API unavailable (expected __TAURI__.core.invoke)");
  }

  throw new Error("No shellOpen implementation available in this environment");
}
