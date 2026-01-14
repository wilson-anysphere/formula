import { getTauriInvokeOrNull } from "./api";

const BLOCKED_PROTOCOLS = new Set(["javascript", "data", "file"]);

function getTauriGlobalOrNull(): any | null {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (globalThis as any).__TAURI__ ?? null;
  } catch {
    // Some hardened host environments (or tests) may define `__TAURI__` with a throwing getter.
    // Treat that as "unavailable" so best-effort callsites can fall back cleanly.
    return null;
  }
}

function safeGetProp(obj: any, prop: string): any | undefined {
  if (!obj) return undefined;
  try {
    return obj[prop];
  } catch {
    return undefined;
  }
}

function parseUrlOrThrow(url: string): URL {
  try {
    return new URL(url);
  } catch {
    throw new Error(`Invalid URL: ${url}`);
  }
}

type TauriShellOpen = (url: string, options?: Record<string, unknown>) => Promise<unknown> | unknown;

function getTauriShellOpenOrNull(): ((url: string) => Promise<void>) | null {
  const tauri = getTauriGlobalOrNull();
  const plugin = safeGetProp(tauri, "plugin");
  const plugins = safeGetProp(tauri, "plugins");
  const shell = safeGetProp(tauri, "shell");
  const pluginShell = safeGetProp(plugin, "shell");
  const pluginsShell = safeGetProp(plugins, "shell");
  const candidates: Array<{ owner: any; fn: unknown }> = [
    // Tauri v1 style
    { owner: shell, fn: safeGetProp(shell, "open") },
    // Tauri v2 plugin style
    { owner: pluginShell, fn: safeGetProp(pluginShell, "open") },
    // Alternate namespaces seen in some builds/tests.
    { owner: pluginsShell, fn: safeGetProp(pluginsShell, "open") },
  ];
  for (const { owner, fn } of candidates) {
    if (typeof fn !== "function") continue;
    const open = fn as TauriShellOpen;
    return async (url: string) => {
      await open.call(owner, url);
    };
  }
  return null;
}

/**
 * Open a URL in the host OS browser (Tauri) when available.
 *
 * - Desktop/Tauri: uses the Rust command `open_external_url` via the Tauri IPC bridge (`core.invoke`).
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
  if ((protocol === "http" || protocol === "https") && (parsed.username !== "" || parsed.password !== "")) {
    throw new Error("Refusing to open URL containing a username/password");
  }

  const tauri = getTauriGlobalOrNull();
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
    throw new Error("Tauri invoke API unavailable (expected core.invoke)");
  }

  throw new Error("No shellOpen implementation available in this environment");
}
