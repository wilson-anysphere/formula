import { t } from "../i18n/index.js";

export type ProviderCloseInfo = { code: number; reason: string };

export const RESERVED_ROOT_GUARD_CLOSE_CODE = 1008;
const RESERVED_ROOT_GUARD_REASON_FRAGMENT = "reserved root mutation";

export const RESERVED_ROOT_GUARD_UI_MESSAGE_KEY = "collab.reservedRootGuard.message";

export function reservedRootGuardUiMessage(): string {
  const out = t(RESERVED_ROOT_GUARD_UI_MESSAGE_KEY);
  // Fallback to a reasonably actionable message if i18n keys are missing.
  if (out === RESERVED_ROOT_GUARD_UI_MESSAGE_KEY) {
    return (
      "The sync server closed the collaboration connection because the reserved root guard is enabled on the sync server (SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED). " +
      "In-doc versioning/branching stores (YjsVersionStore/YjsBranchStore) won't work, so Version History and Branch Manager actions are disabled. " +
      "To use these features, disable SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED on the sync server or configure an out-of-doc store (ApiVersionStore/SQLite)."
    );
  }
  return out;
}

// Preserve the detected error per provider instance so panels can show the banner
// even if they are opened after the close event occurred (or are re-opened).
const providerReservedRootGuardDetected = new WeakSet<object>();
const providerReservedRootGuardMonitors = new WeakMap<object, { subscribers: Set<(detected: boolean) => void> }>();

function ensureReservedRootGuardMonitor(provider: any): {
  key: object;
  monitor: { subscribers: Set<(detected: boolean) => void> };
} | null {
  if (!provider || (typeof provider !== "object" && typeof provider !== "function")) return null;
  const key = provider as object;
  let monitor = providerReservedRootGuardMonitors.get(key);
  if (!monitor) {
    monitor = { subscribers: new Set() };
    providerReservedRootGuardMonitors.set(key, monitor);
    // Install a single long-lived listener per provider so we can surface the
    // error even when panels are not mounted.
    listenForProviderCloseEvents(provider, (info) => {
      if (!isReservedRootGuardDisconnect(info)) return;
      providerReservedRootGuardDetected.add(key);
      for (const cb of Array.from(monitor!.subscribers)) {
        try {
          cb(true);
        } catch {
          // ignore
        }
      }
    });
  }
  return { key, monitor };
}

function coerceReason(reason: unknown): string {
  if (typeof reason === "string") return reason;
  if (reason == null) return "";

  // Node/ws provides a Buffer (Uint8Array) for the reason.
  if (typeof Buffer !== "undefined" && Buffer.isBuffer(reason)) {
    try {
      return reason.toString("utf8");
    } catch {
      // fall through
    }
  }

  if (reason instanceof Uint8Array) {
    try {
      return new TextDecoder("utf-8").decode(reason);
    } catch {
      return String(reason);
    }
  }

  if (reason instanceof ArrayBuffer) {
    try {
      return new TextDecoder("utf-8").decode(new Uint8Array(reason));
    } catch {
      return String(reason);
    }
  }

  try {
    // Many objects (e.g. Buffer) implement a useful `toString()`.
    return String(reason);
  } catch {
    return "";
  }
}

function parseCloseArgs(args: unknown[]): ProviderCloseInfo | null {
  if (args.length === 0) return null;

  if (typeof args[0] === "number") {
    const code = args[0];
    const reason = args.length >= 2 ? coerceReason(args[1]) : "";
    if (!Number.isFinite(code)) return null;
    return { code, reason };
  }

  if (args.length === 2 && typeof args[1] === "number") {
    const code = args[1] as number;
    const reason = coerceReason(args[0]);
    if (!Number.isFinite(code)) return null;
    return { code, reason };
  }

  if (args.length === 1) {
    const arg = args[0] as any;
    if (!arg) return null;
    if (Array.isArray(arg)) return parseCloseArgs(arg);

    if (typeof arg === "object") {
      const codeRaw = (arg as any).code;
      const code = typeof codeRaw === "number" ? codeRaw : Number(codeRaw);
      const reason = coerceReason((arg as any).reason);
      if (!Number.isFinite(code)) return null;
      return { code, reason };
    }
  }

  return null;
}

function attachWsCloseListener(ws: any, onClose: (info: ProviderCloseInfo) => void): (() => void) | null {
  if (!ws || (typeof ws !== "object" && typeof ws !== "function")) return null;

  // Node/ws (EventEmitter-style).
  if (typeof ws.on === "function") {
    const handler = (code: number, reason: unknown) => {
      if (typeof code !== "number") return;
      onClose({ code, reason: coerceReason(reason) });
    };
    try {
      ws.on("close", handler);
    } catch {
      return null;
    }
    return () => {
      try {
        if (typeof ws.off === "function") ws.off("close", handler);
        else if (typeof ws.removeListener === "function") ws.removeListener("close", handler);
      } catch {
        // ignore
      }
    };
  }

  // Browser WebSocket.
  if (typeof ws.addEventListener === "function") {
    const handler = (ev: any) => {
      const code = typeof ev?.code === "number" ? ev.code : Number(ev?.code);
      if (!Number.isFinite(code)) return;
      onClose({ code, reason: coerceReason(ev?.reason) });
    };
    try {
      ws.addEventListener("close", handler);
    } catch {
      return null;
    }
    return () => {
      try {
        if (typeof ws.removeEventListener === "function") ws.removeEventListener("close", handler);
      } catch {
        // ignore
      }
    };
  }

  return null;
}

export function listenForProviderCloseEvents(provider: any | null, onClose: (info: ProviderCloseInfo) => void): () => void {
  if (!provider) return () => {};
  const disposers: Array<() => void> = [];

  let wsDisposer: (() => void) | null = null;
  let currentWs: any = null;

  const refreshWsListener = () => {
    try {
      const nextWs = (provider as any).ws;
      if (nextWs === currentWs) return;
      if (!nextWs) {
        wsDisposer?.();
        wsDisposer = null;
        currentWs = null;
        return;
      }
      wsDisposer?.();
      currentWs = nextWs;
      wsDisposer = attachWsCloseListener(nextWs, onClose);
    } catch {
      // ignore
    }
  };

  // Preferred: some providers (e.g. y-websocket) surface a `connection-close` event with close code/reason.
  if (typeof provider.on === "function") {
    const handler = (...args: unknown[]) => {
      const info = parseCloseArgs(args);
      if (!info) return;
      onClose(info);
    };
    try {
      provider.on("connection-close", handler);
      disposers.push(() => {
        try {
          if (typeof provider.off === "function") provider.off("connection-close", handler);
        } catch {
          // ignore
        }
      });
    } catch {
      // ignore
    }

    // Some providers create/replace `provider.ws` lazily (or on reconnect). Keep
    // refreshing our `ws` close listener when the provider emits status changes.
    const statusHandler = () => refreshWsListener();
    try {
      provider.on("status", statusHandler);
      disposers.push(() => {
        try {
          if (typeof provider.off === "function") provider.off("status", statusHandler);
        } catch {
          // ignore
        }
      });
    } catch {
      // ignore
    }
  }

  // Best-effort: also attach directly to the underlying websocket when exposed.
  refreshWsListener();
  disposers.push(() => {
    try {
      wsDisposer?.();
    } catch {
      // ignore
    }
    wsDisposer = null;
    currentWs = null;
  });

  return () => {
    for (const dispose of disposers) dispose();
  };
}

export function isReservedRootGuardDisconnect(info: ProviderCloseInfo): boolean {
  if (info.code !== RESERVED_ROOT_GUARD_CLOSE_CODE) return false;
  return info.reason.toLowerCase().includes(RESERVED_ROOT_GUARD_REASON_FRAGMENT);
}

export function subscribeToReservedRootGuardDisconnect(
  provider: any | null,
  cb: (detected: boolean) => void,
): () => void {
  const ensured = ensureReservedRootGuardMonitor(provider);
  if (!ensured) return () => {};

  const { key, monitor } = ensured;
  monitor.subscribers.add(cb);
  try {
    cb(providerReservedRootGuardDetected.has(key));
  } catch {
    // ignore
  }
  return () => {
    monitor.subscribers.delete(cb);
  };
}

export function clearReservedRootGuardError(provider: any | null): void {
  if (!provider || (typeof provider !== "object" && typeof provider !== "function")) return;
  const key = provider as object;
  providerReservedRootGuardDetected.delete(key);
  const monitor = providerReservedRootGuardMonitors.get(key);
  if (!monitor) return;
  for (const cb of Array.from(monitor.subscribers)) {
    try {
      cb(false);
    } catch {
      // ignore
    }
  }
}
