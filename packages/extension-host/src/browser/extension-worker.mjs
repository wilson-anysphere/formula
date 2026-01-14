// Import the workspace source directly so the extension host can run in
// environments where workspace package links are incomplete (e.g. minimal agent
// sandboxes).
import * as formulaApi from "../../../extension-api/index.mjs";
import { lockDownTauriGlobals } from "./tauri-globals.mjs";

formulaApi.__setTransport({
  postMessage: (message) => postMessage(message)
});

const nativeFetch = typeof globalThis.fetch === "function" ? globalThis.fetch.bind(globalThis) : null;

let workerData = null;
let extensionModule = null;
let activated = false;
let activationPromise = null;
let sandboxGuardrailsApplied = false;
let importPreflightPromise = null;

let nextInternalRequestId = 1;
const internalPending = new Map();

function createInternalRequestId() {
  return `__internal__${nextInternalRequestId++}`;
}

function serializeError(payload) {
  const out = { message: "Unknown error" };

  try {
    if (payload && typeof payload === "object" && "message" in payload) {
      out.message = String(payload.message);
    } else {
      out.message = String(payload);
    }
  } catch {
    out.message = "Unknown error";
  }

  try {
    if (payload && typeof payload === "object" && "stack" in payload && payload.stack != null) {
      out.stack = String(payload.stack);
    }
  } catch {
    // ignore stack serialization failures
  }

  try {
    if (payload && typeof payload === "object") {
      if (typeof payload.name === "string" && payload.name.trim().length > 0) {
        out.name = String(payload.name);
      }
      if (Object.prototype.hasOwnProperty.call(payload, "code")) {
        const code = payload.code;
        const primitive =
          code == null || typeof code === "string" || typeof code === "number" || typeof code === "boolean";
        out.code = primitive ? code : String(code);
      }
    }
  } catch {
    // ignore metadata serialization failures
  }

  return out;
}

function deserializeError(payload) {
  const message =
    typeof payload === "string" ? payload : String(payload?.message ?? "Unknown error");
  const err = new Error(message);
  if (payload?.stack) err.stack = String(payload.stack);
  if (typeof payload?.name === "string" && payload.name.trim().length > 0) {
    err.name = String(payload.name);
  }
  if (Object.prototype.hasOwnProperty.call(payload ?? {}, "code")) {
    err.code = payload.code;
  }
  return err;
}

function normalizeSandboxOptions(options) {
  const value = options && typeof options === "object" ? options : {};
  return {
    strictImports: value.strictImports !== false,
    disableEval: value.disableEval !== false
  };
}

function internalRpcCall(namespace, method, args) {
  const id = createInternalRequestId();
  postMessage({
    type: "api_call",
    id,
    namespace,
    method,
    args
  });

  return new Promise((resolve, reject) => {
    internalPending.set(id, { resolve, reject });
  });
}

function handleInternalResponse(message) {
  if (!message || typeof message !== "object") return false;
  if (message.type !== "api_result" && message.type !== "api_error") return false;
  const pending = internalPending.get(message.id);
  if (!pending) return false;
  internalPending.delete(message.id);
  if (message.type === "api_result") {
    pending.resolve(message.result);
  } else {
    pending.reject(deserializeError(message.error));
  }
  return true;
}

function init(data) {
  workerData = {
    extensionId: String(data?.extensionId ?? ""),
    extensionPath: String(data?.extensionPath ?? ""),
    extensionUri: String(data?.extensionUri ?? data?.extensionPath ?? ""),
    globalStoragePath: String(data?.globalStoragePath ?? ""),
    workspaceStoragePath: String(data?.workspaceStoragePath ?? ""),
    mainUrl: String(data?.mainUrl ?? ""),
    sandbox: normalizeSandboxOptions(data?.sandbox)
  };

  formulaApi.__setContext({
    extensionId: workerData.extensionId,
    extensionPath: workerData.extensionPath,
    extensionUri: workerData.extensionUri,
    globalStoragePath: workerData.globalStoragePath,
    workspaceStoragePath: workerData.workspaceStoragePath
  });

  applySandboxGuardrails(workerData.sandbox);
}

function findPrototypeWithOwnProperty(obj, prop) {
  let current = Object.getPrototypeOf(obj);
  while (current) {
    if (Object.prototype.hasOwnProperty.call(current, prop)) return current;
    current = Object.getPrototypeOf(current);
  }
  return null;
}

function defineReadOnlyProperty(target, prop, value) {
  try {
    Object.defineProperty(target, prop, {
      value,
      writable: false,
      configurable: false,
      enumerable: true
    });
    return true;
  } catch {
    try {
      // Fall back to assignment when defineProperty is not allowed.
      target[prop] = value;
      return true;
    } catch {
      return false;
    }
  }
}

function lockDownGlobal(prop, value) {
  defineReadOnlyProperty(globalThis, prop, value);

  const protoOwner = findPrototypeWithOwnProperty(globalThis, prop);
  if (protoOwner) {
    defineReadOnlyProperty(protoOwner, prop, value);
  }
}

function lockDownProperty(target, prop, value) {
  if (!target || (typeof target !== "object" && typeof target !== "function")) return;
  defineReadOnlyProperty(target, prop, value);

  const protoOwner = findPrototypeWithOwnProperty(target, prop);
  if (protoOwner) {
    defineReadOnlyProperty(protoOwner, prop, value);
  }
}

function applySandboxGuardrails(sandbox) {
  if (sandboxGuardrailsApplied) return;
  sandboxGuardrailsApplied = true;

  // Tauri injects `__TAURI__` / `__TAURI_IPC__` / `__TAURI_INVOKE__` globals in some contexts. If those are exposed
  // inside the extension worker, untrusted extension code could call native commands directly
  // and bypass Formula's permission model (clipboard, filesystem, etc). Best-effort lock them
  // down to `undefined` before loading any extension modules.
  lockDownTauriGlobals(lockDownGlobal);

  if (sandbox?.disableEval) {
    const blocked = (name) => {
      return function blockedCodegen() {
        throw new Error(`${name} is not allowed in extensions`);
      };
    };

    if (typeof globalThis.eval === "function") {
      lockDownGlobal("eval", blocked("eval"));
    }

    if (typeof globalThis.Function === "function") {
      const NativeFunction = globalThis.Function;
      const DisabledFunction = blocked("Function");
      lockDownGlobal("Function", DisabledFunction);
      lockDownProperty(NativeFunction.prototype, "constructor", DisabledFunction);

      try {
        // AsyncFunction is not exposed as a global; harden the prototype chain so
        // `Object.getPrototypeOf(async function(){}).constructor` cannot be used
        // to create new functions from strings.
        const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
        if (typeof AsyncFunction === "function") {
          lockDownProperty(AsyncFunction.prototype, "constructor", blocked("AsyncFunction"));
        }
      } catch {
        // ignore
      }

      try {
        const GeneratorFunction = Object.getPrototypeOf(function* () {}).constructor;
        if (typeof GeneratorFunction === "function") {
          lockDownProperty(GeneratorFunction.prototype, "constructor", blocked("GeneratorFunction"));
        }
      } catch {
        // ignore
      }

      try {
        const AsyncGeneratorFunction = Object.getPrototypeOf(async function* () {}).constructor;
        if (typeof AsyncGeneratorFunction === "function") {
          lockDownProperty(
            AsyncGeneratorFunction.prototype,
            "constructor",
            blocked("AsyncGeneratorFunction")
          );
        }
      } catch {
        // ignore
      }
    }

    if (typeof globalThis.setTimeout === "function") {
      const native = globalThis.setTimeout;
      lockDownGlobal("setTimeout", (handler, timeout, ...args) => {
        if (typeof handler === "string") {
          throw new Error("setTimeout with a string callback is not allowed in extensions");
        }
        return native(handler, timeout, ...args);
      });
    }

    if (typeof globalThis.setInterval === "function") {
      const native = globalThis.setInterval;
      lockDownGlobal("setInterval", (handler, timeout, ...args) => {
        if (typeof handler === "string") {
          throw new Error("setInterval with a string callback is not allowed in extensions");
        }
        return native(handler, timeout, ...args);
      });
    }
  }
}

if (typeof globalThis.fetch === "function") {
  const permissionedFetch = async (input, init) => {
    return formulaApi.network.fetch(String(input), init);
  };
  lockDownGlobal("fetch", permissionedFetch);
}

if (typeof globalThis.WebSocket === "function") {
  const NativeWebSocket = globalThis.WebSocket;
  const nativeInstances = new WeakMap();

  class PermissionedWebSocket {
    static CONNECTING = 0;
    static OPEN = 1;
    static CLOSING = 2;
    static CLOSED = 3;

    CONNECTING = 0;
    OPEN = 1;
    CLOSING = 2;
    CLOSED = 3;

    constructor(url, protocols) {
      this._url = String(url ?? "");
      this._protocols = protocols;
      this._readyState = PermissionedWebSocket.CONNECTING;
      this._binaryType = "blob";
      this._protocol = "";
      this._extensions = "";
      this._bufferedAmount = 0;
      this._pendingClose = null;
      /** @type {Map<string, Set<Function>>} */
      this._listeners = new Map();

      this.onopen = null;
      this.onmessage = null;
      this.onerror = null;
      this.onclose = null;

      void this._start();
    }

    get url() {
      const ws = nativeInstances.get(this);
      return ws ? ws.url : this._url;
    }

    get readyState() {
      const ws = nativeInstances.get(this);
      return ws ? ws.readyState : this._readyState;
    }

    get bufferedAmount() {
      const ws = nativeInstances.get(this);
      return ws ? ws.bufferedAmount : this._bufferedAmount;
    }

    get extensions() {
      const ws = nativeInstances.get(this);
      return ws ? ws.extensions : this._extensions;
    }

    get protocol() {
      const ws = nativeInstances.get(this);
      return ws ? ws.protocol : this._protocol;
    }

    get binaryType() {
      const ws = nativeInstances.get(this);
      return ws ? ws.binaryType : this._binaryType;
    }

    set binaryType(value) {
      this._binaryType = value;
      const ws = nativeInstances.get(this);
      if (ws) ws.binaryType = value;
    }

    addEventListener(type, listener) {
      if (typeof listener !== "function") return;
      const key = String(type);
      let set = this._listeners.get(key);
      if (!set) {
        set = new Set();
        this._listeners.set(key, set);
      }
      set.add(listener);
    }

    removeEventListener(type, listener) {
      const set = this._listeners.get(String(type));
      if (!set) return;
      set.delete(listener);
      if (set.size === 0) this._listeners.delete(String(type));
    }

    dispatchEvent(event) {
      if (!event || typeof event.type !== "string") return true;
      this._emit(event.type, event);
      return true;
    }

    send(data) {
      const ws = nativeInstances.get(this);
      if (this.readyState !== PermissionedWebSocket.OPEN || !ws) {
        throw new Error("WebSocket is not open");
      }

      ws.send(data);
    }

    close(code, reason) {
      const ws = nativeInstances.get(this);
      if (!ws) {
        this._pendingClose = { code, reason };
        this._readyState = PermissionedWebSocket.CLOSING;
        return;
      }
      ws.close(code, reason);
    }

    async _start() {
      try {
        await internalRpcCall("network", "openWebSocket", [this._url]);
      } catch (err) {
        this._fail(err);
        return;
      }

      let ws;
      try {
        ws =
          this._protocols === undefined ? new NativeWebSocket(this._url) : new NativeWebSocket(this._url, this._protocols);
      } catch (err) {
        this._fail(err);
        return;
      }

      nativeInstances.set(this, ws);
      try {
        ws.binaryType = this._binaryType;
      } catch {
        // ignore
      }

      ws.addEventListener("open", () => {
        this._readyState = PermissionedWebSocket.OPEN;
        this._protocol = ws.protocol;
        this._extensions = ws.extensions;
        this._emit("open", { type: "open", target: this });

        const pendingClose = this._pendingClose;
        if (pendingClose) {
          this._pendingClose = null;
          try {
            ws.close(pendingClose.code, pendingClose.reason);
          } catch {
            // ignore
          }
        }
      });

      ws.addEventListener("message", (event) => {
        this._emit("message", { type: "message", data: event.data, target: this });
      });

      ws.addEventListener("error", () => {
        this._emit("error", { type: "error", target: this });
      });

      ws.addEventListener("close", (event) => {
        this._readyState = PermissionedWebSocket.CLOSED;
        nativeInstances.delete(this);
        this._emit("close", {
          type: "close",
          code: event.code,
          reason: event.reason,
          wasClean: event.wasClean,
          target: this
        });
      });
    }

    _fail(err) {
      this._readyState = PermissionedWebSocket.CLOSED;
      this._emit("error", { type: "error", error: err, target: this });
      this._emit("close", { type: "close", code: 1008, reason: String(err?.message ?? err), wasClean: false, target: this });
    }

    _emit(type, event) {
      const evt = event && typeof event === "object" ? event : { type };
      const propHandler = this[`on${type}`];
      if (typeof propHandler === "function") {
        try {
          propHandler.call(this, evt);
        } catch {
          // ignore
        }
      }

      const set = this._listeners.get(String(type));
      if (!set) return;
      for (const listener of [...set]) {
        try {
          listener.call(this, evt);
        } catch {
          // ignore
        }
      }
    }
  }

  lockDownGlobal("WebSocket", PermissionedWebSocket);
}

// Prevent bypassing network permission checks via XHR. Extensions should use `fetch()`
// (which is replaced with a permission-gated wrapper above).
if (typeof globalThis.XMLHttpRequest === "function") {
  class PermissionedXMLHttpRequest {
    constructor() {
      throw new Error("XMLHttpRequest is not allowed in extensions; use fetch()");
    }
  }

  lockDownGlobal("XMLHttpRequest", PermissionedXMLHttpRequest);
}

if (typeof globalThis.EventSource === "function") {
  class PermissionedEventSource {
    constructor() {
      throw new Error("EventSource is not allowed in extensions");
    }
  }

  lockDownGlobal("EventSource", PermissionedEventSource);
}

if (typeof globalThis.WebTransport === "function") {
  class PermissionedWebTransport {
    constructor() {
      throw new Error("WebTransport is not allowed in extensions");
    }
  }

  lockDownGlobal("WebTransport", PermissionedWebTransport);
}

if (typeof globalThis.RTCPeerConnection === "function") {
  class PermissionedRTCPeerConnection {
    constructor() {
      throw new Error("RTCPeerConnection is not allowed in extensions");
    }
  }

  lockDownGlobal("RTCPeerConnection", PermissionedRTCPeerConnection);
}

if (globalThis.navigator && typeof globalThis.navigator.sendBeacon === "function") {
  lockDownProperty(globalThis.navigator, "sendBeacon", () => {
    throw new Error("navigator.sendBeacon is not allowed in extensions");
  });
}

// Prevent bypassing the permission-gated network APIs by spawning nested workers
// with pristine globals (native fetch/WebSocket/XHR).
if (typeof globalThis.Worker === "function") {
  class PermissionedWorker {
    constructor() {
      throw new Error("Worker is not allowed in extensions");
    }
  }
  lockDownGlobal("Worker", PermissionedWorker);
}

if (typeof globalThis.SharedWorker === "function") {
  class PermissionedSharedWorker {
    constructor() {
      throw new Error("SharedWorker is not allowed in extensions");
    }
  }
  lockDownGlobal("SharedWorker", PermissionedSharedWorker);
}

if (typeof globalThis.importScripts === "function") {
  lockDownGlobal("importScripts", () => {
    throw new Error("importScripts is not allowed in extensions");
  });
}

const IMPORT_PREFLIGHT_LIMITS = {
  maxModules: 200,
  maxTotalBytes: 5 * 1024 * 1024,
  maxModuleBytes: 256 * 1024
};

function scanModuleImports(source, url) {
  // Best-effort module graph validation. This intentionally ignores strings/comments so that
  // JSDoc `import("...")` type references don't trigger dynamic import checks.
  const src = String(source ?? "");
  const len = src.length;
  /** @type {string[]} */
  const specifiers = [];

  const isIdentifierStart = (ch) => /[A-Za-z_$]/.test(ch);
  const isIdentifierChar = (ch) => /[A-Za-z0-9_$]/.test(ch);
  const isDigit = (ch) => /[0-9]/.test(ch);
  const isWhitespace = (ch) => /\s/.test(ch);

  function skipWhitespaceAndComments(idx) {
    let i = idx;
    while (i < len) {
      const ch = src[i];
      if (isWhitespace(ch)) {
        i += 1;
        continue;
      }
      if (ch === "/" && src[i + 1] === "/") {
        i += 2;
        while (i < len && src[i] !== "\n") i += 1;
        continue;
      }
      if (ch === "/" && src[i + 1] === "*") {
        i += 2;
        while (i < len && !(src[i] === "*" && src[i + 1] === "/")) i += 1;
        if (i < len) i += 2;
        continue;
      }
      break;
    }
    return i;
  }

  function parseStringLiteral(idx) {
    const quote = src[idx];
    let i = idx + 1;
    let out = "";
    while (i < len) {
      const ch = src[i];
      if (ch === "\\") {
        i += 1;
        if (i >= len) break;
        out += src[i];
        i += 1;
        continue;
      }
      if (ch === quote) {
        return { value: out, end: i + 1 };
      }
      out += ch;
      i += 1;
    }
    return null;
  }

  function skipString(idx, quote) {
    let i = idx;
    while (i < len) {
      const ch = src[i];
      if (ch === "\\") {
        i += 2;
        continue;
      }
      if (ch === quote) {
        return i + 1;
      }
      i += 1;
    }
    return i;
  }

  function skipRegex(idx) {
    let i = idx;
    let inCharClass = false;
    while (i < len) {
      const ch = src[i];
      if (ch === "\\") {
        i += 2;
        continue;
      }
      if (ch === "[" && !inCharClass) {
        inCharClass = true;
        i += 1;
        continue;
      }
      if (ch === "]" && inCharClass) {
        inCharClass = false;
        i += 1;
        continue;
      }
      if (ch === "/" && !inCharClass) {
        i += 1;
        while (i < len && /[A-Za-z]/.test(src[i])) i += 1;
        return i;
      }
      i += 1;
    }
    return i;
  }

  function tryParseImportSpecifier(idx) {
    let i = skipWhitespaceAndComments(idx);
    if (src[i] === "'" || src[i] === '"') {
      const parsed = parseStringLiteral(i);
      if (parsed) return parsed;
      return null;
    }

    let braceDepth = 0;
    while (i < len) {
      const ch = src[i];
      if (isWhitespace(ch)) {
        i += 1;
        continue;
      }
      if (ch === "/" && src[i + 1] === "/") {
        i += 2;
        while (i < len && src[i] !== "\n") i += 1;
        continue;
      }
      if (ch === "/" && src[i + 1] === "*") {
        i += 2;
        while (i < len && !(src[i] === "*" && src[i + 1] === "/")) i += 1;
        if (i < len) i += 2;
        continue;
      }
      if (ch === "'" || ch === '"') {
        i = skipString(i + 1, ch);
        continue;
      }
      if (ch === "{") {
        braceDepth += 1;
        i += 1;
        continue;
      }
      if (ch === "}") {
        if (braceDepth > 0) braceDepth -= 1;
        i += 1;
        continue;
      }
      if (braceDepth === 0 && isIdentifierStart(ch)) {
        let j = i + 1;
        while (j < len && isIdentifierChar(src[j])) j += 1;
        const ident = src.slice(i, j);
        if (ident === "from") {
          const afterFrom = skipWhitespaceAndComments(j);
          if (src[afterFrom] === "'" || src[afterFrom] === '"') {
            return parseStringLiteral(afterFrom);
          }
        }
        i = j;
        continue;
      }
      if (ch === ";" || ch === "\n") break;
      i += 1;
    }
    return null;
  }

  function tryParseExportSpecifier(idx) {
    return tryParseImportSpecifier(idx);
  }

  let i = 0;
  let state = "code"; // code | template | lineComment | blockComment
  let regexAllowed = true;
  let afterPropertyDot = false;
  const templateBraceStack = [];

  while (i < len) {
    const ch = src[i];

    if (state === "code") {
      if (ch === "{") {
        if (templateBraceStack.length > 0) {
          templateBraceStack[templateBraceStack.length - 1] += 1;
        }
        regexAllowed = true;
        afterPropertyDot = false;
        i += 1;
        continue;
      }

      if (ch === "}" && templateBraceStack.length > 0) {
        const depth = templateBraceStack[templateBraceStack.length - 1];
        if (depth === 0) {
          templateBraceStack.pop();
          state = "template";
          i += 1;
          continue;
        }
        templateBraceStack[templateBraceStack.length - 1] -= 1;
        regexAllowed = false;
        afterPropertyDot = false;
        i += 1;
        continue;
      }

      if (isWhitespace(ch)) {
        i += 1;
        continue;
      }

      if (ch === "'" || ch === '"') {
        i = skipString(i + 1, ch);
        regexAllowed = false;
        afterPropertyDot = false;
        continue;
      }

      if (ch === "`") {
        state = "template";
        regexAllowed = false;
        afterPropertyDot = false;
        i += 1;
        continue;
      }

      if (ch === "/" && src[i + 1] === "/") {
        state = "lineComment";
        i += 2;
        continue;
      }

      if (ch === "/" && src[i + 1] === "*") {
        state = "blockComment";
        i += 2;
        continue;
      }

      if (ch === ".") {
        if (src[i + 1] === "." && src[i + 2] === ".") {
          afterPropertyDot = false;
          regexAllowed = true;
          i += 3;
          continue;
        }

        afterPropertyDot = true;
        regexAllowed = true;
        i += 1;
        continue;
      }

      if (ch === "/") {
        if (regexAllowed) {
          i = skipRegex(i + 1);
          regexAllowed = false;
          afterPropertyDot = false;
          continue;
        }
        regexAllowed = true;
        afterPropertyDot = false;
        i += 1;
        continue;
      }

      if (ch === "(" || ch === "[" || ch === "," || ch === ";" || ch === ":" || ch === "?" || ch === "=") {
        regexAllowed = true;
        afterPropertyDot = false;
        i += 1;
        continue;
      }

      if (ch === ")" || ch === "]") {
        regexAllowed = false;
        afterPropertyDot = false;
        i += 1;
        continue;
      }

      if ((ch === "+" || ch === "-") && src[i + 1] === ch) {
        regexAllowed = Boolean(regexAllowed);
        afterPropertyDot = false;
        i += 2;
        continue;
      }

      if (
        ch === "!" ||
        ch === "~" ||
        ch === "&" ||
        ch === "|" ||
        ch === "*" ||
        ch === "%" ||
        ch === "^" ||
        ch === "<" ||
        ch === ">" ||
        ch === "+" ||
        ch === "-"
      ) {
        regexAllowed = true;
        afterPropertyDot = false;
        i += 1;
        continue;
      }

      if (isIdentifierStart(ch)) {
        let j = i + 1;
        while (j < len && isIdentifierChar(src[j])) j += 1;
        const ident = src.slice(i, j);

        if (!afterPropertyDot && ident === "import") {
          const afterImport = skipWhitespaceAndComments(j);
          if (src[afterImport] === "(") {
            const argStart = skipWhitespaceAndComments(afterImport + 1);
            let detail = "";
            if (src[argStart] === "'" || src[argStart] === '"') {
              const parsed = parseStringLiteral(argStart);
              if (parsed) detail = ` (attempted to import '${parsed.value}')`;
            }
            throw new Error(`Dynamic import is not allowed in extensions${detail}${url ? ` (in ${url})` : ""}`);
          }

          if (src[afterImport] !== ".") {
            const parsed = tryParseImportSpecifier(afterImport);
            if (parsed?.value) {
              specifiers.push(parsed.value);
              if (parsed.end) {
                i = parsed.end;
                regexAllowed = false;
                afterPropertyDot = false;
                continue;
              }
            }
          }
        }

        if (!afterPropertyDot && ident === "export") {
          const parsed = tryParseExportSpecifier(j);
          if (parsed?.value) {
            specifiers.push(parsed.value);
            if (parsed.end) {
              i = parsed.end;
              regexAllowed = false;
              afterPropertyDot = false;
              continue;
            }
          }
        }

        regexAllowed = false;
        afterPropertyDot = false;
        i = j;
        continue;
      }

      if (isDigit(ch)) {
        let j = i + 1;
        while (j < len && /[0-9._xobA-Fa-f]/.test(src[j])) j += 1;
        regexAllowed = false;
        afterPropertyDot = false;
        i = j;
        continue;
      }

      afterPropertyDot = false;
      i += 1;
      continue;
    }

    if (state === "template") {
      if (ch === "\\") {
        i += 2;
        continue;
      }
      if (ch === "`") {
        state = "code";
        regexAllowed = false;
        afterPropertyDot = false;
        i += 1;
        continue;
      }
      if (ch === "$" && src[i + 1] === "{") {
        templateBraceStack.push(0);
        state = "code";
        regexAllowed = true;
        afterPropertyDot = false;
        i += 2;
        continue;
      }
      i += 1;
      continue;
    }

    if (state === "lineComment") {
      if (ch === "\n") {
        state = "code";
        regexAllowed = true;
      }
      i += 1;
      continue;
    }

    if (state === "blockComment") {
      if (ch === "*" && src[i + 1] === "/") {
        state = "code";
        regexAllowed = true;
        i += 2;
        continue;
      }
      i += 1;
      continue;
    }

    i += 1;
  }

  return specifiers;
}

function assertAllowedStaticImport(specifier, parentUrl) {
  const request = String(specifier ?? "");
  // Extensions are allowed to import the Formula extension API. In Vite dev the alias for
  // `@formula/extension-api` is rewritten to an absolute `/@fs/.../packages/extension-api/index.mjs`
  // specifier, so accept that form as well.
  const requestNoQuery = request.split("?", 1)[0] ?? request;
  if (request === "@formula/extension-api" || request === "formula") return { type: "virtual" };
  if (requestNoQuery.startsWith("/@fs/") && requestNoQuery.endsWith("/packages/extension-api/index.mjs")) {
    return { type: "virtual" };
  }

  // In-memory extension loaders (eg: web marketplace installs) may rewrite module specifiers to
  // `data:`/`blob:` URLs that contain already-verified code. Allow these, but only when they are
  // imported from an in-memory module as well (prevents remote/network-loaded extensions from
  // smuggling additional code via URL imports).
  if (request.startsWith("data:") || request.startsWith("blob:")) {
    const parent = String(parentUrl ?? "");
    if (parent.startsWith("data:") || parent.startsWith("blob:")) {
      return { type: "inline" };
    }
    throw new Error(
      `Disallowed import specifier '${request}' in ${parentUrl}: URL/protocol imports are not allowed; data/blob URL imports are only allowed from in-memory modules`
    );
  }

  if (request.startsWith("./") || request.startsWith("../")) return { type: "relative" };

  if (/^[A-Za-z][A-Za-z0-9+.-]*:/.test(request) || request.startsWith("//")) {
    throw new Error(
      `Disallowed import specifier '${request}' in ${parentUrl}: URL/protocol imports are not allowed; bundle dependencies instead`
    );
  }
  if (request.startsWith("/")) {
    throw new Error(
      `Disallowed import specifier '${request}' in ${parentUrl}: absolute imports are not allowed; use a relative path instead`
    );
  }

  throw new Error(
    `Disallowed import specifier '${request}' in ${parentUrl}: only relative imports ('./' or '../') and '@formula/extension-api' (or 'formula') are allowed`
  );
}

async function fetchModuleSource(url, rootUrl) {
  if (!nativeFetch) {
    throw new Error("Strict import validation requires fetch support in this runtime");
  }
  const response = await nativeFetch(url);
  const responseUrl = typeof response?.url === "string" ? response.url.trim() : "";
  const effectiveUrl =
    responseUrl.length > 0 ? responseUrl : String(url);
  if (
    rootUrl &&
    effectiveUrl &&
    !effectiveUrl.startsWith(rootUrl) &&
    (response?.redirected === true || effectiveUrl !== String(url))
  ) {
    throw new Error(
      `Extension module redirected outside the extension base URL: '${url}' resolved to '${effectiveUrl}'`
    );
  }
  if (!response.ok) {
    throw new Error(`Failed to fetch extension module: ${url} (${response.status})`);
  }
  const buffer = await response.arrayBuffer();
  const size = buffer.byteLength;
  return { source: new TextDecoder().decode(buffer), size, url: effectiveUrl };
}

async function validateModuleGraph(entryUrl, extensionRootUrl, limits = IMPORT_PREFLIGHT_LIMITS) {
  const entry = String(entryUrl ?? "");

  // `data:` and `blob:` module URLs are not hierarchical, so we cannot meaningfully apply the
  // extensionRootUrl prefix check (and `new URL("./", entry)` would throw). These entrypoints are
  // expected to be fully self-contained (no relative imports) and to only import other in-memory
  // modules.
  const enforceRoot = Boolean(extensionRootUrl) && !(entry.startsWith("data:") || entry.startsWith("blob:"));
  const root = enforceRoot ? new URL("./", extensionRootUrl).href : null;
  if (root && entry && !entry.startsWith(root)) {
    throw new Error(
      `Extension entrypoint must resolve inside the extension base URL: '${entry}' is outside '${root}'`
    );
  }
  /** @type {string[]} */
  const queue = [entry];
  const visited = new Set();
  const requested = new Set();
  let fetchCount = 0;
  let totalBytes = 0;

  while (queue.length > 0) {
    const requestedUrl = queue.shift();
    if (!requestedUrl) continue;
    if (requested.has(requestedUrl)) continue;
    requested.add(requestedUrl);

    if (fetchCount >= limits.maxModules) {
      throw new Error(
        `Extension module graph exceeded limit of ${limits.maxModules} modules (starting from ${entryUrl})`
      );
    }

    const { source, size, url } = await fetchModuleSource(requestedUrl, root);
    fetchCount += 1;

    if (size > limits.maxModuleBytes) {
      throw new Error(
        `Extension module too large: ${url} (${size} bytes; max ${limits.maxModuleBytes} bytes)`
      );
    }
    totalBytes += size;
    if (totalBytes > limits.maxTotalBytes) {
      throw new Error(
        `Extension module graph too large: exceeded ${limits.maxTotalBytes} bytes (starting from ${entryUrl})`
      );
    }

    if (visited.has(url)) continue;

    visited.add(url);

    const imports = scanModuleImports(source, url);
    for (const specifier of imports) {
      const kind = assertAllowedStaticImport(specifier, url);
      if (kind.type === "inline") {
        queue.push(String(specifier));
        continue;
      }
      if (kind.type !== "relative") continue;

      let resolved = "";
      try {
        resolved = new URL(specifier, url).href;
      } catch (error) {
        throw new Error(
          `Failed to resolve import specifier '${specifier}' in ${url}: ${String(error?.message ?? error)}`
        );
      }
      if (root && !resolved.startsWith(root)) {
        throw new Error(
          `Disallowed import specifier '${specifier}' in ${url}: resolved outside the extension base URL (${resolved})`
        );
      }
      queue.push(resolved);
    }
  }
}

async function activateExtension() {
  if (activated) return;
  if (activationPromise) return activationPromise;
  if (!workerData) throw new Error("Extension worker not initialized");

  activationPromise = (async () => {
    if (!extensionModule) {
      if (workerData.sandbox?.strictImports) {
        try {
          importPreflightPromise ??= validateModuleGraph(workerData.mainUrl, workerData.extensionPath);
          await importPreflightPromise;
        } catch (error) {
          // Allow retries on transient preflight failures by clearing the cached promise.
          importPreflightPromise = null;
          throw error;
        }
      }
      // The entrypoint URL is provided at runtime by the host, so Vite/Rollup
      // cannot statically analyze it. Suppress Vite's dynamic import warning.
      extensionModule = await import(/* @vite-ignore */ workerData.mainUrl);
    }

    const activateFn = extensionModule.activate || extensionModule.default?.activate;
    if (typeof activateFn !== "function") {
      throw new Error(`Extension entrypoint does not export an activate() function`);
    }

    const context = {
      extensionId: workerData.extensionId,
      extensionPath: workerData.extensionPath,
      extensionUri: workerData.extensionUri,
      globalStoragePath: workerData.globalStoragePath,
      workspaceStoragePath: workerData.workspaceStoragePath,
      subscriptions: []
    };

    await activateFn(context);
    activated = true;
  })();

  try {
    await activationPromise;
  } finally {
    if (!activated) activationPromise = null;
  }
}

self.addEventListener("message", (event) => {
  void (async () => {
    const message = event.data;
    if (!message || typeof message !== "object") return;

    if (message.type === "init") {
      init(message);
      return;
    }

    if (message.type === "activate") {
      try {
        await activateExtension();
        postMessage({ type: "activate_result", id: message.id });
      } catch (error) {
        postMessage({
          type: "activate_error",
          id: message.id,
          error: serializeError(error)
        });
      }
      return;
    }

    if (handleInternalResponse(message)) {
      return;
    }

    await Promise.resolve(formulaApi.__handleMessage(message));
  })().catch((error) => {
    // Message handlers ignore returned promises; swallow errors here so failures
    // don't become unhandled rejections in the worker.
    try {
      console.error("Unhandled extension worker message error:", error);
    } catch {
      // ignore logging failures
    }
  });
});
