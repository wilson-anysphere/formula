const { parentPort, workerData } = require("node:worker_threads");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const Module = require("node:module");

const extensionRoot = fs.realpathSync(workerData.extensionPath);
const mainPath = path.resolve(workerData.mainPath);
const builtinModules = new Set(
  Module.builtinModules.map((m) => (m.startsWith("node:") ? m.slice("node:".length) : m))
);

const sandboxGlobal = Object.create(null);
const sandboxContext = vm.createContext(sandboxGlobal, {
  codeGeneration: { strings: false, wasm: false }
});

const SandboxError = vm.runInContext("Error", sandboxContext);
const SandboxObject = vm.runInContext("Object", sandboxContext);
const SandboxArray = vm.runInContext("Array", sandboxContext);

function createSandboxError(message) {
  return new SandboxError(String(message));
}

function serializeErrorForTransport(error) {
  const payload = { message: "Unknown error" };

  try {
    if (error && typeof error === "object" && "message" in error) {
      payload.message = String(error.message);
    } else {
      payload.message = String(error);
    }
  } catch {
    payload.message = "Unknown error";
  }

  try {
    if (error && typeof error === "object" && "stack" in error && error.stack != null) {
      payload.stack = String(error.stack);
    }
  } catch {
    // ignore
  }

  try {
    if (error && typeof error === "object") {
      if (typeof error.name === "string" && error.name.trim().length > 0) {
        payload.name = String(error.name);
      }
      if (Object.prototype.hasOwnProperty.call(error, "code")) {
        const code = error.code;
        const primitive =
          code == null || typeof code === "string" || typeof code === "number" || typeof code === "boolean";
        payload.code = primitive ? code : String(code);
      }
    }
  } catch {
    // ignore
  }

  return payload;
}

function hardenHostFunction(fn) {
  Object.setPrototypeOf(fn, null);
  return fn;
}

const hostPostMessage = hardenHostFunction((message) => {
  try {
    parentPort.postMessage(message);
  } catch (error) {
    throw createSandboxError(
      `Failed to send message from extension worker: ${String(error?.message ?? error)}`
    );
  }
});

const installSandboxGlobals = vm.runInContext(
  `
((postMessage) => {
  function safeSerializeLogArg(arg) {
    if (arg instanceof Error) {
      return { error: { message: arg.message, stack: arg.stack } };
    }
    if (typeof arg === "string") return arg;
    if (typeof arg === "number" || typeof arg === "boolean" || arg === null) return arg;
    try {
      return JSON.parse(JSON.stringify(arg));
    } catch {
      try {
        return String(arg);
      } catch {
        return "[Unserializable]";
      }
    }
  }

  const console = {};
  for (const level of ["log", "info", "warn", "error"]) {
    console[level] = (...args) => {
      try {
        postMessage({ type: "log", level, args: args.map(safeSerializeLogArg) });
      } catch {
        // ignore
      }
    };
  }

  const process = Object.freeze({
    binding(name) {
      throw new Error(
        \`process.binding() is not allowed in extensions (attempted '\${String(name)}')\`
      );
    }
  });

  // Prevent a known vm sandbox escape: extensions can install a custom
  // Error.prepareStackTrace that receives CallSite objects. When the stack
  // contains host frames, CallSite#getFunction() can return *host* functions,
  // whose constructors bypass codeGeneration: { strings: false }.
  try {
    Object.defineProperty(Error, "prepareStackTrace", {
      configurable: false,
      enumerable: false,
      get() {
        return undefined;
      },
      set() {
        throw new Error("Error.prepareStackTrace is not allowed in extensions");
      }
    });
  } catch {
    // ignore
  }

  return { console: Object.freeze(console), process };
})
`,
  sandboxContext
);

const { console: sandboxConsole, process: sandboxProcess } = installSandboxGlobals(hostPostMessage);
sandboxGlobal.console = sandboxConsole;
sandboxGlobal.process = sandboxProcess;

// Provide timer APIs without leaking Node.js Timeout objects into the VM realm.
// Extensions (and our own permission tests) expect browser-like numeric timer handles.
let nextTimeoutId = 1;
const timeouts = new Map(); // id -> NodeJS.Timeout
let nextIntervalId = 1;
const intervals = new Map(); // id -> NodeJS.Timeout
let nextImmediateId = 1;
const immediates = new Map(); // id -> NodeJS.Immediate

function normalizeDelay(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return 0;
  return Math.max(0, num);
}

const hostSetTimeout = hardenHostFunction((callback, delay, ...args) => {
  const id = nextTimeoutId++;
  const handle = setTimeout(() => {
    timeouts.delete(id);
    callback(...args);
  }, normalizeDelay(delay));
  timeouts.set(id, handle);
  return id;
});

const hostClearTimeout = hardenHostFunction((id) => {
  const handle = timeouts.get(id);
  if (!handle) return;
  timeouts.delete(id);
  clearTimeout(handle);
});

const hostSetImmediate = hardenHostFunction((callback, ...args) => {
  const id = nextImmediateId++;
  const handle = setImmediate(() => {
    immediates.delete(id);
    // Intentionally do not catch errors: unhandled exceptions should crash the worker
    // so the ExtensionHost can restart it.
    callback(...args);
  });
  immediates.set(id, handle);
  return id;
});

const hostClearImmediate = hardenHostFunction((id) => {
  const handle = immediates.get(id);
  if (!handle) return;
  immediates.delete(id);
  clearImmediate(handle);
});

const hostSetInterval = hardenHostFunction((callback, delay, ...args) => {
  const id = nextIntervalId++;
  const handle = setInterval(() => {
    callback(...args);
  }, normalizeDelay(delay));
  intervals.set(id, handle);
  return id;
});

const hostClearInterval = hardenHostFunction((id) => {
  const handle = intervals.get(id);
  if (!handle) return;
  intervals.delete(id);
  clearInterval(handle);
});

const installSandboxTimers = vm.runInContext(
  `
((hostSetTimeout, hostClearTimeout, hostSetInterval, hostClearInterval, hostSetImmediate, hostClearImmediate) => {
  return {
    // Never pass extension callbacks directly to host timers. If we do, the callback's
    // arguments.callee.caller chain can expose a host function, whose
    // constructor.constructor(...) bypasses codeGeneration: { strings: false }.
    setTimeout: (cb, delay, ...args) => {
      const wrapped = (...innerArgs) => cb(...innerArgs);
      return hostSetTimeout(wrapped, delay, ...args);
    },
    clearTimeout: (id) => hostClearTimeout(id),
    setInterval: (cb, delay, ...args) => {
      const wrapped = (...innerArgs) => cb(...innerArgs);
      return hostSetInterval(wrapped, delay, ...args);
    },
    clearInterval: (id) => hostClearInterval(id),
    setImmediate: (cb, ...args) => {
      const wrapped = (...innerArgs) => cb(...innerArgs);
      return hostSetImmediate(wrapped, ...args);
    },
    clearImmediate: (id) => hostClearImmediate(id)
  };
})
`,
  sandboxContext
);

const sandboxTimers = installSandboxTimers(
  hostSetTimeout,
  hostClearTimeout,
  hostSetInterval,
  hostClearInterval,
  hostSetImmediate,
  hostClearImmediate
);
sandboxGlobal.setTimeout = sandboxTimers.setTimeout;
sandboxGlobal.clearTimeout = sandboxTimers.clearTimeout;
sandboxGlobal.setImmediate = sandboxTimers.setImmediate;
sandboxGlobal.clearImmediate = sandboxTimers.clearImmediate;
sandboxGlobal.setInterval = sandboxTimers.setInterval;
sandboxGlobal.clearInterval = sandboxTimers.clearInterval;

function normalizeBuiltinRequest(request) {
  return request.startsWith("node:") ? request.slice("node:".length) : request;
}

function assertAllowedModuleRequest(request) {
  if (typeof request !== "string") {
    throw createSandboxError(`Invalid require specifier: ${String(request)}`);
  }

  const normalized = normalizeBuiltinRequest(request);
  if (builtinModules.has(normalized)) {
    throw createSandboxError(
      `Access to Node builtin module '${normalized}' is not allowed in extensions`
    );
  }

  if (request === "@formula/extension-api" || request === "formula") return;

  // Support both POSIX and Windows-style relative specifiers. Extensions may be authored
  // using either `/` or `\\` separators, but we only permit relative filesystem imports.
  const isRelative =
    request.startsWith("./") || request.startsWith("../") || request.startsWith(".\\") || request.startsWith("..\\");
  if (!isRelative) {
    throw createSandboxError(
      `Only relative extension imports are allowed (attempted to require '${request}')`
    );
  }
}

function assertWithinExtensionRoot(resolvedPath, request) {
  const real = fs.realpathSync(resolvedPath);
  const relative = path.relative(extensionRoot, real);
  const inside =
    relative === "" ||
    (!relative.startsWith(".." + path.sep) && relative !== ".." && !path.isAbsolute(relative));
  if (!inside) {
    throw createSandboxError(
      `Extensions cannot require modules outside their extension folder: '${request}' resolved to '${real}'`
    );
  }
  return real;
}

function resolveAsFile(candidate) {
  try {
    const stat = fs.statSync(candidate);
    if (stat.isFile()) return candidate;
  } catch {
    // ignore
  }
  return null;
}

function resolveAsDirectory(candidate) {
  try {
    const stat = fs.statSync(candidate);
    if (!stat.isDirectory()) return null;
  } catch {
    return null;
  }

  return (
    resolveAsFile(path.join(candidate, "index.js")) ?? resolveAsFile(path.join(candidate, "index.json"))
  );
}

function resolveExtensionModulePath(request, parentFilename) {
  assertAllowedModuleRequest(request);
  if (request === "@formula/extension-api" || request === "formula") {
    return { type: "virtual", id: "@formula/extension-api" };
  }

  const baseDir = parentFilename ? path.dirname(parentFilename) : extensionRoot;
  const resolvedBase = path.resolve(baseDir, request);

  const direct = resolveAsFile(resolvedBase) ?? resolveAsDirectory(resolvedBase);
  if (direct) return { type: "file", filename: assertWithinExtensionRoot(direct, request) };

  const withJs = resolveAsFile(`${resolvedBase}.js`);
  if (withJs) return { type: "file", filename: assertWithinExtensionRoot(withJs, request) };

  const withJson = resolveAsFile(`${resolvedBase}.json`);
  if (withJson) return { type: "file", filename: assertWithinExtensionRoot(withJson, request) };

  throw createSandboxError(`Cannot find module '${request}'`);
}

function checkForDynamicImport(source, filename) {
  // `import("node:fs")` bypasses Module._load and can reach Node builtins via the ESM loader.
  // Node's vm dynamic import hooks still require flags in some runtimes; we proactively
  // scan extension code and reject any dynamic import usage.
  const src = String(source);
  const len = src.length;

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
        return out;
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

  let i = 0;
  let state = "code"; // code | single | double | template | regex | lineComment | blockComment
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
        // Distinguish property access (`obj.import()`) from spread (`...import()`).
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
        // division operator
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
        // ++ / -- prefix vs postfix changes whether a regex literal may follow.
        // If a regex is allowed here, treat it as a prefix operator (expression still expected).
        // Otherwise it's postfix (expression ended).
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
        if (ident === "import" && !afterPropertyDot) {
          const afterImport = skipWhitespaceAndComments(j);
          if (src[afterImport] === "(") {
            const argStart = skipWhitespaceAndComments(afterImport + 1);
            let specifier = null;
            if (src[argStart] === "'" || src[argStart] === '"') {
              specifier = parseStringLiteral(argStart);
            }
            const detail = specifier ? ` (attempted to import '${specifier}')` : "";
            throw createSandboxError(`Dynamic import is not allowed in extensions${detail}`);
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

      // Unknown token char, keep scanning.
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
      }
      i += 1;
      continue;
    }

    if (state === "blockComment") {
      if (ch === "*" && src[i + 1] === "/") {
        state = "code";
        i += 2;
        continue;
      }
      i += 1;
      continue;
    }
  }

  if (templateBraceStack.length > 0) {
    throw createSandboxError(
      `Unterminated template literal expression while scanning '${String(filename)}'`
    );
  }
}

const moduleCache = new Map(); // real filename -> sandbox module record
const makeRequire = vm.runInContext("(hostRequire) => (specifier) => hostRequire(specifier)", sandboxContext);
const sandboxJSONParse = vm.runInContext("JSON.parse", sandboxContext);

function loadModuleFromFile(filename) {
  const ext = path.extname(filename);
  const realFilename = fs.realpathSync(filename);

  if (moduleCache.has(realFilename)) {
    return moduleCache.get(realFilename).exports;
  }

  const module = new SandboxObject();
  module.id = realFilename;
  module.filename = realFilename;
  module.loaded = false;
  module.exports = new SandboxObject();

  moduleCache.set(realFilename, module);

  const hostRequire = hardenHostFunction((specifier) => {
    try {
      const resolved = resolveExtensionModulePath(specifier, realFilename);
      if (resolved.type === "virtual") return formulaApi;
      return loadModuleFromFile(resolved.filename);
    } catch (error) {
      if (error instanceof SandboxError) throw error;
      const err = createSandboxError(String(error?.message ?? error));
      if (error?.stack) err.stack = String(error.stack);
      throw err;
    }
  });

  const requireFn = makeRequire(hostRequire);

  if (ext === ".json") {
    try {
      const raw = fs.readFileSync(realFilename, "utf8");
      module.exports = sandboxJSONParse(String(raw));
      module.loaded = true;
      return module.exports;
    } catch (error) {
      if (error instanceof SandboxError) throw error;
      const err = createSandboxError(String(error?.message ?? error));
      if (error?.stack) err.stack = String(error.stack);
      throw err;
    }
  }

  if (ext !== ".js" && ext !== ".cjs") {
    throw createSandboxError(`Unsupported module type '${ext}' for '${realFilename}'`);
  }

  let code;
  try {
    code = fs.readFileSync(realFilename, "utf8");
  } catch (error) {
    throw createSandboxError(`Failed to read module '${realFilename}': ${String(error?.message ?? error)}`);
  }

  checkForDynamicImport(code, realFilename);

  const wrapped = `(function (exports, require, module, __filename, __dirname) {\n'use strict';\n${code}\n});`;

  let wrapperFn;
  try {
    const script = new vm.Script(wrapped, { filename: realFilename });
    wrapperFn = script.runInContext(sandboxContext);
  } catch (error) {
    const err = createSandboxError(String(error?.message ?? error));
    if (error?.stack) err.stack = String(error.stack);
    throw err;
  }

  try {
    wrapperFn.call(
      module.exports,
      module.exports,
      requireFn,
      module,
      realFilename,
      path.dirname(realFilename)
    );
    module.loaded = true;
    return module.exports;
  } catch (error) {
    if (error instanceof SandboxError) throw error;
    const err = createSandboxError(String(error?.message ?? error));
    if (error?.stack) err.stack = String(error.stack);
    throw err;
  }
}

// The extension API itself is trusted host code, but it must still be evaluated inside the VM
// realm so objects/promises returned to extensions originate from the sandbox. The API package is
// allowed to `require()` *relative* helper modules within its own folder (eg: ./src/runtime.js),
// but must not load Node builtins or arbitrary files.
const extensionApiRoot = fs.realpathSync(path.dirname(workerData.apiModulePath));
const apiModuleCache = new Map(); // real filename -> sandbox module record

function assertAllowedExtensionApiRequest(request) {
  if (typeof request !== "string") {
    throw createSandboxError(`Invalid require specifier: ${String(request)}`);
  }

  const normalized = normalizeBuiltinRequest(request);
  if (builtinModules.has(normalized)) {
    throw createSandboxError(
      `Access to Node builtin module '${normalized}' is not allowed in extensions`
    );
  }

  const isRelative =
    request.startsWith("./") ||
    request.startsWith("../") ||
    request.startsWith(".\\") ||
    request.startsWith("..\\");
  if (!isRelative) {
    throw createSandboxError("The Formula extension API cannot require external modules");
  }
}

function assertWithinExtensionApiRoot(resolvedPath, request) {
  const real = fs.realpathSync(resolvedPath);
  const relative = path.relative(extensionApiRoot, real);
  const inside =
    relative === "" ||
    (!relative.startsWith(".." + path.sep) && relative !== ".." && !path.isAbsolute(relative));
  if (!inside) {
    throw createSandboxError(
      `The Formula extension API cannot require modules outside its package folder: '${request}' resolved to '${real}'`
    );
  }
  return real;
}

function resolveExtensionApiModulePath(request, parentFilename) {
  assertAllowedExtensionApiRequest(request);

  const baseDir = parentFilename ? path.dirname(parentFilename) : extensionApiRoot;
  const resolvedBase = path.resolve(baseDir, request);

  const direct = resolveAsFile(resolvedBase) ?? resolveAsDirectory(resolvedBase);
  if (direct) return { filename: assertWithinExtensionApiRoot(direct, request) };

  const withJs = resolveAsFile(`${resolvedBase}.js`);
  if (withJs) return { filename: assertWithinExtensionApiRoot(withJs, request) };

  const withJson = resolveAsFile(`${resolvedBase}.json`);
  if (withJson) return { filename: assertWithinExtensionApiRoot(withJson, request) };

  throw createSandboxError(`Cannot find module '${request}'`);
}

function loadExtensionApiModuleFromFile(filename) {
  const ext = path.extname(filename);
  const realFilename = fs.realpathSync(filename);

  if (apiModuleCache.has(realFilename)) {
    return apiModuleCache.get(realFilename).exports;
  }

  const module = new SandboxObject();
  module.id = realFilename;
  module.filename = realFilename;
  module.loaded = false;
  module.exports = new SandboxObject();

  apiModuleCache.set(realFilename, module);

  const hostRequire = hardenHostFunction((specifier) => {
    try {
      const resolved = resolveExtensionApiModulePath(specifier, realFilename);
      return loadExtensionApiModuleFromFile(resolved.filename);
    } catch (error) {
      if (error instanceof SandboxError) throw error;
      const err = createSandboxError(String(error?.message ?? error));
      if (error?.stack) err.stack = String(error.stack);
      throw err;
    }
  });

  const requireFn = makeRequire(hostRequire);

  if (ext === ".json") {
    try {
      const raw = fs.readFileSync(realFilename, "utf8");
      module.exports = sandboxJSONParse(String(raw));
      module.loaded = true;
      return module.exports;
    } catch (error) {
      if (error instanceof SandboxError) throw error;
      const err = createSandboxError(String(error?.message ?? error));
      if (error?.stack) err.stack = String(error.stack);
      throw err;
    }
  }

  if (ext !== ".js" && ext !== ".cjs") {
    throw createSandboxError(`Unsupported module type '${ext}' for '${realFilename}'`);
  }

  let code;
  try {
    code = fs.readFileSync(realFilename, "utf8");
  } catch (error) {
    throw createSandboxError(`Failed to read module '${realFilename}': ${String(error?.message ?? error)}`);
  }

  checkForDynamicImport(code, realFilename);

  const wrapped = `(function (exports, require, module, __filename, __dirname) {\n'use strict';\n${code}\n});`;

  let wrapperFn;
  try {
    const script = new vm.Script(wrapped, { filename: realFilename });
    wrapperFn = script.runInContext(sandboxContext);
  } catch (error) {
    const err = createSandboxError(String(error?.message ?? error));
    if (error?.stack) err.stack = String(error.stack);
    throw err;
  }

  try {
    wrapperFn.call(
      module.exports,
      module.exports,
      requireFn,
      module,
      realFilename,
      path.dirname(realFilename)
    );
    module.loaded = true;
    return module.exports;
  } catch (error) {
    if (error instanceof SandboxError) throw error;
    const err = createSandboxError(String(error?.message ?? error));
    if (error?.stack) err.stack = String(error.stack);
    throw err;
  }
}

// Load the extension API module inside the sandbox so returned objects/promises
// are created within the VM realm (preventing constructor-based escapes).
const apiCode = fs.readFileSync(workerData.apiModulePath, "utf8");
checkForDynamicImport(apiCode, workerData.apiModulePath);
const apiWrapper = `(function (exports, require, module, __filename, __dirname) {\n'use strict';\n${apiCode}\n});`;

const apiModule = new SandboxObject();
apiModule.id = workerData.apiModulePath;
apiModule.filename = workerData.apiModulePath;
apiModule.loaded = false;
apiModule.exports = new SandboxObject();

let formulaApi;
try {
  const script = new vm.Script(apiWrapper, { filename: workerData.apiModulePath });
  const fn = script.runInContext(sandboxContext);
  const apiRequire = makeRequire(
    hardenHostFunction((specifier) => {
      const resolved = resolveExtensionApiModulePath(specifier, workerData.apiModulePath);
      return loadExtensionApiModuleFromFile(resolved.filename);
    })
  );
  fn.call(
    apiModule.exports,
    apiModule.exports,
    apiRequire,
    apiModule,
    workerData.apiModulePath,
    path.dirname(workerData.apiModulePath)
  );
  apiModule.loaded = true;
  formulaApi = vm.runInContext(
    "globalThis[Symbol.for('formula.extensionApi.api')]",
    sandboxContext
  );
  if (!formulaApi) {
    throw createSandboxError("@formula/extension-api runtime failed to initialize");
  }
} catch (error) {
  parentPort.postMessage({
    type: "activate_error",
    id: "startup",
    error: serializeErrorForTransport(error)
  });
  throw error;
}

const createTransport = vm.runInContext(
  "(postMessage) => ({ postMessage: (message) => postMessage(message) })",
  sandboxContext
);
const createContextObject = vm.runInContext(
  `(extensionId, extensionPath, extensionUri, globalStoragePath, workspaceStoragePath) => ({
    extensionId: String(extensionId),
    extensionPath: String(extensionPath),
    extensionUri: extensionUri == null ? undefined : String(extensionUri),
    globalStoragePath: globalStoragePath == null ? undefined : String(globalStoragePath),
    workspaceStoragePath: workspaceStoragePath == null ? undefined : String(workspaceStoragePath)
  })`,
  sandboxContext
);

formulaApi.__setTransport(createTransport(hostPostMessage));
formulaApi.__setContext(
  createContextObject(
    workerData.extensionId,
    workerData.extensionPath,
    workerData.extensionUri,
    workerData.globalStoragePath,
    workerData.workspaceStoragePath
  )
);

vm.runInContext(
  "(formulaApi) => { globalThis.fetch = (input, init) => formulaApi.network.fetch(String(input), init); }",
  sandboxContext
)(formulaApi);

// Permission-gated WebSocket: the host uses `network.openWebSocket` as a permission check.
// We expose a WebSocket implementation in the VM realm that never surfaces the native
// WebSocket instance to extension code (it is stored in a host-side WeakMap).
if (typeof globalThis.WebSocket === "function") {
  const NativeWebSocket = globalThis.WebSocket;
  const nativeWebSockets = new WeakMap(); // sandbox wrapper -> native ws

  const hostWsCreate = hardenHostFunction((wrapper, url, protocols) => {
    try {
      const ws =
        protocols === undefined ? new NativeWebSocket(url) : new NativeWebSocket(url, protocols);
      nativeWebSockets.set(wrapper, ws);

      ws.addEventListener("open", () => {
        try {
          wrapper._handleHostOpen(String(ws.protocol ?? ""), String(ws.extensions ?? ""));
        } catch {
          // ignore
        }
      });

      ws.addEventListener("message", (event) => {
        const data = typeof event?.data === "string" ? event.data : null;
        try {
          wrapper._handleHostMessage(data);
        } catch {
          // ignore
        }
      });

      ws.addEventListener("error", () => {
        try {
          wrapper._handleHostError();
        } catch {
          // ignore
        }
      });

      ws.addEventListener("close", (event) => {
        nativeWebSockets.delete(wrapper);
        try {
          // Never pass host objects into the VM realm. Any host object (even a plain
          // `{}`) exposes the host's `Object`/`Function` constructors via
          // `obj.constructor.constructor(...)`, which is a full sandbox escape.
          wrapper._handleHostClose(
            Number(event?.code ?? 1006),
            String(event?.reason ?? ""),
            Boolean(event?.wasClean)
          );
        } catch {
          // ignore
        }
      });
    } catch (error) {
      throw createSandboxError(String(error?.message ?? error));
    }
  });

  const hostWsSend = hardenHostFunction((wrapper, data) => {
    try {
      const ws = nativeWebSockets.get(wrapper);
      if (!ws) {
        throw new Error("WebSocket is not open");
      }
      ws.send(data);
    } catch (error) {
      throw createSandboxError(String(error?.message ?? error));
    }
  });

  const hostWsClose = hardenHostFunction((wrapper, code, reason) => {
    try {
      const ws = nativeWebSockets.get(wrapper);
      if (!ws) return;
      ws.close(code, reason);
    } catch (error) {
      throw createSandboxError(String(error?.message ?? error));
    }
  });

  const installSandboxWebSocket = vm.runInContext(
    `
((formulaApi, wsCreate, wsSend, wsClose) => {
  class PermissionedWebSocket {
    static CONNECTING = 0;
    static OPEN = 1;
    static CLOSING = 2;
    static CLOSED = 3;

    constructor(url, protocols) {
      this._url = String(url ?? "");
      this._protocols = protocols;
      this._readyState = PermissionedWebSocket.CONNECTING;
      this._binaryType = "blob";
      this._protocol = "";
      this._extensions = "";
      this._bufferedAmount = 0;
      this._pendingClose = null;
      this._listeners = new Map();

      this.onopen = null;
      this.onmessage = null;
      this.onerror = null;
      this.onclose = null;

      void this._start();
    }

    get url() {
      return this._url;
    }

    get readyState() {
      return this._readyState;
    }

    get bufferedAmount() {
      return this._bufferedAmount;
    }

    get extensions() {
      return this._extensions;
    }

    get protocol() {
      return this._protocol;
    }

    get binaryType() {
      return this._binaryType;
    }

    set binaryType(value) {
      this._binaryType = String(value ?? "blob");
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
      if (this.readyState !== PermissionedWebSocket.OPEN) {
        throw new Error("WebSocket is not open");
      }
      wsSend(this, data);
    }

    close(code, reason) {
      if (this.readyState === PermissionedWebSocket.CLOSED) return;
      if (this.readyState === PermissionedWebSocket.CONNECTING) {
        this._pendingClose = { code, reason };
        this._readyState = PermissionedWebSocket.CLOSING;
        return;
      }
      this._readyState = PermissionedWebSocket.CLOSING;
      wsClose(this, code, reason);
    }

    async _start() {
      try {
        if (typeof formulaApi?.network?.openWebSocket !== "function") {
          throw new Error("WebSocket API is not available in this extension runtime");
        }
        await formulaApi.network.openWebSocket(this._url);
      } catch (err) {
        this._fail(err);
        return;
      }

      try {
        wsCreate(this, this._url, this._protocols);
      } catch (err) {
        this._fail(err);
      }
    }

    _handleHostOpen(protocol, extensions) {
      this._readyState = PermissionedWebSocket.OPEN;
      this._protocol = String(protocol ?? "");
      this._extensions = String(extensions ?? "");
      this._emit("open", { type: "open", target: this });

      const pendingClose = this._pendingClose;
      if (pendingClose) {
        this._pendingClose = null;
        try {
          wsClose(this, pendingClose.code, pendingClose.reason);
        } catch {
          // ignore
        }
      }
    }

    _handleHostMessage(data) {
      this._emit("message", { type: "message", data, target: this });
    }

    _handleHostError() {
      this._emit("error", { type: "error", target: this });
    }

    _handleHostClose(code, reason, wasClean) {
      this._readyState = PermissionedWebSocket.CLOSED;
      this._emit("close", {
        type: "close",
        code: Number(code ?? 1006),
        reason: String(reason ?? ""),
        wasClean: Boolean(wasClean),
        target: this
      });
    }

    _fail(err) {
      this._readyState = PermissionedWebSocket.CLOSED;
      this._emit("error", { type: "error", error: err, target: this });
      this._emit("close", {
        type: "close",
        code: 1008,
        reason: String(err?.message ?? err),
        wasClean: false,
        target: this
      });
    }

    _emit(type, event) {
      const evt = event && typeof event === "object" ? event : { type };
      const propHandler = this[\`on\${type}\`];
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

  globalThis.WebSocket = PermissionedWebSocket;
})
`,
    sandboxContext
  );

  installSandboxWebSocket(formulaApi, hostWsCreate, hostWsSend, hostWsClose);
}

function toSandbox(value, seen = new Map()) {
  if (value === null) return null;
  const type = typeof value;
  if (type === "function") {
    throw createSandboxError("Cannot transfer functions into the extension sandbox");
  }
  if (type !== "object") return value;

  if (seen.has(value)) return seen.get(value);

  if (Array.isArray(value)) {
    const arr = new SandboxArray();
    seen.set(value, arr);
    for (const item of value) {
      arr.push(toSandbox(item, seen));
    }
    return arr;
  }

  const obj = new SandboxObject();
  seen.set(value, obj);
  for (const [k, v] of Object.entries(value)) {
    obj[k] = toSandbox(v, seen);
  }
  return obj;
}

let extensionModule = null;
let activated = false;
let activationPromise = null;

async function activateExtension() {
  if (activated) return;
  if (activationPromise) return activationPromise;

  activationPromise = (async () => {
    if (!extensionModule) {
      const resolvedMain = assertWithinExtensionRoot(mainPath, mainPath);
      extensionModule = loadModuleFromFile(resolvedMain);
    }

    const activateFn = extensionModule.activate || extensionModule.default?.activate;
    if (typeof activateFn !== "function") {
      throw createSandboxError("Extension entrypoint does not export an activate() function");
    }

    const context = new SandboxObject();
    context.extensionId = String(workerData.extensionId);
    context.extensionPath = String(workerData.extensionPath);
    if (workerData.extensionUri != null) context.extensionUri = String(workerData.extensionUri);
    if (workerData.globalStoragePath != null) context.globalStoragePath = String(workerData.globalStoragePath);
    if (workerData.workspaceStoragePath != null)
      context.workspaceStoragePath = String(workerData.workspaceStoragePath);
    context.subscriptions = new SandboxArray();

    await activateFn(context);
    activated = true;
  })();

  try {
    await activationPromise;
  } finally {
    // Allow retries after failures (e.g. transient errors) while keeping successful activations fast.
    if (!activated) activationPromise = null;
  }
}

function safeParentPostMessage(payload) {
  try {
    parentPort.postMessage(payload);
  } catch {
    // ignore
  }
}

parentPort.on("message", (message) => {
  void (async () => {
    if (!message || typeof message !== "object") return;

    if (message.type === "activate") {
      try {
        await activateExtension();
        safeParentPostMessage({ type: "activate_result", id: message.id });
      } catch (error) {
        safeParentPostMessage({
          type: "activate_error",
          id: message.id,
          error: serializeErrorForTransport(error)
        });
      }
      return;
    }

    try {
      await Promise.resolve(formulaApi.__handleMessage(toSandbox(message)));
    } catch {
      // ignore
    }
  })().catch(() => {
    // ignore
  });
});
