import ts from "typescript";

/**
 * @typedef {{ sheetName: string, address: string }} Selection
 */

/**
 * @param {Selection} selection
 */
function serializeSelection(selection) {
  return {
    sheetName: JSON.stringify(selection.sheetName),
    address: JSON.stringify(selection.address),
  };
}

function formatTypeScriptDiagnostics(diagnostics) {
  const errors = diagnostics.filter((d) => d.category === ts.DiagnosticCategory.Error);
  if (errors.length === 0) return null;

  const host = {
    getCanonicalFileName: (fileName) => fileName,
    getCurrentDirectory: () => "",
    getNewLine: () => "\n",
  };

  return ts.formatDiagnostics(errors, host);
}

/**
 * Build a TypeScript program that:
 * - bootstraps the Formula workbook API (`ctx`)
 * - compiles + executes the user code inside an async function
 * - flushes any queued workbook event handlers before completing
 *
 * The output is intended to run in a sandbox where `globalThis.__hostRpc`
 * is defined (capability-based host RPC).
 *
 * @param {{
 *   code: string,
 *   activeSheetName: string,
 *   selection: Selection,
 * }} params
 */
export function buildBootstrapJavaScript({ activeSheetName, selection }) {
  const activeSheet = JSON.stringify(activeSheetName);
  const sel = serializeSelection(selection);

  // NOTE: This string is executed inside the sandbox (worker/vm). Keep it
  // deterministic and free of host-only dependencies.
  return `"use strict";
// Formula Script Runtime bootstrap.
// This file is generated at runtime; it must be deterministic and sandbox-safe.

const __hostRpc = globalThis.__hostRpc;
if (typeof __hostRpc !== "function") {
  throw new Error("Formula scripting sandbox misconfigured: missing globalThis.__hostRpc");
}

function __rpc(method, params) {
  return __hostRpc(method, params);
}

function __assertPositiveInt(value, label) {
  if (!Number.isFinite(value) || value <= 0 || !Number.isInteger(value)) {
    throw new Error(label + " must be a positive integer. Received: " + value);
  }
}

function __colLabel(index1) {
  // 1-based column index -> A1 label.
  __assertPositiveInt(index1, "col");
  let n = index1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function __cellA1(row1, col1) {
  __assertPositiveInt(row1, "row");
  __assertPositiveInt(col1, "col");
  return __colLabel(col1) + String(row1);
}

function __createRangeProxy(sheetName, address) {
  return {
    address,
    getValues: () => __rpc("range.getValues", { sheetName, address }),
    setValues: (values) => __rpc("range.setValues", { sheetName, address, values }),
    getFormulas: () => __rpc("range.getFormulas", { sheetName, address }),
    setFormulas: (formulas) => __rpc("range.setFormulas", { sheetName, address, formulas }),
    getFormats: () => __rpc("range.getFormats", { sheetName, address }),
    setFormats: (formats) => __rpc("range.setFormats", { sheetName, address, formats }),
    getValue: () => __rpc("range.getValue", { sheetName, address }),
    setValue: (value) => __rpc("range.setValue", { sheetName, address, value }),
    getFormat: () => __rpc("range.getFormat", { sheetName, address }),
    setFormat: (format) => __rpc("range.setFormat", { sheetName, address, format }),
  };
}

function __createSheetProxy(name) {
  return {
    name,
    getRange: (address) => __createRangeProxy(name, address),
    getCell: (row, col) => __createRangeProxy(name, __cellA1(row, col)),
    getUsedRange: async () => {
      const address = await __rpc("sheet.getUsedRange", { sheetName: name });
      return __createRangeProxy(name, address);
    },
  };
}

function __createWorkbookProxy() {
  return {
    getSheets: async () => {
      const names = await __rpc("workbook.getSheets", null);
      return names.map((n) => __createSheetProxy(n));
    },
    getSheet: (name) => __createSheetProxy(name),
    getActiveSheetName: () => __rpc("workbook.getActiveSheetName", null),
    getSelection: () => __rpc("workbook.getSelection", null),
    setSelection: (sheetName, address) => __rpc("workbook.setSelection", { sheetName, address }),
  };
}

const __eventListeners = new Map();
let __eventChain = Promise.resolve();

function __addEventListener(eventType, handler) {
  if (typeof handler !== "function") {
    throw new Error("Event handler must be a function");
  }
  let bucket = __eventListeners.get(eventType);
  if (!bucket) {
    bucket = new Set();
    __eventListeners.set(eventType, bucket);
  }
  bucket.add(handler);
  return () => bucket.delete(handler);
}

function __dispatchEvent(eventType, payload) {
  const bucket = __eventListeners.get(eventType);
  if (!bucket || bucket.size === 0) return;

  const handlers = Array.from(bucket);
  __eventChain = __eventChain.then(async () => {
    for (const handler of handlers) {
      await handler(payload);
    }
  });
}

globalThis.__formulaDispatchEvent = __dispatchEvent;

async function __flushEvents() {
  // Ensure we wait for events that are queued while we're already flushing.
  while (true) {
    const snapshot = __eventChain;
    await snapshot;
    if (snapshot === __eventChain) return;
  }
}

 const ctx = {
   workbook: __createWorkbookProxy(),
   activeSheet: __createSheetProxy(${activeSheet}),
   selection: __createRangeProxy(${sel.sheetName}, ${sel.address}),
   ui: {
     log: (...args) => console.log(...args),
     alert: (message) => __rpc("ui.alert", { message }),
     confirm: (message) => __rpc("ui.confirm", { message }),
     prompt: (message, defaultValue) => __rpc("ui.prompt", { message, defaultValue }),
   },
   alert: (message) => __rpc("ui.alert", { message }),
   confirm: (message) => __rpc("ui.confirm", { message }),
   prompt: (message, defaultValue) => __rpc("ui.prompt", { message, defaultValue }),
   fetch: typeof globalThis.fetch === "function" ? (...args) => globalThis.fetch(...args) : undefined,
   console: globalThis.console,
   events: {
     onEdit: (handler) => __addEventListener("edit", handler),
     onSelectionChange: (handler) => __addEventListener("selectionChange", handler),
     onFormatChange: (handler) => __addEventListener("formatChange", handler),
     flush: () => __flushEvents(),
   },
  };
`;
}

/**
 * @param {string} code
 */
function isModuleScript(code) {
  return /\bexport\s+default\b/.test(code);
}

function findUnsupportedModuleSyntax(code) {
  const source = ts.createSourceFile("user-script.ts", code, ts.ScriptTarget.ES2022, true, ts.ScriptKind.TS);

  /** @type {string[]} */
  const staticImports = [];
  /** @type {string[]} */
  const dynamicImports = [];

  /** @param {import("typescript").Node} node */
  function visit(node) {
    if (ts.isImportDeclaration(node) && ts.isStringLiteral(node.moduleSpecifier)) {
      staticImports.push(node.moduleSpecifier.text);
    }

    if (ts.isExportDeclaration(node) && node.moduleSpecifier && ts.isStringLiteral(node.moduleSpecifier)) {
      staticImports.push(node.moduleSpecifier.text);
    }

    if (ts.isCallExpression(node) && node.expression.kind === ts.SyntaxKind.ImportKeyword) {
      const [arg] = node.arguments;
      if (arg && ts.isStringLiteral(arg)) {
        dynamicImports.push(arg.text);
      } else {
        dynamicImports.push("<dynamic>");
      }
    }

    ts.forEachChild(node, visit);
  }

  visit(source);
  return { staticImports, dynamicImports };
}

/**
 * @param {{ code: string }} params
 * @returns {{ kind: "script" | "module", ts: string, moduleKind: "none" | "commonjs" }}
 */
export function buildTypeScriptProgram({ code }) {
  const { staticImports, dynamicImports } = findUnsupportedModuleSyntax(code);
  if (dynamicImports.length > 0) {
    const specifier = dynamicImports[0];
    const err = new Error(`Dynamic import() is not supported in scripts (${specifier})`);
    err.name = "DynamicImportNotSupportedError";
    throw err;
  }
  if (staticImports.length > 0) {
    const specifier = staticImports[0];
    const err = new Error(`Imports are not supported in scripts (${specifier})`);
    err.name = "ImportsNotSupportedError";
    throw err;
  }

  if (isModuleScript(code)) {
    return {
      kind: "module",
      moduleKind: "commonjs",
      ts: code
    };
  }

  return {
    kind: "script",
    moduleKind: "none",
    ts: `// Formula Script Runtime wrapper.

async function __formulaUserMain(ctx) {
${code}
}

async function __formulaRuntimeMain() {
  const result = await __formulaUserMain(ctx);
  await ctx.events.flush();
  return result;
}

__formulaRuntimeMain();
`
  };
}

/**
 * Convenience helper for workers: returns the bootstrap JS plus the TS wrapper.
 *
 * @param {{ code: string, activeSheetName: string, selection: Selection }} params
 */
export function buildSandboxedScript({ code, activeSheetName, selection }) {
  const program = buildTypeScriptProgram({ code });
  return {
    bootstrap: buildBootstrapJavaScript({ activeSheetName, selection }),
    ts: program.ts,
    moduleKind: program.moduleKind,
    kind: program.kind
  };
}

export function buildModuleRunnerJavaScript({ moduleJs }) {
  // Wrap the CommonJS-emitted module in a closure so `exports`/`module` exist
  // before TypeScript's prologue executes.
  return `// Formula Script Runtime wrapper (module style).
// Module scripts must export a default async function main(ctx).

const __formulaExports = {};
const __formulaModule = { exports: __formulaExports };

(function(exports, module) {
${moduleJs}
})(__formulaExports, __formulaModule);

async function __formulaRuntimeMain() {
  const __formulaMain = __formulaModule.exports?.default ?? __formulaExports.default;
  if (typeof __formulaMain !== "function") {
    throw new Error("Script must export a default function");
  }
  const result = await __formulaMain(ctx);
  await ctx.events.flush();
  return result;
}

__formulaRuntimeMain();
`;
}

/**
 * @param {string} tsSource
 * @param {{ moduleKind?: "none" | "commonjs" }} [options]
 * @returns {{ js: string }}
 */
export function transpileTypeScript(tsSource, options = {}) {
  const moduleKind = options.moduleKind ?? "none";
  const result = ts.transpileModule(tsSource, {
    compilerOptions: {
      target: ts.ScriptTarget.ES2022,
      module: moduleKind === "commonjs" ? ts.ModuleKind.CommonJS : ts.ModuleKind.None,
    },
    reportDiagnostics: true,
    fileName: "user-script.ts",
  });

  const formatted = formatTypeScriptDiagnostics(result.diagnostics ?? []);
  if (formatted) {
    const err = new Error(formatted);
    err.name = "TypeScriptCompileError";
    throw err;
  }

  return { js: result.outputText };
}

export function serializeError(err) {
  if (err instanceof Error) {
    return {
      name: err.name,
      message: err.message,
      stack: err.stack,
    };
  }
  if (typeof err === "string") {
    return { message: err };
  }
  try {
    return { message: JSON.stringify(err) };
  } catch {
    return { message: "Unknown error" };
  }
}
