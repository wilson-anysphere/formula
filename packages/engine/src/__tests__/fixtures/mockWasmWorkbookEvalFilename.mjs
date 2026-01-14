// Minimal ESM module that emulates a wasm-bindgen build of `crates/formula-wasm`,
// but implements just enough evaluation to exercise workbook file metadata plumbing.
//
// This is loaded by `packages/engine/src/engine.worker.ts` via dynamic import (runtime string),
// so it must be plain JS (not TS) to avoid relying on Vite/Vitest transforms.

export default async function init() {}

function normalizeSheet(sheet) {
  return typeof sheet === "string" && sheet.trim() !== "" ? sheet : "Sheet1";
}

function normalizeInput(value) {
  if (value === undefined || value === null) return null;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
  return null;
}

function isFormula(value) {
  return typeof value === "string" && value.trimStart().startsWith("=");
}

function parseCellFilenameReferenceArg(formula) {
  const body = formula.trim().replace(/^=/, "").trim();
  const m = /^CELL\s*\(\s*\"filename\"\s*(?:,\s*(.+?)\s*)?\)\s*$/i.exec(body);
  if (!m) return null;
  const raw = typeof m[1] === "string" ? m[1].trim() : "";
  return raw ? raw : "";
}

function sheetNameFromReferenceArg(arg) {
  const raw = String(arg ?? "").trim();
  if (!raw) return null;

  const bang = raw.lastIndexOf("!");
  if (bang < 0) return null;
  const token = raw.slice(0, bang).trim();
  if (!token) return null;

  const quoted = /^'((?:[^']|'')+)'$/.exec(token);
  if (quoted) return quoted[1].replace(/''/g, "'").trim() || null;

  return token;
}

function matchesInfoDirectory(formula) {
  const body = formula.trim().replace(/^=/, "").trim();
  return /^INFO\s*\(\s*\"directory\"\s*\)\s*$/i.test(body);
}

function workbookDirForExcel(dir) {
  const d = String(dir ?? "");
  if (d === "") return "";
  if (d.endsWith("/") || d.endsWith("\\")) return d;
  const lastSlash = d.lastIndexOf("/");
  const lastBackslash = d.lastIndexOf("\\");
  let sep = "/";
  if (lastSlash >= 0 && lastBackslash >= 0) {
    sep = lastSlash > lastBackslash ? "/" : "\\";
  } else if (lastBackslash >= 0) {
    sep = "\\";
  }
  return d + sep;
}

export class WasmWorkbook {
  constructor() {
    this.directory = null;
    this.filename = null;
    this.infoDirectory = null;
    this.inputsBySheet = new Map();
    this.valuesBySheet = new Map();
  }

  _sheetMap(map, sheet) {
    const key = normalizeSheet(sheet);
    let m = map.get(key);
    if (!m) {
      m = new Map();
      map.set(key, m);
    }
    return m;
  }

  toJson() {
    return "{}";
  }

  setWorkbookFileMetadata(directory, filename) {
    this.directory = directory ?? null;
    this.filename = filename ?? null;
  }

  setEngineInfo(info) {
    if (!info || typeof info !== "object") {
      throw new Error("setEngineInfo: info must be an object");
    }
    if ("directory" in info) {
      const raw = info.directory;
      this.infoDirectory = typeof raw === "string" && raw.trim() !== "" ? raw.trim() : null;
    }
  }

  setCell(address, value, sheet) {
    const sheetName = normalizeSheet(sheet);
    const input = normalizeInput(value);
    const inputs = this._sheetMap(this.inputsBySheet, sheetName);
    const values = this._sheetMap(this.valuesBySheet, sheetName);

    if (input == null) {
      inputs.delete(address);
      values.delete(address);
      return;
    }

    inputs.set(address, input);
    // Recalculate updates values; for now clear any stale value.
    values.delete(address);
  }

  getCell(address, sheet) {
    const sheetName = normalizeSheet(sheet);
    const inputs = this._sheetMap(this.inputsBySheet, sheetName);
    const values = this._sheetMap(this.valuesBySheet, sheetName);
    return {
      sheet: sheetName,
      address,
      input: inputs.get(address) ?? null,
      value: values.get(address) ?? null,
    };
  }

  getRange(_range, _sheet) {
    return [];
  }

  setRange(_range, _values, _sheet) {}

  recalculate(sheet) {
    const targetSheet = sheet ? normalizeSheet(sheet) : null;
    const changes = [];

    const sheetEntries = targetSheet
      ? [[targetSheet, this._sheetMap(this.inputsBySheet, targetSheet)]]
      : Array.from(this.inputsBySheet.entries());

    for (const [sheetName, inputs] of sheetEntries) {
      const values = this._sheetMap(this.valuesBySheet, sheetName);
      for (const [address, input] of inputs.entries()) {
        let computed = null;
        if (isFormula(input)) {
          const refArg = parseCellFilenameReferenceArg(input);
          if (refArg !== null) {
            if (!this.filename) {
              computed = "";
            } else {
              const dirRaw = typeof this.directory === "string" ? this.directory : "";
              const dir = dirRaw.trim() !== "" ? workbookDirForExcel(dirRaw) : "";
              const refSheetName = sheetNameFromReferenceArg(refArg);
              computed = dir ? `${dir}[${this.filename}]${refSheetName ?? sheetName}` : `[${this.filename}]${refSheetName ?? sheetName}`;
            }
          } else if (matchesInfoDirectory(input)) {
            if (this.infoDirectory) {
              computed = workbookDirForExcel(this.infoDirectory);
            } else {
              const dirRaw = typeof this.directory === "string" ? this.directory : "";
              const dir = dirRaw.trim() !== "" ? workbookDirForExcel(dirRaw) : "";
              computed = this.filename && dir ? dir : "#N/A";
            }
          }
        } else {
          computed = input;
        }

        if (values.get(address) !== computed) {
          values.set(address, computed);
          changes.push({ sheet: sheetName, address, value: computed });
        }
      }
    }

    return changes;
  }

  static fromJson(_json) {
    return new WasmWorkbook();
  }
}

// Editor-tooling exports (unused by these tests, but included to satisfy worker expectations).
export function lexFormula() {
  return [];
}
export function parseFormulaPartial() {
  return { ast: null, error: null, context: { function: null } };
}
