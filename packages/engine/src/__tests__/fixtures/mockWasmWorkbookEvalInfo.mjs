// Minimal ESM module that emulates a wasm-bindgen build of `crates/formula-wasm`,
// but implements just enough evaluation to exercise `INFO()` metadata plumbing.
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

function parseInfoKey(formula) {
  const body = formula.trim().replace(/^=/, "").trim();
  const match = /^INFO\s*\(\s*\"([^\"]*)\"\s*\)\s*$/i.exec(body);
  if (!match) return null;
  return match[1];
}

function normalizeOriginA1(origin) {
  const s = String(origin ?? "").trim();
  if (s === "") return null;

  const match = /^\$?([A-Za-z]+)\$?([1-9][0-9]*)$/.exec(s);
  if (!match) {
    throw new Error("origin must be an A1 address");
  }
  const [, colRaw, rowRaw] = match;
  const row = Number(rowRaw);
  if (!Number.isInteger(row) || row < 1) {
    throw new Error("origin must be an A1 address");
  }
  const col = colRaw.toUpperCase();
  return `$${col}$${row}`;
}

function legacyOriginForExcel(origin) {
  const s = String(origin ?? "").trim();
  if (s === "") return "";
  const match = /^\$?([A-Za-z]+)\$?([1-9][0-9]*)$/.exec(s);
  if (!match) return s;
  const [, colRaw, rowRaw] = match;
  const row = Number(rowRaw);
  if (!Number.isInteger(row) || row < 1) return s;
  const col = colRaw.toUpperCase();
  return `$${col}$${row}`;
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

function normalizeInfoString(value, { key, allowEmpty = false } = {}) {
  if (value === null || value === undefined) return null;
  if (typeof value !== "string") {
    throw new Error(`setEngineInfo: ${key} must be a string`);
  }
  const trimmed = value.trim();
  if (trimmed === "" && !allowEmpty) return null;
  return trimmed;
}

function normalizeInfoNumber(value, { key } = {}) {
  if (value === null || value === undefined) return null;
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`setEngineInfo: ${key} must be a finite number`);
  }
  return value;
}

export class WasmWorkbook {
  constructor() {
    this.inputsBySheet = new Map();
    this.valuesBySheet = new Map();

    // INFO metadata.
    this.system = null;
    this.directory = null;
    this.osversion = null;
    this.release = null;
    this.version = null;
    this.memavail = null;
    this.totmem = null;
    // Excel INFO("origin") is derived from the sheet view state (top-left visible cell).
    // Model it per-sheet and default to "$A$1".
    this.sheetOriginBySheet = new Map();

    // Legacy `EngineInfo.origin` / `origin_by_sheet` plumbing (still supported by the real engine).
    this.infoOrigin = null;
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

  setEngineInfo(info) {
    if (!info || typeof info !== "object") {
      throw new Error("setEngineInfo: info must be an object");
    }

    if ("system" in info) this.system = normalizeInfoString(info.system, { key: "system" });
    if ("directory" in info) this.directory = normalizeInfoString(info.directory, { key: "directory" });
    if ("osversion" in info) this.osversion = normalizeInfoString(info.osversion, { key: "osversion" });
    if ("release" in info) this.release = normalizeInfoString(info.release, { key: "release" });
    if ("version" in info) this.version = normalizeInfoString(info.version, { key: "version" });
    if ("memavail" in info) this.memavail = normalizeInfoNumber(info.memavail, { key: "memavail" });
    if ("totmem" in info) this.totmem = normalizeInfoNumber(info.totmem, { key: "totmem" });
  }

  setSheetOrigin(sheet, origin) {
    const sheetName = normalizeSheet(sheet);
    const normalized = normalizeOriginA1(origin);
    if (normalized == null) {
      this.sheetOriginBySheet.delete(sheetName);
    } else {
      this.sheetOriginBySheet.set(sheetName, normalized);
    }
  }

  setInfoOrigin(origin) {
    this.infoOrigin = normalizeInfoString(origin, { key: "origin" });
  }

  setInfoOriginForSheet(sheet, origin) {
    // The real wasm engine treats `setInfoOriginForSheet` as a legacy alias for `setSheetOrigin`,
    // not as a string-based metadata fallback. Keep the mock in sync so tests reflect production
    // behavior (A1 normalization + per-sheet precedence).
    this.setSheetOrigin(sheet, origin);
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

  _infoValue(sheetName, key) {
    const k = String(key ?? "").trim().toLowerCase();
    switch (k) {
      case "system":
        return this.system ?? "pcdos";
      case "directory":
        if (this.directory != null && this.directory !== "") {
          return workbookDirForExcel(this.directory);
        }
        return "#N/A";
      case "osversion":
        return this.osversion ?? "#N/A";
      case "release":
        return this.release ?? "#N/A";
      case "version":
        return this.version ?? "#N/A";
      case "memavail":
        return typeof this.memavail === "number" && Number.isFinite(this.memavail) ? this.memavail : "#N/A";
      case "totmem":
        return typeof this.totmem === "number" && Number.isFinite(this.totmem) ? this.totmem : "#N/A";
      case "origin": {
        const sheetOrigin = this.sheetOriginBySheet.get(sheetName);
        if (sheetOrigin != null) return sheetOrigin;

        if (this.infoOrigin != null) return legacyOriginForExcel(this.infoOrigin);

        return "$A$1";
      }
      default:
        return "#VALUE!";
    }
  }

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
          const infoKey = parseInfoKey(input);
          if (infoKey != null) {
            computed = this._infoValue(sheetName, infoKey);
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
