import { showToast } from "../../extensions/ui.js";
import { createLazyImport } from "../../startup/lazyImport.js";

// Translation tables from the Rust engine (canonical <-> localized function names).
// Keep these in sync with `crates/formula-engine/src/locale/data/*.tsv`.
//
// We use these to support formula-bar hints for localized function names like
// de-DE `SUMME(...)` by mapping them back to canonical signatures (`SUM`).
import DE_DE_FUNCTION_TSV from "../../../../../crates/formula-engine/src/locale/data/de-DE.tsv?raw";
import ES_ES_FUNCTION_TSV from "../../../../../crates/formula-engine/src/locale/data/es-ES.tsv?raw";
import FR_FR_FUNCTION_TSV from "../../../../../crates/formula-engine/src/locale/data/fr-FR.tsv?raw";
import { normalizeFormulaLocaleId } from "../../spreadsheet/formulaLocale.js";

type FunctionParam = { name: string; optional?: boolean };

type FunctionSignature = {
  name: string;
  params: FunctionParam[];
  summary: string;
};

type CatalogFunction = {
  name: string;
  min_args: number;
  max_args: number;
  arg_types?: string[];
};

type FunctionCatalogModule = { default: { functions?: CatalogFunction[] } };

const loadFunctionCatalogModule = createLazyImport(() => import("../../../../../shared/functionCatalog.mjs"), {
  label: "Function catalog",
  onError: (err) => {
    console.error("[formula][desktop] Failed to load function catalog:", err);
    try {
      showToast("Failed to load function catalog. Please reload the app.", "error");
    } catch {
      // ignore
    }
  },
});

let catalogByName: Map<string, CatalogFunction> | null = null;
let catalogByNameInitPromise: Promise<Map<string, CatalogFunction> | null> | null = null;

async function ensureCatalogByName(): Promise<Map<string, CatalogFunction> | null> {
  if (catalogByName) return catalogByName;
  if (!catalogByNameInitPromise) {
    catalogByNameInitPromise = (async () => {
      const mod = (await loadFunctionCatalogModule()) as FunctionCatalogModule | null;
      if (!mod) return null;
      const map = new Map<string, CatalogFunction>();
      for (const fn of mod.default?.functions ?? []) {
        if (fn?.name) map.set(fn.name.toUpperCase(), fn);
      }
      catalogByName = map;
      return catalogByName;
    })().finally(() => {
      catalogByNameInitPromise = null;
    });
  }
  return catalogByNameInitPromise;
}

/**
 * Best-effort: begin loading the function catalog in the background. Callers that render function
 * hints can invoke this when the user starts editing a formula.
 */
export function preloadFunctionSignatureCatalog(): Promise<void> {
  return ensureCatalogByName()
    .then(() => {})
    .catch(() => {
      // Best-effort: preloading is opportunistic; callers should treat failures as a no-op.
    });
}

export function isFunctionSignatureCatalogReady(): boolean {
  return Boolean(catalogByName && catalogByName.size > 0);
}

function casefoldIdent(ident: string): string {
  // Mirror Rust's locale behavior (`casefold_ident` / `casefold`): Unicode-aware uppercasing.
  return String(ident ?? "").toUpperCase();
}

type FunctionTranslationMap = Map<string, string>;

function parseFunctionTranslationsTsv(tsv: string): FunctionTranslationMap {
  const localizedToCanonical: FunctionTranslationMap = new Map();
  for (const rawLine of String(tsv ?? "").split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;
    const [canonical, localized] = line.split("\t");
    if (!canonical || !localized) continue;
    const canonUpper = casefoldIdent(canonical.trim());
    const locUpper = casefoldIdent(localized.trim());
    // Only store translations that differ; identity entries can fall back to `casefoldIdent`.
    if (canonUpper && locUpper && canonUpper !== locUpper) {
      localizedToCanonical.set(locUpper, canonUpper);
    }
  }
  return localizedToCanonical;
}

const FUNCTION_TRANSLATIONS_BY_LOCALE: Record<string, FunctionTranslationMap> = {
  "de-DE": parseFunctionTranslationsTsv(DE_DE_FUNCTION_TSV),
  "fr-FR": parseFunctionTranslationsTsv(FR_FR_FUNCTION_TSV),
  "es-ES": parseFunctionTranslationsTsv(ES_ES_FUNCTION_TSV),
};

const FUNCTION_SIGNATURE_CACHE = new Map<string, FunctionSignature | null>();
const CATALOG_SIGNATURE_CACHE = new Map<string, FunctionSignature | null>();

const FUNCTION_SIGNATURES: Record<string, FunctionSignature> = {
  DATE: {
    name: "DATE",
    params: [{ name: "year" }, { name: "month" }, { name: "day" }],
    summary: "Returns the serial number of a particular date.",
  },
  DAY: {
    name: "DAY",
    params: [{ name: "serial_number" }],
    summary: "Converts a serial number to a day of the month.",
  },
  MONTH: {
    name: "MONTH",
    params: [{ name: "serial_number" }],
    summary: "Converts a serial number to a month.",
  },
  YEAR: {
    name: "YEAR",
    params: [{ name: "serial_number" }],
    summary: "Converts a serial number to a year.",
  },
  SUM: {
    name: "SUM",
    params: [{ name: "number1" }, { name: "number2", optional: true }],
    summary: "Adds all the numbers in a range of cells.",
  },
  COUNT: {
    name: "COUNT",
    params: [{ name: "value1" }, { name: "value2", optional: true }],
    summary: "Counts the number of cells that contain numbers.",
  },
  COUNTA: {
    name: "COUNTA",
    params: [{ name: "value1" }, { name: "value2", optional: true }],
    summary: "Counts the number of non-empty cells.",
  },
  COUNTBLANK: {
    name: "COUNTBLANK",
    params: [{ name: "range" }],
    summary: "Counts the number of blank cells within a range.",
  },
  COUNTIF: {
    name: "COUNTIF",
    params: [{ name: "range" }, { name: "criteria" }],
    summary: "Counts the number of cells within a range that meet the given criteria.",
  },
  AVERAGE: {
    name: "AVERAGE",
    params: [{ name: "number1" }, { name: "number2", optional: true }],
    summary: "Returns the average (arithmetic mean) of its arguments.",
  },
  MAX: {
    name: "MAX",
    params: [{ name: "number1" }, { name: "number2", optional: true }],
    summary: "Returns the largest value in a set of values.",
  },
  MIN: {
    name: "MIN",
    params: [{ name: "number1" }, { name: "number2", optional: true }],
    summary: "Returns the smallest value in a set of values.",
  },
  ROUND: {
    name: "ROUND",
    params: [{ name: "number" }, { name: "num_digits" }],
    summary: "Rounds a number to a specified number of digits.",
  },
  ROUNDUP: {
    name: "ROUNDUP",
    params: [{ name: "number" }, { name: "num_digits" }],
    summary: "Rounds a number up, away from zero.",
  },
  ROUNDDOWN: {
    name: "ROUNDDOWN",
    params: [{ name: "number" }, { name: "num_digits" }],
    summary: "Rounds a number down, toward zero.",
  },
  SUMPRODUCT: {
    name: "SUMPRODUCT",
    params: [{ name: "array1" }, { name: "array2" }],
    summary: "Returns the sum of the products of corresponding array components.",
  },
  IF: {
    name: "IF",
    params: [
      { name: "logical_test" },
      { name: "value_if_true" },
      { name: "value_if_false", optional: true },
    ],
    summary: "Checks whether a condition is met, and returns one value if TRUE and another value if FALSE.",
  },
  IFERROR: {
    name: "IFERROR",
    params: [{ name: "value" }, { name: "value_if_error" }],
    summary: "Returns a value you specify if a formula evaluates to an error; otherwise returns the formula result.",
  },
  IFNA: {
    name: "IFNA",
    params: [{ name: "value" }, { name: "value_if_na" }],
    summary: "Returns a value you specify if a formula evaluates to #N/A; otherwise returns the formula result.",
  },
  ISERROR: {
    name: "ISERROR",
    params: [{ name: "value" }],
    summary: "Checks whether a value is an error.",
  },
  NA: {
    name: "NA",
    params: [],
    summary: "Returns the #N/A error value.",
  },
  VLOOKUP: {
    name: "VLOOKUP",
    params: [
      { name: "lookup_value" },
      { name: "table_array" },
      { name: "col_index_num" },
      { name: "range_lookup", optional: true },
    ],
    summary: "Looks for a value in the leftmost column of a table, then returns a value in the same row from a specified column.",
  },
  HLOOKUP: {
    name: "HLOOKUP",
    params: [
      { name: "lookup_value" },
      { name: "table_array" },
      { name: "row_index_num" },
      { name: "range_lookup", optional: true },
    ],
    summary: "Looks for a value in the top row of a table, then returns a value in the same column from a specified row.",
  },
  XLOOKUP: {
    name: "XLOOKUP",
    params: [
      { name: "lookup_value" },
      { name: "lookup_array" },
      { name: "return_array" },
      { name: "if_not_found", optional: true },
      { name: "match_mode", optional: true },
      { name: "search_mode", optional: true },
    ],
    summary: "Looks up a value in a range or an array.",
  },
  INDEX: {
    name: "INDEX",
    params: [
      { name: "array" },
      { name: "row_num" },
      { name: "column_num", optional: true },
    ],
    summary: "Returns the value of an element in a table or an array.",
  },
  MATCH: {
    name: "MATCH",
    params: [
      { name: "lookup_value" },
      { name: "lookup_array" },
      { name: "match_type", optional: true },
    ],
    summary: "Looks up values in a reference or array.",
  },
  TODAY: {
    name: "TODAY",
    params: [],
    summary: "Returns the current date.",
  },
  NOW: {
    name: "NOW",
    params: [],
    summary: "Returns the current date and time.",
  },
  RAND: {
    name: "RAND",
    params: [],
    summary: "Returns a random number between 0 and 1.",
  },
  RANDBETWEEN: {
    name: "RANDBETWEEN",
    params: [{ name: "bottom" }, { name: "top" }],
    summary: "Returns a random integer between the numbers you specify.",
  },
  SEQUENCE: {
    name: "SEQUENCE",
    params: [
      { name: "rows" },
      { name: "columns", optional: true },
      { name: "start", optional: true },
      { name: "step", optional: true },
    ],
    summary: "Generates a list of sequential numbers in an array.",
  },
  TAKE: {
    name: "TAKE",
    params: [{ name: "array" }, { name: "rows", optional: true }, { name: "columns", optional: true }],
    summary: "Returns a specified number of contiguous rows or columns from the start or end of an array.",
  },
  DROP: {
    name: "DROP",
    params: [{ name: "array" }, { name: "rows", optional: true }, { name: "columns", optional: true }],
    summary: "Excludes a specified number of rows or columns from the start or end of an array.",
  },
  CHOOSECOLS: {
    name: "CHOOSECOLS",
    params: [{ name: "array" }, { name: "col_num1" }, { name: "col_num2", optional: true }, { name: "…", optional: true }],
    summary: "Returns the specified columns from an array.",
  },
  CHOOSEROWS: {
    name: "CHOOSEROWS",
    params: [{ name: "array" }, { name: "row_num1" }, { name: "row_num2", optional: true }, { name: "…", optional: true }],
    summary: "Returns the specified rows from an array.",
  },
  EXPAND: {
    name: "EXPAND",
    params: [
      { name: "array" },
      { name: "rows" },
      { name: "columns", optional: true },
      { name: "pad_with", optional: true },
    ],
    summary: "Expands an array to the specified row and column dimensions.",
  },
  TRANSPOSE: {
    name: "TRANSPOSE",
    params: [{ name: "array" }],
    summary: "Returns the transpose of an array or range.",
  },
  CONCAT: {
    name: "CONCAT",
    params: [{ name: "text1" }, { name: "text2", optional: true }],
    summary: "Combines the text from multiple ranges and/or strings.",
  },
  CONCATENATE: {
    name: "CONCATENATE",
    params: [{ name: "text1" }, { name: "text2", optional: true }],
    summary: "Combines several text strings into one text string.",
  },
  LEFT: {
    name: "LEFT",
    params: [{ name: "text" }, { name: "num_chars", optional: true }],
    summary: "Returns the leftmost characters from a text string.",
  },
  RIGHT: {
    name: "RIGHT",
    params: [{ name: "text" }, { name: "num_chars", optional: true }],
    summary: "Returns the rightmost characters from a text string.",
  },
  MID: {
    name: "MID",
    params: [{ name: "text" }, { name: "start_num" }, { name: "num_chars" }],
    summary: "Returns a specific number of characters from a text string starting at the position you specify.",
  },
  LEN: {
    name: "LEN",
    params: [{ name: "text" }],
    summary: "Returns the number of characters in a text string.",
  },
  TRIM: {
    name: "TRIM",
    params: [{ name: "text" }],
    summary: "Removes leading/trailing spaces and reduces multiple internal spaces to a single space.",
  },
  UPPER: {
    name: "UPPER",
    params: [{ name: "text" }],
    summary: "Converts text to uppercase.",
  },
  LOWER: {
    name: "LOWER",
    params: [{ name: "text" }],
    summary: "Converts text to lowercase.",
  },
  FIND: {
    name: "FIND",
    params: [{ name: "find_text" }, { name: "within_text" }, { name: "start_num", optional: true }],
    summary: "Finds one text string within another (case-sensitive).",
  },
  SEARCH: {
    name: "SEARCH",
    params: [{ name: "find_text" }, { name: "within_text" }, { name: "start_num", optional: true }],
    summary: "Finds one text string within another (not case-sensitive).",
  },
  SUBSTITUTE: {
    name: "SUBSTITUTE",
    params: [
      { name: "text" },
      { name: "old_text" },
      { name: "new_text" },
      { name: "instance_num", optional: true },
    ],
    summary: "Substitutes new text for old text in a text string.",
  },
  TEXTSPLIT: {
    name: "TEXTSPLIT",
    params: [
      { name: "text" },
      { name: "col_delimiter" },
      { name: "row_delimiter", optional: true },
      { name: "ignore_empty", optional: true },
      { name: "match_mode", optional: true },
      { name: "pad_with", optional: true },
    ],
    summary: "Splits text into rows and columns using delimiters and returns an array.",
  },
};

export function getFunctionSignature(name: string, opts: { localeId?: string } = {}): FunctionSignature | null {
  const requested = casefoldIdent(name);
  const lookup = requested.startsWith("_XLFN.") ? requested.slice("_XLFN.".length) : requested;

  // If the requested name is localized (e.g. `SUMME` in de-DE), map it back to the canonical
  // name (`SUM`) so we can reuse the curated signature list / function catalog metadata.
  const localeId =
    opts.localeId?.trim?.() ||
    (typeof document !== "undefined" ? document.documentElement?.lang : "")?.trim?.() ||
    "en-US";
  const formulaLocaleId = normalizeFormulaLocaleId(localeId);

  // Cache by the *effective* formula locale ID so language-only / variant locale IDs
  // can reuse signatures (e.g. `de`, `de_DE.UTF-8`, `de-AT` -> `de-DE`).
  const cacheKey = `${formulaLocaleId ?? localeId}\0${requested}`;
  if (FUNCTION_SIGNATURE_CACHE.has(cacheKey)) {
    return FUNCTION_SIGNATURE_CACHE.get(cacheKey) ?? null;
  }

  const localeMap = formulaLocaleId ? FUNCTION_TRANSLATIONS_BY_LOCALE[formulaLocaleId] : undefined;
  const canonical = localeMap?.get(lookup) ?? lookup;

  const known = FUNCTION_SIGNATURES[canonical] ?? signatureFromCatalog(canonical);
  if (!known) return null;

  // Preserve any `_xlfn.` prefix *and* localized naming in the displayed name so formula-bar hints
  // match the text users see/typed.
  const result = requested === canonical ? known : { ...known, name: requested };
  FUNCTION_SIGNATURE_CACHE.set(cacheKey, result);
  return result;
}

type SignaturePart = { text: string; kind: "name" | "param" | "paramActive" | "punct" };

export function signatureParts(
  sig: FunctionSignature,
  activeParamIndex: number | null,
  opts?: { argSeparator?: string }
): SignaturePart[] {
  const argSeparator = opts?.argSeparator ?? ", ";
  const parts: SignaturePart[] = [{ text: `${sig.name}(`, kind: "name" }];
  sig.params.forEach((param, index) => {
    if (index > 0) parts.push({ text: argSeparator, kind: "punct" });
    const isActive = activeParamIndex !== null && activeParamIndex === index;
    parts.push({
      text: param.optional ? `[${param.name}]` : param.name,
      kind: isActive ? "paramActive" : "param",
    });
  });
  parts.push({ text: ")", kind: "punct" });
  return parts;
}

function signatureFromCatalog(name: string): FunctionSignature | null {
  if (!catalogByName) {
    // Kick off lazy-load for next time; the catalog is large and we avoid parsing it on initial render.
    void preloadFunctionSignatureCatalog();
    return null;
  }

  if (CATALOG_SIGNATURE_CACHE.has(name)) {
    return CATALOG_SIGNATURE_CACHE.get(name) ?? null;
  }

  const fn = catalogByName.get(name);
  if (!fn) {
    CATALOG_SIGNATURE_CACHE.set(name, null);
    return null;
  }

  const sig: FunctionSignature = {
    name,
    params: buildParams(fn.min_args, fn.max_args, fn.arg_types),
    summary: "",
  };
  CATALOG_SIGNATURE_CACHE.set(name, sig);
  return sig;
}

function buildParams(minArgs: number, maxArgs: number, argTypes: string[] | undefined): FunctionParam[] {
  const MAX_PARAMS = 5;

  if (!Number.isFinite(minArgs) || !Number.isFinite(maxArgs) || minArgs < 0 || maxArgs < 0) {
    return [];
  }

  if (maxArgs <= MAX_PARAMS) {
    const out: FunctionParam[] = [];
    for (let i = 1; i <= maxArgs; i++) {
      out.push({ name: paramNameFromCatalogTypes(i, maxArgs, argTypes), optional: i > minArgs });
    }
    return out;
  }

  const requiredShown = Math.min(minArgs, MAX_PARAMS - 1);
  const out: FunctionParam[] = [];
  for (let i = 1; i <= requiredShown; i++) out.push({ name: paramNameFromCatalogTypes(i, maxArgs, argTypes) });

  if (minArgs > requiredShown) {
    out.push({ name: "…" });
    return out;
  }

  out.push({ name: "…", optional: true });
  return out;
}

function paramNameFromCatalogTypes(index1: number, maxArgs: number, argTypes: string[] | undefined): string {
  const index0 = index1 - 1;
  if (!Array.isArray(argTypes) || argTypes.length === 0) return `arg${index1}`;

  let valueType: string | undefined;
  if (argTypes.length === 1 && maxArgs > 1) {
    valueType = argTypes[0];
  } else {
    valueType = argTypes[index0] ?? argTypes[argTypes.length - 1];
  }

  switch (valueType) {
    case "number":
      return `number${index1}`;
    case "text":
      return `text${index1}`;
    case "bool":
      return `logical${index1}`;
    case "any":
      return `value${index1}`;
    default:
      return `arg${index1}`;
  }
}
