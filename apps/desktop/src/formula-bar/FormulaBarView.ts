import { FormulaBarModel } from "./FormulaBarModel.js";
import { type RangeAddress } from "../spreadsheet/a1.js";
import { searchFunctionResults } from "../command-palette/commandPaletteSearch.js";
import FUNCTION_NAMES from "../../../../shared/functionNames.mjs";
import {
  assignFormulaReferenceColors,
  extractFormulaReferences,
  tokenizeFormula,
  toggleA1AbsoluteAtCursor,
  type ExtractFormulaReferencesOptions,
  type FormulaReferenceRange,
} from "@formula/spreadsheet-frontend";
import type { EngineClient, FormulaParseOptions } from "@formula/engine";
import { ContextMenu, type ContextMenuItem } from "../menus/contextMenu.js";
import {
  getFunctionSignature,
  isFunctionSignatureCatalogReady,
  preloadFunctionSignatureCatalog,
  signatureParts,
} from "./highlight/functionSignatures.js";
import {
  normalizeFormulaLocaleId,
  type FormulaLocaleId,
} from "../spreadsheet/formulaLocale.js";

// Translation tables from the Rust engine (canonical <-> localized function names).
// Keep these in sync with `crates/formula-engine/src/locale/data/*.tsv`.
import DE_DE_FUNCTION_TSV from "../../../../crates/formula-engine/src/locale/data/de-DE.tsv?raw";
import ES_ES_FUNCTION_TSV from "../../../../crates/formula-engine/src/locale/data/es-ES.tsv?raw";
import FR_FR_FUNCTION_TSV from "../../../../crates/formula-engine/src/locale/data/fr-FR.tsv?raw";

type FixFormulaErrorWithAiInfo = {
  address: string;
  /** The committed formula text currently stored in the active cell. */
  input: string;
  /** The current formula bar draft (may differ from `input` while editing). */
  draft: string;
  value: unknown;
  explanation: NonNullable<ReturnType<FormulaBarModel["errorExplanation"]>>;
};

type FormulaReferenceHighlight = {
  range: FormulaReferenceRange;
  color: string;
  text: string;
  index: number;
  active?: boolean;
};

type ReferenceHighlightMode = "editing" | "errorPanel" | "none";

type NameBoxMenuItem = {
  /**
   * User-visible label. For named ranges/tables this is typically the workbook-defined name.
   */
  label: string;
  /**
   * Navigation reference. When provided, selecting the item behaves like typing the reference
   * into the name box + Enter.
   *
   * When omitted/null, selection falls back to populating + selecting the name box input text.
   */
  reference?: string | null;
  enabled?: boolean;
};

const FORMULA_BAR_EXPANDED_STORAGE_KEY = "formula:ui:formulaBarExpanded";
const FORMULA_BAR_MIN_HEIGHT = 24;
const FORMULA_BAR_MAX_HEIGHT_COLLAPSED = 140;
const FORMULA_BAR_MAX_HEIGHT_EXPANDED = 360;

let formulaBarExpandedFallback: boolean | null = null;

function getFormulaBarSessionStorage(): Storage | null {
  try {
    if (typeof window === "undefined") return null;
    return window.sessionStorage ?? null;
  } catch {
    return null;
  }
}

function loadFormulaBarExpandedState(): boolean {
  const storage = getFormulaBarSessionStorage();
  if (storage) {
    try {
      const raw = storage.getItem(FORMULA_BAR_EXPANDED_STORAGE_KEY);
      if (raw === "true") return true;
      if (raw === "false") return false;
    } catch {
      // ignore
    }
  }
  return formulaBarExpandedFallback ?? false;
}

function storeFormulaBarExpandedState(expanded: boolean): void {
  const storage = getFormulaBarSessionStorage();
  if (storage) {
    try {
      storage.setItem(FORMULA_BAR_EXPANDED_STORAGE_KEY, expanded ? "true" : "false");
      return;
    } catch {
      // Fall back to in-memory storage.
    }
  }

  formulaBarExpandedFallback = expanded;
}

const INDENT_WIDTH = 2;
const MAX_INDENT_LEVEL = 20;

function isWhitespace(ch: string): boolean {
  return ch === " " || ch === "\t" || ch === "\n" || ch === "\r";
}

/**
 * Computes indentation (spaces) to insert after an Alt+Enter newline while editing
 * a formula.
 *
 * The indentation level is derived from formula structure up to the caret:
 * - Count `(` minus `)` while ignoring any parentheses that appear inside string literals.
 * - Clamp indentation to a reasonable maximum so pathological inputs don't create huge whitespace runs.
 */
function computeFormulaIndentation(text: string, cursor: number): string {
  const cursorPos = Math.max(0, Math.min(cursor, text.length));

  let parenDepth = 0;
  let lastSignificantOutsideString: string | null = null;
  let inString = false;

  const tokens = tokenizeFormula(text);

  for (const token of tokens) {
    if (token.start >= cursorPos) break;

    if (token.type === "string") {
      // `tokenizeFormula` returns unterminated strings as a single token that runs to the
      // end of the input. If the cursor is at the end of an unterminated string, treat it
      // as "inside" the string so we avoid inserting indentation into the literal.
      const unterminated = !token.text.endsWith('"');
      if (token.start < cursorPos && (cursorPos < token.end || (unterminated && cursorPos === token.end))) {
        inString = true;
      }
      continue;
    }

    if (token.type === "punctuation") {
      if (token.text === "(") parenDepth += 1;
      else if (token.text === ")") parenDepth = Math.max(0, parenDepth - 1);
    }

    if (token.type !== "whitespace") {
      lastSignificantOutsideString = token.text;
    }
  }

  // Don't auto-indent when breaking a string literal; extra spaces would become part of the string.
  if (inString) return "";

  // Optional continuation indent: if the preceding token is a comma but we're not inside parentheses
  // (e.g. array constants / union expressions), add a single indentation level.
  if (parenDepth === 0 && (lastSignificantOutsideString === "," || lastSignificantOutsideString === ";")) {
    parenDepth = 1;
  }

  const clamped = Math.max(0, Math.min(parenDepth, MAX_INDENT_LEVEL));
  return " ".repeat(clamped * INDENT_WIDTH);
}

type NameBoxDropdownItemKind = "namedRange" | "table" | "recent";

type FunctionPickerItem = { name: string; signature?: string; summary?: string };

type FunctionSignature = NonNullable<ReturnType<typeof getFunctionSignature>>;

// `shared/functionNames.mjs` is a generated, already-sorted list (codepoint order). Avoid
// re-sorting at runtime on the hot startup path.
const ALL_FUNCTION_NAMES_SORTED: string[] = Array.isArray(FUNCTION_NAMES) ? FUNCTION_NAMES : [];

const COMMON_FUNCTION_NAMES = [
  "SUM",
  "AVERAGE",
  "COUNT",
  "MIN",
  "MAX",
  "IF",
  "IFERROR",
  "XLOOKUP",
  "VLOOKUP",
  "INDEX",
  "MATCH",
  "TODAY",
  "NOW",
  "ROUND",
  "CONCAT",
  "SEQUENCE",
  "FILTER",
  "TEXTSPLIT",
];

const DEFAULT_FUNCTION_NAMES: string[] = (() => {
  const byName = new Map(ALL_FUNCTION_NAMES_SORTED.map((name) => [name.toUpperCase(), name]));
  const out: string[] = [];
  const seen = new Set<string>();

  for (const name of COMMON_FUNCTION_NAMES) {
    const resolved = byName.get(name.toUpperCase());
    if (!resolved) continue;
    out.push(resolved);
    seen.add(resolved.toUpperCase());
  }
  for (const name of ALL_FUNCTION_NAMES_SORTED) {
    if (out.length >= 50) break;
    const key = name.toUpperCase();
    if (seen.has(key)) continue;
    out.push(name);
    seen.add(key);
  }
  return out;
})();

function buildFunctionPickerItems(query: string, limit: number, localeId: string): FunctionPickerItem[] {
  const trimmed = String(query ?? "").trim();
  const cappedLimit = Math.max(0, Math.floor(limit));
  if (cappedLimit === 0) return [];

  if (!trimmed) {
    const tables = getFunctionTranslationTables(localeId);
    const out: FunctionPickerItem[] = [];
    const seen = new Set<string>();
    for (const canonicalName of DEFAULT_FUNCTION_NAMES) {
      if (out.length >= cappedLimit) break;
      const upper = canonicalName.toUpperCase();
      const localized = tables?.canonicalToLocalized.get(upper) ?? canonicalName;
      const key = localized.toUpperCase();
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(functionPickerItemFromName(localized, localeId));
    }
    return out;
  }

  return searchFunctionResults(trimmed, { limit: cappedLimit, localeId }).map((res) => functionPickerItemFromName(res.name, localeId));
}

function renderFunctionPickerList(opts: {
  listEl: HTMLUListElement;
  query: string;
  items: FunctionPickerItem[];
  selectedIndex: number;
  onSelect: (index: number) => void;
}): HTMLLIElement[] {
  const { listEl, query, items, selectedIndex, onSelect } = opts;

  listEl.innerHTML = "";

  if (items.length === 0) {
    const empty = document.createElement("li");
    empty.className = "command-palette__empty";
    empty.textContent = query.trim() ? "No matching functions" : "Type to search functions";
    empty.setAttribute("role", "presentation");
    listEl.appendChild(empty);
    return [];
  }

  const itemEls: HTMLLIElement[] = [];
  for (let i = 0; i < items.length; i += 1) {
    const fn = items[i]!;
    const li = document.createElement("li");
    li.className = "command-palette__item";
    li.dataset.testid = `formula-function-picker-item-${fn.name}`;
    li.setAttribute("role", "option");
    li.setAttribute("aria-selected", i === selectedIndex ? "true" : "false");

    const icon = document.createElement("div");
    icon.className = "command-palette__item-icon command-palette__item-icon--function";
    icon.textContent = "fx";

    const main = document.createElement("div");
    main.className = "command-palette__item-main";

    const label = document.createElement("div");
    label.className = "command-palette__item-label";
    label.textContent = fn.name;

    main.appendChild(label);

    const signatureOrSummary = (() => {
      const summary = fn.summary?.trim?.() ?? "";
      const signature = fn.signature?.trim?.() ?? "";
      if (signature && summary) return `${signature} — ${summary}`;
      if (signature) return signature;
      if (summary) return summary;
      return "";
    })();

    if (signatureOrSummary) {
      const desc = document.createElement("div");
      desc.className = "command-palette__item-description command-palette__item-description--mono";
      desc.textContent = signatureOrSummary;
      main.appendChild(desc);
    }

    li.appendChild(icon);
    li.appendChild(main);

    li.addEventListener("mousedown", (e) => {
      // Keep focus in the search input so we can handle selection consistently.
      e.preventDefault();
    });
    li.addEventListener("click", () => onSelect(i));

    listEl.appendChild(li);
    itemEls.push(li);
  }

  return itemEls;
}

function functionPickerItemFromName(name: string, localeId: string): FunctionPickerItem {
  const sig = getFunctionSignature(name, { localeId });
  const signature = sig ? formatSignature(sig, localeId) : undefined;
  const summary = sig?.summary?.trim?.() ? sig.summary.trim() : undefined;
  return { name, signature, summary };
}

function formatSignature(sig: FunctionSignature, localeId: string): string {
  const parts = signatureParts(sig, null, { argSeparator: inferArgSeparator(localeId) });
  return parts.map((p) => p.text).join("");
}

type CompletionContext = {
  /**
   * Inclusive start index of the identifier-like token in the input.
   */
  replaceStart: number;
  /**
   * Exclusive end index of the identifier-like token in the input.
   */
  replaceEnd: number;
  /**
   * Token prefix typed before the caret (used for filtering + casing preservation).
   */
  typedPrefix: string;
  /**
   * Optional function qualifier prefix (e.g. `_xlfn.`) that is preserved during insertion.
   */
  qualifier: string;
  /**
   * Uppercased prefix used for matching against the catalog.
   */
  matchPrefixUpper: string;
};

type FunctionSuggestion = { name: string; signature: string };

let AUTOCOMPLETE_INSTANCE_ID = 0;

const FUNCTION_ENTRIES: Array<{ name: string; upper: string }> = (() => {
  // Deduplicate case-insensitively; we shouldn't surface multiple entries for the same
  // function name just because the catalog casing differs.
  const out: Array<{ name: string; upper: string }> = [];
  const seen = new Set<string>();
  for (const name of ALL_FUNCTION_NAMES_SORTED) {
    const upper = name.toUpperCase();
    if (seen.has(upper)) continue;
    seen.add(upper);
    out.push({ name, upper });
  }
  return out;
})();

const FUNCTION_NAMES_UPPER = new Set(FUNCTION_ENTRIES.map((e) => e.upper));

type FunctionTranslationTables = {
  localizedToCanonical: Map<string, string>;
  canonicalToLocalized: Map<string, string>;
  localizedNamesUpper: Set<string>;
};

function casefoldIdent(ident: string): string {
  // Mirror Rust's locale behavior (`casefold_ident` / `casefold`): Unicode-aware uppercasing.
  return String(ident ?? "").toUpperCase();
}

function parseFunctionTranslationsTsv(tsv: string): FunctionTranslationTables {
  const localizedToCanonical: Map<string, string> = new Map();
  const canonicalToLocalized: Map<string, string> = new Map();
  const localizedNamesUpper: Set<string> = new Set();

  for (const rawLine of String(tsv ?? "").split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;
    const [canonical, localized] = line.split("\t");
    if (!canonical || !localized) continue;

    const canonUpper = casefoldIdent(canonical.trim());
    const localizedTrimmed = localized.trim();
    const locUpper = casefoldIdent(localizedTrimmed);
    if (!canonUpper || !locUpper) continue;

    // Only store translations that differ; identity entries can fall back to `casefoldIdent`.
    if (canonUpper !== locUpper) {
      localizedToCanonical.set(locUpper, canonUpper);
      canonicalToLocalized.set(canonUpper, localizedTrimmed);
    }
    localizedNamesUpper.add(locUpper);
  }

  return { localizedToCanonical, canonicalToLocalized, localizedNamesUpper };
}

const FUNCTION_TRANSLATIONS_BY_LOCALE: Record<FormulaLocaleId, FunctionTranslationTables> = {
  "de-DE": parseFunctionTranslationsTsv(DE_DE_FUNCTION_TSV),
  "fr-FR": parseFunctionTranslationsTsv(FR_FR_FUNCTION_TSV),
  "es-ES": parseFunctionTranslationsTsv(ES_ES_FUNCTION_TSV),
  // Minimal locales (no translated function names yet, but they are valid engine locale ids).
  "ja-JP": { localizedToCanonical: new Map(), canonicalToLocalized: new Map(), localizedNamesUpper: new Set() },
  "zh-CN": { localizedToCanonical: new Map(), canonicalToLocalized: new Map(), localizedNamesUpper: new Set() },
  "zh-TW": { localizedToCanonical: new Map(), canonicalToLocalized: new Map(), localizedNamesUpper: new Set() },
  "ko-KR": { localizedToCanonical: new Map(), canonicalToLocalized: new Map(), localizedNamesUpper: new Set() },
  // Canonical locale: no translations needed, but keep the table shape consistent.
  "en-US": { localizedToCanonical: new Map(), canonicalToLocalized: new Map(), localizedNamesUpper: new Set() },
};

function getFunctionTranslationTables(localeId: string): FunctionTranslationTables | null {
  const formulaLocaleId = normalizeFormulaLocaleId(localeId);
  if (!formulaLocaleId) return null;
  return FUNCTION_TRANSLATIONS_BY_LOCALE[formulaLocaleId] ?? null;
}

function isKnownFunctionNameUpper(nameUpper: string, localeId: string): boolean {
  if (FUNCTION_NAMES_UPPER.has(nameUpper)) return true;
  const tables = getFunctionTranslationTables(localeId);
  return Boolean(tables?.localizedNamesUpper.has(nameUpper));
}

const ARG_SEPARATOR_CACHE = new Map<string, string>();

function inferArgSeparator(localeId: string): string {
  // Prefer the formula engine's normalized locale IDs so UI separators match parsing semantics
  // for language/region variants (e.g. `de-CH` is treated as `de-DE` by the engine today).
  const locale = normalizeFormulaLocaleId(localeId) ?? "en-US";
  const cached = ARG_SEPARATOR_CACHE.get(locale);
  if (cached) return cached;

  try {
    const parts = new Intl.NumberFormat(locale).formatToParts(1.1);
    const decimal = parts.find((p) => p.type === "decimal")?.value ?? ".";
    const sep = decimal === "," ? "; " : ", ";
    ARG_SEPARATOR_CACHE.set(locale, sep);
    return sep;
  } catch {
    return ", ";
  }
}

const SIGNATURE_PREVIEW_CACHE = new Map<string, string>();

const UNICODE_LETTER_RE: RegExp | null = (() => {
  try {
    return new RegExp("^\\p{Alphabetic}$", "u");
  } catch {
    return null;
  }
})();

const UNICODE_ALNUM_RE: RegExp | null = (() => {
  try {
    return new RegExp("^[\\p{Alphabetic}\\p{Number}]$", "u");
  } catch {
    return null;
  }
})();

function isUnicodeAlphabetic(ch: string): boolean {
  if (UNICODE_LETTER_RE) return UNICODE_LETTER_RE.test(ch);
  return (ch >= "A" && ch <= "Z") || (ch >= "a" && ch <= "z");
}

function isUnicodeAlphanumeric(ch: string): boolean {
  if (UNICODE_ALNUM_RE) return UNICODE_ALNUM_RE.test(ch);
  return isUnicodeAlphabetic(ch) || (ch >= "0" && ch <= "9");
}

function isIdentifierChar(ch: string): boolean {
  // Match the formula tokenizer's identifier rules closely enough for completion.
  // Excel function names allow dots (e.g. `COVARIANCE.P`) and digits (e.g. `LOG10`).
  return (
    ch === "_" ||
    ch === "." ||
    isUnicodeAlphanumeric(ch)
  );
}

function clampCursor(input: string, cursorPosition: number): number {
  if (!Number.isInteger(cursorPosition)) return input.length;
  if (cursorPosition < 0) return 0;
  if (cursorPosition > input.length) return input.length;
  return cursorPosition;
}

function firstNonWhitespaceIndex(text: string): number {
  for (let i = 0; i < text.length; i += 1) {
    if (!isWhitespace(text[i]!)) return i;
  }
  return -1;
}

function isFormulaText(text: string): boolean {
  const firstNonWhitespace = firstNonWhitespaceIndex(text);
  return firstNonWhitespace >= 0 && text[firstNonWhitespace] === "=";
}

function findCompletionContext(input: string, cursorPosition: number, localeId: string): CompletionContext | null {
  const cursor = clampCursor(input, cursorPosition);

  if (!isFormulaText(input)) return null;

  // Require a collapsed selection (caller ensures selectionStart === selectionEnd).
  let replaceStart = cursor;
  while (replaceStart > 0 && isIdentifierChar(input[replaceStart - 1]!)) replaceStart -= 1;

  let replaceEnd = cursor;
  while (replaceEnd < input.length && isIdentifierChar(input[replaceEnd]!)) replaceEnd += 1;

  const fullToken = input.slice(replaceStart, replaceEnd);
  const fullTokenUpper = fullToken.toUpperCase();
  // Avoid showing function autocomplete when the caret is inside a likely A1 cell reference
  // (e.g. `=A1+1` with the caret between `A` and `1`). In these cases the dropdown can
  // unexpectedly consume Escape/Tab, making the formula bar feel "stuck".
  if (/^[A-Za-z]{1,3}\d{1,7}$/.test(fullToken) && !isKnownFunctionNameUpper(fullTokenUpper, localeId)) return null;

  const typedPrefix = input.slice(replaceStart, cursor);
  if (typedPrefix.length < 1) return null;

  // Only trigger on identifier-looking starts.
  // (We handle `_xlfn.` separately below.)
  const firstChar = typedPrefix[0] ?? "";
  if (!(firstChar === "_" || isUnicodeAlphabetic(firstChar))) return null;

  // Avoid suggesting functions while the caret is inside a likely A1-style cell reference
  // (e.g. `=A1`, `=XFD1048576`). This prevents the autocomplete dropdown from stealing
  // Escape/Tab semantics while users edit references.
  const fullIdent = input.slice(replaceStart, replaceEnd);
  const fullUpper = fullIdent.toUpperCase();
  if (/^[A-Z]{1,3}[0-9]+$/.test(fullUpper) && !isKnownFunctionNameUpper(fullUpper, localeId)) return null;

  // Ensure we're at the start of an expression-like position:
  // `=VLO`, `=1+VLO`, `=SUM(VLO`, `=SUM(A, VLO)`
  let prev = replaceStart - 1;
  while (prev >= 0 && isWhitespace(input[prev]!)) prev -= 1;
  if (prev < 0) return null;

  const prevChar = input[prev]!;
  const startsExpression =
    prevChar === "=" || prevChar === "(" || prevChar === "," || prevChar === ";" || "+-*/^&=><%@".includes(prevChar);
  if (!startsExpression) return null;

  // In argument positions (after `(` or `,`), very short alphabetic identifiers are
  // much more likely to be column/range refs (e.g. `SUM(A` / `SUM(AB`) than function
  // names. Be conservative here so we don't steal Tab from range completion.
  if (prevChar === "(" || prevChar === "," || prevChar === ";") {
    if (/^[A-Za-z]+$/.test(typedPrefix)) {
      if (typedPrefix.length === 1 && !isKnownFunctionNameUpper(typedPrefix.toUpperCase(), localeId)) return null;
      if (typedPrefix.length === 2 && !isKnownFunctionNameUpper(typedPrefix.toUpperCase(), localeId)) return null;
    }
  }

  // Support Excel `_xlfn.` function prefix in pasted formulas.
  const upper = typedPrefix.toUpperCase();
  const qualifierUpper = "_XLFN.";
  if (upper.startsWith(qualifierUpper)) {
    const qualifier = typedPrefix.slice(0, qualifierUpper.length);
    const rest = typedPrefix.slice(qualifierUpper.length);
    return {
      replaceStart,
      replaceEnd,
      typedPrefix,
      qualifier,
      matchPrefixUpper: rest.toUpperCase(),
    };
  }

  return {
    replaceStart,
    replaceEnd,
    typedPrefix,
    qualifier: "",
    matchPrefixUpper: upper,
  };
}

function signaturePreview(name: string, localeId: string): string {
  const cacheKey = `${normalizeFormulaLocaleId(localeId) ?? localeId}\0${name}`;
  const cached = SIGNATURE_PREVIEW_CACHE.get(cacheKey);
  if (cached) return cached;

  const sig = getFunctionSignature(name, { localeId });
  if (!sig) {
    const fallback = "(…)";
    // When the signature catalog is still loading, avoid permanently caching the fallback.
    // (Once the catalog is ready, a subsequent call can fill the cache with the real signature.)
    if (isFunctionSignatureCatalogReady()) {
      SIGNATURE_PREVIEW_CACHE.set(cacheKey, fallback);
    }
    return fallback;
  }

  // The dropdown already shows the function name; display just the argument list for
  // a compact "signature preview" (Excel-like).
  const parts = signatureParts(sig, null, { argSeparator: inferArgSeparator(localeId) });
  if (parts.length < 2) return "(…)";

  // `signatureParts` yields: `${NAME}(` + [params/separators] + `)`.
  const inner = parts
    .slice(1, -1)
    .map((p) => p.text)
    .join("");
  const args = `(${inner})`;
  const summary = sig.summary?.trim?.() ?? "";
  const out = summary ? `${args} — ${summary}` : args;
  SIGNATURE_PREVIEW_CACHE.set(cacheKey, out);
  return out;
}

function preserveTypedCasing(typedPrefix: string, canonical: string): string {
  if (!typedPrefix) return canonical;
  if (typedPrefix.length >= canonical.length) return typedPrefix;

  // Infer case preference from the *letters* the user typed (ignore digits, dots, underscores).
  // This yields nicer results for common patterns like:
  //   "=vlo"  -> "=vlookup("
  //   "=VLO"  -> "=VLOOKUP("
  //   "=Vlo"  -> "=Vlookup("
  let letters = "";
  for (const ch of typedPrefix) {
    if (isUnicodeAlphabetic(ch)) letters += ch;
  }
  if (!letters) return typedPrefix + canonical.slice(typedPrefix.length);

  const lower = letters.toLowerCase();
  const upper = letters.toUpperCase();
  if (letters === lower) return canonical.toLowerCase();
  if (letters === upper) return canonical.toUpperCase();

  // Title-ish casing: first letter uppercase, remainder lowercase.
  if (letters[0] === upper[0] && letters.slice(1) === lower.slice(1)) {
    const lowered = canonical.toLowerCase();
    let firstLetterIdx = -1;
    for (let i = 0; i < lowered.length; i += 1) {
      if (isUnicodeAlphabetic(lowered[i]!)) {
        firstLetterIdx = i;
        break;
      }
    }
    if (firstLetterIdx >= 0) {
      return lowered.slice(0, firstLetterIdx) + lowered[firstLetterIdx]!.toUpperCase() + lowered.slice(firstLetterIdx + 1);
    }
    return lowered;
  }

  // Fallback: preserve the exact prefix the user typed and append the canonical tail.
  return typedPrefix + canonical.slice(typedPrefix.length);
}

function buildSuggestions(params: {
  prefixUpper: string;
  limit: number;
  localeId: string;
  allowLocalized: boolean;
}): FunctionSuggestion[] {
  const out: FunctionSuggestion[] = [];
  const { prefixUpper, limit, localeId, allowLocalized } = params;

  const tables = allowLocalized ? getFunctionTranslationTables(localeId) : null;
  const suppressCanonical = new Set<string>();

  if (tables) {
    // Prefer localized aliases when they match the prefix, and suppress the canonical
    // version of those same functions to avoid duplicate suggestions.
    for (const fn of FUNCTION_ENTRIES) {
      const localized = tables.canonicalToLocalized.get(fn.upper);
      if (!localized) continue;
      const localizedUpper = casefoldIdent(localized);
      if (prefixUpper && !localizedUpper.startsWith(prefixUpper)) continue;
      out.push({ name: localized, signature: signaturePreview(localized, localeId) });
      suppressCanonical.add(fn.upper);
      if (out.length >= limit) return out;
    }
  }

  for (const fn of FUNCTION_ENTRIES) {
    if (suppressCanonical.has(fn.upper)) continue;
    if (prefixUpper && !fn.upper.startsWith(prefixUpper)) continue;
    out.push({ name: fn.name, signature: signaturePreview(fn.name, localeId) });
    if (out.length >= limit) break;
  }

  return out;
}

interface FormulaBarFunctionAutocompleteControllerOptions {
  formulaBar: FormulaBarView;
  /**
   * Element used as the positioning context (should be `position: relative`).
   */
  anchor: HTMLElement;
  maxItems?: number;
}

class FormulaBarFunctionAutocompleteController {
  readonly #formulaBar: FormulaBarView;
  readonly #textarea: HTMLTextAreaElement;
  readonly #maxItems: number;

  readonly #dropdownEl: HTMLDivElement;
  readonly #listboxId: string;
  readonly #optionIdPrefix: string;
  #itemEls: HTMLButtonElement[] = [];

  #context: CompletionContext | null = null;
  #suggestions: FunctionSuggestion[] = [];
  #selectedIndex = 0;
  #activeDescendantId: string | null = null;
  #isComposing = false;
  #signatureCatalogRefreshPromise: Promise<void> | null = null;

  readonly #unsubscribe: Array<() => void> = [];

  #scheduleSignatureCatalogRefresh(): void {
    if (this.#signatureCatalogRefreshPromise) return;

    this.#signatureCatalogRefreshPromise = preloadFunctionSignatureCatalog()
      .then(() => {
        if (!isFunctionSignatureCatalogReady()) return;
        // Don't auto-open the dropdown; only refresh if it's already visible.
        if (!this.isOpen()) return;
        this.update();
      })
      .finally(() => {
        this.#signatureCatalogRefreshPromise = null;
      });
  }

  constructor(opts: FormulaBarFunctionAutocompleteControllerOptions) {
    this.#formulaBar = opts.formulaBar;
    this.#textarea = opts.formulaBar.textarea;
    this.#maxItems = Math.max(1, Math.min(50, opts.maxItems ?? 12));

    AUTOCOMPLETE_INSTANCE_ID += 1;
    this.#listboxId = `formula-function-autocomplete-${AUTOCOMPLETE_INSTANCE_ID}`;
    this.#optionIdPrefix = `${this.#listboxId}-option`;

    const dropdown = document.createElement("div");
    dropdown.className = "formula-bar-function-autocomplete";
    dropdown.dataset.testid = "formula-function-autocomplete";
    dropdown.setAttribute("role", "listbox");
    dropdown.setAttribute("aria-label", "Function suggestions");
    dropdown.id = this.#listboxId;
    dropdown.hidden = true;
    opts.anchor.appendChild(dropdown);
    this.#dropdownEl = dropdown;

    // Keep the textarea focused while navigating the listbox, using the
    // active-descendant pattern for screen readers.
    this.#textarea.setAttribute("aria-haspopup", "listbox");
    this.#textarea.setAttribute("aria-controls", this.#listboxId);
    this.#textarea.setAttribute("aria-expanded", "false");
    this.#textarea.setAttribute("aria-autocomplete", "list");

    const updateNow = () => this.update();
    const onBlur = () => this.close();
    const onCompositionStart = () => {
      this.#isComposing = true;
      this.close();
    };
    const onCompositionEnd = () => {
      this.#isComposing = false;
      // Ensure the dropdown can respond to the next keydown immediately (e.g. ArrowDown),
      // while still scheduling a microtask update to catch cases where the final composition
      // text is applied after this callback.
      this.update();
      queueMicrotask(() => this.update());
    };
    this.#textarea.addEventListener("input", updateNow);
    this.#textarea.addEventListener("click", updateNow);
    this.#textarea.addEventListener("keyup", updateNow);
    this.#textarea.addEventListener("focus", updateNow);
    this.#textarea.addEventListener("select", updateNow);
    this.#textarea.addEventListener("blur", onBlur);
    this.#textarea.addEventListener("compositionstart", onCompositionStart);
    this.#textarea.addEventListener("compositionend", onCompositionEnd);

    this.#unsubscribe.push(() => this.#textarea.removeEventListener("input", updateNow));
    this.#unsubscribe.push(() => this.#textarea.removeEventListener("click", updateNow));
    this.#unsubscribe.push(() => this.#textarea.removeEventListener("keyup", updateNow));
    this.#unsubscribe.push(() => this.#textarea.removeEventListener("focus", updateNow));
    this.#unsubscribe.push(() => this.#textarea.removeEventListener("select", updateNow));
    this.#unsubscribe.push(() => this.#textarea.removeEventListener("blur", onBlur));
    this.#unsubscribe.push(() => this.#textarea.removeEventListener("compositionstart", onCompositionStart));
    this.#unsubscribe.push(() => this.#textarea.removeEventListener("compositionend", onCompositionEnd));
  }

  destroy(): void {
    for (const stop of this.#unsubscribe.splice(0)) stop();
    this.close();
    // Clean up ARIA attributes we own (FormulaBarView does not manage these).
    this.#textarea.removeAttribute("aria-haspopup");
    this.#textarea.removeAttribute("aria-controls");
    this.#textarea.removeAttribute("aria-expanded");
    this.#textarea.removeAttribute("aria-autocomplete");
    this.#dropdownEl.remove();
  }

  isOpen(): boolean {
    return !this.#dropdownEl.hidden;
  }

  close(): void {
    // Always clear state/ARIA even if the dropdown is already hidden (defensive:
    // DOM state may be manipulated externally and we don't want stale activedescendant ids).
    this.#formulaBar.root.classList.remove("formula-bar--function-autocomplete-open");
    this.#context = null;
    this.#suggestions = [];
    this.#selectedIndex = 0;
    this.#activeDescendantId = null;
    this.#textarea.removeAttribute("aria-activedescendant");
    this.#textarea.setAttribute("aria-expanded", "false");

    if (this.#dropdownEl.hidden) return;
    this.#dropdownEl.hidden = true;
    this.#dropdownEl.textContent = "";
    this.#itemEls = [];
  }

  update(): void {
    if (this.#isComposing) {
      this.close();
      return;
    }
    if (!this.#formulaBar.model.isEditing) {
      this.close();
      return;
    }

    const start = this.#textarea.selectionStart ?? this.#textarea.value.length;
    const end = this.#textarea.selectionEnd ?? this.#textarea.value.length;
    if (start !== end) {
      this.close();
      return;
    }

    // If the caret is inside a reference token (A1 ref, named range, structured ref),
    // suppress function autocomplete so Escape/Tab keep their reference-editing semantics.
    //
    // Note: FormulaBarView registers its own input/selection listeners *before*
    // constructing this controller, so `model.activeReferenceIndex()` is up-to-date
    // by the time we handle the same DOM event.
    if (this.#formulaBar.model.activeReferenceIndex() != null) {
      this.close();
      return;
    }

    const input = this.#textarea.value;
    const localeId = this.#formulaBar.currentLocaleId();
    const ctx = findCompletionContext(input, start, localeId);
    if (!ctx) {
      this.close();
      return;
    }

    const suggestions = buildSuggestions({
      prefixUpper: ctx.matchPrefixUpper,
      limit: this.#maxItems,
      localeId,
      // `_xlfn.` prefix is primarily used with canonical English names; avoid localizing completions in this mode.
      allowLocalized: ctx.qualifier.length === 0,
    });
    if (suggestions.length === 0) {
      this.close();
      return;
    }
    // If we're showing placeholder signatures, lazily load the function catalog metadata and
    // refresh the dropdown when it becomes available.
    if (!isFunctionSignatureCatalogReady() && suggestions.some((s) => s.signature === "(…)")) {
      this.#scheduleSignatureCatalogRefresh();
    }

    // Preserve selection when possible.
    const prevSelectedName = this.#suggestions[this.#selectedIndex]?.name ?? null;
    let nextIndex = 0;
    if (prevSelectedName) {
      const found = suggestions.findIndex((s) => s.name === prevSelectedName);
      if (found >= 0) nextIndex = found;
    }

    this.#context = ctx;
    this.#suggestions = suggestions;
    this.#selectedIndex = Math.min(nextIndex, suggestions.length - 1);
    this.#render();
  }

  handleKeyDown(e: KeyboardEvent): boolean {
    if (!this.isOpen()) return false;

    if (e.key === "ArrowDown") {
      e.preventDefault();
      this.#selectedIndex = Math.min(this.#selectedIndex + 1, this.#suggestions.length - 1);
      this.#syncSelection();
      return true;
    }

    if (e.key === "ArrowUp") {
      e.preventDefault();
      this.#selectedIndex = Math.max(this.#selectedIndex - 1, 0);
      this.#syncSelection();
      return true;
    }

    if (e.key === "PageDown") {
      e.preventDefault();
      this.#selectedIndex = Math.min(this.#selectedIndex + 5, this.#suggestions.length - 1);
      this.#syncSelection();
      return true;
    }

    if (e.key === "PageUp") {
      e.preventDefault();
      this.#selectedIndex = Math.max(this.#selectedIndex - 5, 0);
      this.#syncSelection();
      return true;
    }

    if (e.key === "Home") {
      e.preventDefault();
      this.#selectedIndex = 0;
      this.#syncSelection();
      return true;
    }

    if (e.key === "End") {
      e.preventDefault();
      this.#selectedIndex = Math.max(0, this.#suggestions.length - 1);
      this.#syncSelection();
      return true;
    }

    if (e.key === "Escape") {
      e.preventDefault();
      this.close();
      return true;
    }

    // Excel/editor-style commit character: typing `(` completes the selected function name.
    if (e.key === "(") {
      e.preventDefault();
      this.acceptSelected();
      return true;
    }

    // Match typical editor UX:
    // - Tab accepts the selected item (Shift+Tab should keep its usual meaning in the formula bar)
    // - Enter accepts (Shift+Enter remains available for formula-bar commit/navigation semantics)
    if ((e.key === "Tab" && !e.shiftKey) || (e.key === "Enter" && !e.altKey && !e.shiftKey)) {
      e.preventDefault();
      this.acceptSelected();
      return true;
    }

    return false;
  }

  acceptSelected(): void {
    const ctx = this.#context;
    const selected = this.#suggestions[this.#selectedIndex] ?? null;
    if (!ctx || !selected) {
      this.close();
      return;
    }

    const input = this.#textarea.value;
    // Preserve the user-typed casing for the function name portion while keeping
    // any `_xlfn.` qualifier prefix intact (Excel compatibility).
    const typedNamePrefix = ctx.typedPrefix.slice(ctx.qualifier.length);
    const casedName = preserveTypedCasing(typedNamePrefix, selected.name);
    const insertedName = `${ctx.qualifier}${casedName}`;
    // Avoid duplicating the opening paren if the user already has one in the text (e.g. editing `=VLO()`).
    const hasParen = input[ctx.replaceEnd] === "(";
    const inserted = hasParen ? insertedName : `${insertedName}(`;

    const nextText = input.slice(0, ctx.replaceStart) + inserted + input.slice(ctx.replaceEnd);
    // Always place the caret inside the parens (after `(`). When `(` already exists,
    // step over it after inserting the function name.
    const nextCursor = ctx.replaceStart + inserted.length + (hasParen ? 1 : 0);

    this.#textarea.value = nextText;
    this.#textarea.setSelectionRange(nextCursor, nextCursor);
    // Keep editing inside the formula bar.
    try {
      this.#textarea.focus({ preventScroll: true });
    } catch {
      this.#textarea.focus();
    }

    // Notify FormulaBarView listeners (model sync + highlight render).
    this.#textarea.dispatchEvent(new Event("input"));
    this.close();
  }

  #render(): void {
    this.#dropdownEl.hidden = false;
    this.#dropdownEl.textContent = "";
    this.#itemEls = [];
    this.#activeDescendantId = null;
    this.#textarea.setAttribute("aria-expanded", "true");
    this.#formulaBar.root.classList.add("formula-bar--function-autocomplete-open");

    for (let i = 0; i < this.#suggestions.length; i += 1) {
      const item = this.#suggestions[i]!;
      const button = document.createElement("button");
      button.type = "button";
      button.className = "formula-bar-function-autocomplete-item";
      // Keep focus on the textarea; options are managed via aria-activedescendant.
      button.tabIndex = -1;
      button.setAttribute("role", "option");
      button.dataset.testid = "formula-function-autocomplete-item";
      button.dataset.name = item.name;
      button.setAttribute("aria-selected", i === this.#selectedIndex ? "true" : "false");
      const id = `${this.#optionIdPrefix}-${i}`;
      button.id = id;
      if (i === this.#selectedIndex) {
        this.#activeDescendantId = id;
      }

      // Prevent the mousedown from stealing focus from the textarea.
      button.addEventListener("mousedown", (e) => e.preventDefault());
      button.addEventListener("mouseenter", () => {
        this.#selectedIndex = i;
        this.#syncSelection();
      });
      button.addEventListener("click", () => {
        this.#selectedIndex = i;
        this.acceptSelected();
      });

      const name = document.createElement("div");
      name.className = "formula-bar-function-autocomplete-name";
      const ctx = this.#context;
      const typedName = ctx ? ctx.typedPrefix.slice(ctx.qualifier.length) : "";
      const matchLen = Math.max(0, Math.min(item.name.length, typedName.length));
      if (matchLen > 0) {
        const match = document.createElement("span");
        match.className = "formula-bar-function-autocomplete-match";
        match.textContent = item.name.slice(0, matchLen);
        name.appendChild(match);
        name.appendChild(document.createTextNode(item.name.slice(matchLen)));
      } else {
        name.textContent = item.name;
      }

      const sig = document.createElement("div");
      sig.className = "formula-bar-function-autocomplete-signature";
      sig.textContent = item.signature;

      button.appendChild(name);
      button.appendChild(sig);

      this.#dropdownEl.appendChild(button);
      this.#itemEls.push(button);
    }

    this.#syncSelection();
  }

  #syncSelection(): void {
    this.#activeDescendantId = null;
    for (let i = 0; i < this.#itemEls.length; i += 1) {
      const el = this.#itemEls[i]!;
      const selected = i === this.#selectedIndex;
      el.setAttribute("aria-selected", selected ? "true" : "false");
      if (selected) {
        this.#activeDescendantId = el.id || null;
        try {
          el.scrollIntoView({ block: "nearest" });
        } catch {
          // jsdom doesn't implement layout/scroll; ignore.
        }
      }
    }

    if (this.#activeDescendantId) {
      this.#textarea.setAttribute("aria-activedescendant", this.#activeDescendantId);
    } else {
      this.#textarea.removeAttribute("aria-activedescendant");
    }
  }
}

type NameBoxDropdownItem = {
  /**
   * Stable key used for ARIA ids + testing.
   */
  key: string;
  kind: NameBoxDropdownItemKind;
  /**
   * User-facing label shown in the dropdown and written into the name box on selection.
   */
  label: string;
  /**
   * Reference string passed to `onGoTo`.
   *
   * This may differ from `label` (e.g. tables use `Table1[#All]`).
   */
  reference: string;
  /**
   * Optional secondary text (e.g. the resolved A1 range).
   */
  description?: string;
};

interface NameBoxDropdownProvider {
  getItems(): NameBoxDropdownItem[];
}

type FormulaBarViewOptions = FormulaBarViewToolingOptions & {
  nameBoxDropdownProvider?: NameBoxDropdownProvider;
};

interface FormulaBarViewCallbacks {
  onBeginEdit?: (activeCellAddress: string) => void;
  onCommit: (text: string, commit: FormulaBarCommit) => void;
  onCancel?: () => void;
  /**
   * Navigate the active selection to the provided A1/name/table reference.
   *
   * Return `true` only when navigation actually occurred. Return `false` when the
   * reference could not be parsed/resolved so the view can show "invalid reference"
   * feedback (Excel-style name box behavior).
   */
  onGoTo?: (reference: string) => boolean;
  onOpenNameBoxMenu?: () => void | Promise<void>;
  getNameBoxMenuItems?: () => NameBoxMenuItem[];
  onHoverRange?: (range: RangeAddress | null) => void;
  /**
   * Like `onHoverRange`, but includes the original reference text (e.g. `Sheet2!A1:B2`)
   * so consumers can display a label and/or enforce sheet-qualified preview behavior.
   */
  onHoverRangeWithText?: (range: RangeAddress | null, refText: string | null) => void;
  onReferenceHighlights?: (
    highlights: Array<{ range: FormulaReferenceRange; color: string; text: string; index: number; active?: boolean }>
  ) => void;
  onFixFormulaErrorWithAi?: (info: FixFormulaErrorWithAiInfo) => void;
}

let errorPanelIdCounter = 0;
function nextErrorPanelId(): string {
  errorPanelIdCounter += 1;
  return `formula-bar-error-panel-${errorPanelIdCounter}`;
}

let nameBoxErrorIdCounter = 0;
function nextNameBoxErrorId(): string {
  nameBoxErrorIdCounter += 1;
  return `formula-bar-name-box-error-${nameBoxErrorIdCounter}`;
}

let nameBoxListIdCounter = 0;
function nextNameBoxListId(): string {
  nameBoxListIdCounter += 1;
  return `formula-name-box-list-${nameBoxListIdCounter}`;
}

let functionPickerListIdCounter = 0;
function nextFunctionPickerListId(): string {
  functionPickerListIdCounter += 1;
  return `formula-function-picker-list-${functionPickerListIdCounter}`;
}

type FormulaBarCommitReason = "enter" | "tab" | "command";

interface FormulaBarCommit {
  reason: FormulaBarCommitReason;
  /**
   * Shift modifier for enter/tab (Shift+Enter moves up, Shift+Tab moves left).
   */
  shift: boolean;
}

type FormulaBarViewToolingOptions = {
  /**
   * Returns the current WASM engine instance (may be null while the worker/WASM is still loading).
   *
   * When present, the formula bar will use `lexFormulaPartial` / `parseFormulaPartial` for syntax
   * highlighting, function parameter hints, and syntax error spans.
   */
  getWasmEngine?: () => EngineClient | null;
  /**
   * Formula locale id (e.g. "en-US", "de-DE"). Defaults to `document.documentElement.lang` and then "en-US".
   */
  getLocaleId?: () => string;
  referenceStyle?: NonNullable<FormulaParseOptions["referenceStyle"]>;
};

export class FormulaBarView {
  readonly model = new FormulaBarModel();

  readonly root: HTMLElement;
  readonly textarea: HTMLTextAreaElement;

  #destroyed = false;
  #domAbort = new AbortController();

  #readOnly = false;
  #isComposing = false;
  #isFunctionPickerComposing = false;
  #isNameBoxComposing = false;

  #scheduledRender:
    | { id: number; kind: "raf" }
    | { id: ReturnType<typeof setTimeout>; kind: "timeout" }
    | null = null;
  #pendingRender: { preserveTextareaValue: boolean } | null = null;
  #lastHighlightHtml: string | null = null;
  #lastHighlightDraft: string | null = null;
  #lastHighlightIsFormulaEditing = false;
  #lastHighlightHadGhost = false;
  #lastHighlightCursor: number | null = null;
  #lastHighlightGhost: string | null = null;
  #lastHighlightPreviewText: string | null = null;
  #lastActiveReferenceIndex: number | null = null;
  #lastHighlightSpans: ReturnType<FormulaBarModel["highlightedSpans"]> | null = null;
  #lastColoredReferences: ReturnType<FormulaBarModel["coloredReferences"]> | null = null;
  #lastHintIsFormulaEditing: boolean | null = null;
  #lastHintSyntaxKey: string | null = null;
  #lastHint: ReturnType<FormulaBarModel["functionHint"]> | null = null;
  #lastHintArgPreviewKey: string | null = null;
  #lastHintArgPreviewRhs: string | null = null;
  #lastErrorExplanation: ReturnType<FormulaBarModel["errorExplanation"]> | null = null;
  // Use a non-null sentinel so the first render always syncs the error panel state.
  #lastErrorExplanationAddress: string | null = "__init__";
  #lastRootIsEditingClass: boolean | null = null;
  #lastHintHasSyntaxErrorClass: boolean | null = null;
  #lastShowEditingActions: boolean | null = null;
  #lastErrorFixAiDisabled: boolean | null = null;
  #lastErrorShowRangesDisabled: boolean | null = null;
  #lastErrorShowRangesPressed: boolean | null = null;
  #lastErrorShowRangesText: string | null = null;
  #referenceElsByIndex: Array<HTMLElement[] | undefined> | null = null;
  #lastAdjustedHeightDraft: string | null = null;
  #lastAdjustedHeightIsEditing = false;
  #lastAdjustedHeightIsExpanded = false;

  #argumentPreviewProvider: ((expr: string) => unknown | Promise<unknown>) | null = null;
  #argumentPreviewKey: string | null = null;
  #argumentPreviewDisplayKey: string | null = null;
  #argumentPreviewDisplayExpr: string | null = null;
  #argumentPreviewValue: unknown | null = null;
  #argumentPreviewDisplayValue: string | null = null;
  #argumentPreviewPending = false;
  #argumentPreviewTimer: ReturnType<typeof setTimeout> | null = null;
  #argumentPreviewRequestId = 0;

  #functionAutocomplete: FormulaBarFunctionAutocompleteController;
  #nameBoxEl: HTMLDivElement;
  #nameBoxDropdownEl: HTMLButtonElement;
  #nameBoxDropdownPopupEl: HTMLDivElement;
  #nameBoxDropdownListEl: HTMLDivElement;
  #nameBoxDropdownProvider: NameBoxDropdownProvider | null = null;
  #isNameBoxDropdownOpen = false;
  #nameBoxDropdownOriginalAddressValue: string | null = null;
  #nameBoxDropdownAllItems: NameBoxDropdownItem[] = [];
  #nameBoxDropdownQuery = "";
  #nameBoxDropdownFilteredItems: NameBoxDropdownItem[] = [];
  #nameBoxDropdownOptionEls: HTMLElement[] = [];
  #nameBoxDropdownActiveIndex: number = -1;
  #nameBoxDropdownRecentKeys: string[] = [];
  #nameBoxDropdownPointerDownListener: ((e: PointerEvent) => void) | null = null;
  #nameBoxDropdownFocusInListener: ((e: FocusEvent) => void) | null = null;
  #nameBoxDropdownScrollListener: ((e: Event) => void) | null = null;
  #nameBoxDropdownResizeListener: (() => void) | null = null;
  #nameBoxDropdownBlurListener: (() => void) | null = null;
  #cancelButtonEl: HTMLButtonElement;
  #commitButtonEl: HTMLButtonElement;
  #fxButtonEl: HTMLButtonElement;
  #expandButtonEl: HTMLButtonElement;
  #addressEl: HTMLInputElement;
  #nameBoxErrorEl: HTMLDivElement;
  #highlightEl: HTMLElement;
  #hintEl: HTMLElement;
  #errorButton: HTMLButtonElement;
  #errorPanel: HTMLElement;
  #errorTitleEl: HTMLElement;
  #errorDescEl: HTMLElement;
  #errorSuggestionsEl: HTMLUListElement;
  #errorFixAiButton: HTMLButtonElement;
  #errorShowRangesButton: HTMLButtonElement;
  #errorCloseButton: HTMLButtonElement;
  #isErrorPanelOpen = false;
  #errorPanelReferenceHighlights: FormulaReferenceHighlight[] | null = null;
  #hoverOverride: RangeAddress | null = null;
  #hoverOverrideText: string | null = null;
  #lastEmittedHoverRange: { startRow: number; startCol: number; endRow: number; endCol: number } | null = null;
  #lastEmittedHoverText: string | null = null;
  #lastEmittedReferenceHighlightsMode: ReferenceHighlightMode = "none";
  #lastEmittedReferenceHighlightsColoredRefs: ReturnType<FormulaBarModel["coloredReferences"]> | null = null;
  #lastEmittedReferenceHighlightsActiveIndex: number | null = null;
  #lastEmittedReferenceHighlightsErrorPanel: FormulaReferenceHighlight[] | null = null;
  #selectedReferenceIndex: number | null = null;
  #mouseDownSelectedReferenceIndex: number | null = null;
  #nameBoxValue = "A1";
  #isExpanded = false;
  #callbacks: FormulaBarViewCallbacks;
  #tooling: FormulaBarViewToolingOptions | null = null;
  #toolingRequestId = 0;
  #toolingScheduled:
    | { id: number; kind: "raf" }
    | { id: ReturnType<typeof setTimeout>; kind: "timeout" }
    | null = null;
  #toolingAbort: AbortController | null = null;
  #toolingPending: {
    requestId: number;
    draft: string;
    cursor: number;
    localeId: string;
    referenceStyle: NonNullable<FormulaParseOptions["referenceStyle"]>;
  } | null = null;
  #toolingLexCache:
    | {
        draft: string;
        localeId: string;
        referenceStyle: NonNullable<FormulaParseOptions["referenceStyle"]>;
        lexResult: Awaited<ReturnType<EngineClient["lexFormulaPartial"]>>;
      }
    | null = null;
  #nameBoxMenu: ContextMenu | null = null;
  #restoreNameBoxFocusOnMenuClose = false;
  #nameBoxMenuEscapeListener: ((e: KeyboardEvent) => void) | null = null;

  #functionPickerEl: HTMLDivElement;
  #functionPickerInputEl: HTMLInputElement;
  #functionPickerListEl: HTMLUListElement;
  #functionPickerOpen = false;
  #functionPickerItems: FunctionPickerItem[] = [];
  #functionPickerItemEls: HTMLLIElement[] = [];
  #functionPickerSelectedIndex = 0;
  #functionPickerAnchorSelection: { start: number; end: number } | null = null;
  #functionPickerDocMouseDown = (e: MouseEvent) => this.#onFunctionPickerDocMouseDown(e);
  #nameBoxErrorId: string;
  #isNameBoxInvalid = false;

  constructor(root: HTMLElement, callbacks: FormulaBarViewCallbacks, opts: FormulaBarViewOptions = {}) {
    this.root = root;
    this.#callbacks = callbacks;
    this.#nameBoxDropdownProvider = opts.nameBoxDropdownProvider ?? null;
    this.#tooling =
      typeof opts.getWasmEngine === "function" || typeof opts.getLocaleId === "function" || opts.referenceStyle != null
        ? opts
        : null;

    root.classList.add("formula-bar");
    this.#isExpanded = loadFormulaBarExpandedState();
    root.classList.toggle("formula-bar--expanded", this.#isExpanded);

    const row = document.createElement("div");
    row.className = "formula-bar-row";

    const address = document.createElement("input");
    address.className = "formula-bar-address";
    address.dataset.testid = "formula-address";
    address.setAttribute("aria-label", "Name box");
    // When the listbox dropdown provider is present, the name box input participates in a listbox
    // (type-to-filter + aria-activedescendant). Otherwise, the name box supports a menu-style
    // affordance (Alt+Down / F4) via the ContextMenu fallback.
    address.setAttribute(
      "aria-haspopup",
      this.#nameBoxDropdownProvider ? "listbox" : this.#callbacks.getNameBoxMenuItems || this.#callbacks.onOpenNameBoxMenu ? "menu" : "false"
    );
    address.setAttribute("aria-expanded", "false");
    address.autocomplete = "off";
    address.spellcheck = false;
    address.value = "A1";

    const nameBox = document.createElement("div");
    nameBox.className = "formula-bar-name-box";

    const nameBoxWrapper = document.createElement("div");
    nameBoxWrapper.className = "formula-bar-name-box-wrapper";

    const nameBoxDropdown = document.createElement("button");
    nameBoxDropdown.className = "formula-bar-name-box-dropdown";
    nameBoxDropdown.dataset.testid = "name-box-dropdown";
    nameBoxDropdown.type = "button";
    nameBoxDropdown.textContent = "▾";
    nameBoxDropdown.title = "Name box menu";
    nameBoxDropdown.setAttribute("aria-label", "Open name box menu");
    nameBoxDropdown.setAttribute("aria-haspopup", this.#nameBoxDropdownProvider ? "listbox" : "menu");
    nameBoxDropdown.setAttribute("aria-expanded", "false");

    nameBox.appendChild(address);
    nameBox.appendChild(nameBoxDropdown);

    const nameBoxError = document.createElement("div");
    nameBoxError.className = "formula-bar-name-box-error";
    nameBoxError.hidden = true;
    nameBoxError.textContent = "Invalid reference";
    nameBoxError.setAttribute("role", "tooltip");
    const nameBoxErrorId = nextNameBoxErrorId();
    nameBoxError.id = nameBoxErrorId;

    nameBoxWrapper.appendChild(nameBox);
    nameBoxWrapper.appendChild(nameBoxError);

    const actions = document.createElement("div");
    actions.className = "formula-bar-actions";

    const cancelButton = document.createElement("button");
    cancelButton.className = "formula-bar-action-button formula-bar-action-button--cancel";
    cancelButton.type = "button";
    cancelButton.textContent = "✕";
    cancelButton.title = "Cancel (Esc)";
    cancelButton.setAttribute("aria-label", "Cancel edit");

    const commitButton = document.createElement("button");
    commitButton.className = "formula-bar-action-button formula-bar-action-button--commit";
    commitButton.type = "button";
    commitButton.textContent = "✓";
    commitButton.title = "Enter (↵)";
    commitButton.setAttribute("aria-label", "Commit edit");

    const fxButton = document.createElement("button");
    fxButton.className = "formula-bar-action-button formula-bar-action-button--fx";
    fxButton.type = "button";
    fxButton.textContent = "fx";
    fxButton.title = "Insert function";
    fxButton.setAttribute("aria-label", "Insert function");
    fxButton.setAttribute("aria-haspopup", "dialog");
    fxButton.setAttribute("aria-expanded", "false");
    fxButton.dataset.testid = "formula-fx-button";

    actions.appendChild(cancelButton);
    actions.appendChild(commitButton);
    actions.appendChild(fxButton);

    const editor = document.createElement("div");
    editor.className = "formula-bar-editor";

    const highlight = document.createElement("pre");
    highlight.className = "formula-bar-highlight";
    highlight.dataset.testid = "formula-highlight";
    highlight.setAttribute("aria-hidden", "true");

    const textarea = document.createElement("textarea");
    textarea.className = "formula-bar-input";
    textarea.dataset.testid = "formula-input";
    textarea.setAttribute("aria-label", "Formula bar");
    textarea.spellcheck = false;
    textarea.autocapitalize = "off";
    textarea.autocomplete = "off";
    textarea.wrap = "off";
    textarea.rows = 1;

    editor.appendChild(highlight);
    editor.appendChild(textarea);

    const expandButton = document.createElement("button");
    expandButton.className = "formula-bar-expand-button";
    expandButton.type = "button";
    expandButton.dataset.testid = "formula-expand-button";
    // Label/text synced after elements are assigned.
    expandButton.textContent = "▾";
    expandButton.title = "Expand formula bar";
    expandButton.setAttribute("aria-label", "Expand formula bar");
    expandButton.setAttribute("aria-pressed", "false");

    const errorButton = document.createElement("button");
    errorButton.className = "formula-bar-error-button";
    errorButton.type = "button";
    errorButton.textContent = "!";
    errorButton.title = "Show error details";
    errorButton.setAttribute("aria-label", "Show formula error");
    errorButton.setAttribute("aria-expanded", "false");
    errorButton.setAttribute("aria-haspopup", "dialog");
    errorButton.dataset.testid = "formula-error-button";
    errorButton.hidden = true;
    errorButton.disabled = true;

    const errorPanel = document.createElement("div");
    errorPanel.className = "formula-bar-error-panel";
    errorPanel.dataset.testid = "formula-error-panel";
    errorPanel.hidden = true;

    const errorPanelId = nextErrorPanelId();
    errorPanel.id = errorPanelId;
    errorButton.setAttribute("aria-controls", errorPanelId);
    errorPanel.setAttribute("role", "dialog");
    errorPanel.setAttribute("aria-modal", "false");

    const errorHeader = document.createElement("div");
    errorHeader.className = "formula-bar-error-header";

    const errorTitle = document.createElement("div");
    errorTitle.className = "formula-bar-error-title";
    errorTitle.id = `${errorPanelId}-title`;

    const errorCloseButton = document.createElement("button");
    errorCloseButton.className = "formula-bar-error-close";
    errorCloseButton.type = "button";
    errorCloseButton.textContent = "✕";
    errorCloseButton.title = "Dismiss (Esc)";
    errorCloseButton.setAttribute("aria-label", "Dismiss formula error details");
    errorCloseButton.dataset.testid = "formula-error-close";

    errorHeader.appendChild(errorTitle);
    errorHeader.appendChild(errorCloseButton);

    const errorDesc = document.createElement("div");
    errorDesc.className = "formula-bar-error-desc";
    errorDesc.id = `${errorPanelId}-desc`;

    const errorSuggestions = document.createElement("ul");
    errorSuggestions.className = "formula-bar-error-suggestions";

    const errorActions = document.createElement("div");
    errorActions.className = "formula-bar-error-actions";

    const errorFixAiButton = document.createElement("button");
    errorFixAiButton.className = "formula-bar-error-action formula-bar-error-action--primary";
    errorFixAiButton.type = "button";
    errorFixAiButton.textContent = "Fix with AI";
    errorFixAiButton.dataset.testid = "formula-error-fix-ai";

    const errorShowRangesButton = document.createElement("button");
    errorShowRangesButton.className = "formula-bar-error-action formula-bar-error-action--secondary";
    errorShowRangesButton.type = "button";
    errorShowRangesButton.textContent = "Show referenced ranges";
    errorShowRangesButton.setAttribute("aria-pressed", "false");
    errorShowRangesButton.dataset.testid = "formula-error-show-ranges";

    errorActions.appendChild(errorFixAiButton);
    errorActions.appendChild(errorShowRangesButton);

    errorPanel.appendChild(errorHeader);
    errorPanel.appendChild(errorDesc);
    errorPanel.appendChild(errorSuggestions);
    errorPanel.appendChild(errorActions);
    errorPanel.setAttribute("aria-labelledby", errorTitle.id);
    errorPanel.setAttribute("aria-describedby", errorDesc.id);

    row.appendChild(nameBoxWrapper);
    row.appendChild(actions);
    row.appendChild(editor);
    row.appendChild(expandButton);
    row.appendChild(errorButton);

    const hint = document.createElement("div");
    hint.className = "formula-bar-hint";
    hint.dataset.testid = "formula-hint";

    const functionPicker = document.createElement("div");
    functionPicker.className = "formula-function-picker";
    functionPicker.dataset.testid = "formula-function-picker";
    functionPicker.hidden = true;
    functionPicker.setAttribute("role", "dialog");
    functionPicker.setAttribute("aria-label", "Insert function");

    const functionPickerPanel = document.createElement("div");
    functionPickerPanel.className = "command-palette";
    functionPickerPanel.dataset.testid = "formula-function-picker-panel";

    const functionPickerInput = document.createElement("input");
    functionPickerInput.className = "command-palette__input";
    functionPickerInput.dataset.testid = "formula-function-picker-input";
    functionPickerInput.placeholder = "Search functions";
    functionPickerInput.setAttribute("role", "combobox");
    functionPickerInput.setAttribute("aria-autocomplete", "list");
    functionPickerInput.setAttribute("aria-expanded", "false");
    functionPickerInput.setAttribute("aria-label", "Search functions");
    // Avoid browser spellcheck/autofill UI interfering with keyboard navigation + e2e tests.
    functionPickerInput.spellcheck = false;
    functionPickerInput.autocapitalize = "off";
    functionPickerInput.autocomplete = "off";

    const functionPickerList = document.createElement("ul");
    functionPickerList.className = "command-palette__list";
    functionPickerList.dataset.testid = "formula-function-picker-list";
    functionPickerList.id = nextFunctionPickerListId();
    functionPickerList.setAttribute("role", "listbox");
    // Ensure there is at least one tabbable element besides the input so Tab doesn't escape.
    functionPickerList.tabIndex = 0;
    functionPickerInput.setAttribute("aria-controls", functionPickerList.id);
    functionPickerInput.setAttribute("aria-haspopup", "listbox");

    functionPickerPanel.appendChild(functionPickerInput);
    functionPickerPanel.appendChild(functionPickerList);
    functionPicker.appendChild(functionPickerPanel);

    const nameBoxDropdownPopup = document.createElement("div");
    nameBoxDropdownPopup.className = "formula-bar-name-box-popup";
    nameBoxDropdownPopup.hidden = true;
    nameBoxDropdownPopup.dataset.testid = "formula-name-box-popup";

    const nameBoxDropdownList = document.createElement("div");
    nameBoxDropdownList.className = "formula-bar-name-box-popup-list";
    nameBoxDropdownList.dataset.testid = "formula-name-box-list";
    nameBoxDropdownList.id = nextNameBoxListId();
    nameBoxDropdownList.setAttribute("role", "listbox");
    nameBoxDropdownList.setAttribute("aria-label", "Name box menu");
    nameBoxDropdownPopup.appendChild(nameBoxDropdownList);

    root.appendChild(row);
    root.appendChild(hint);
    root.appendChild(errorPanel);
    root.appendChild(functionPicker);
    root.appendChild(nameBoxDropdownPopup);

    this.textarea = textarea;
    this.#nameBoxEl = nameBox;
    this.#nameBoxDropdownEl = nameBoxDropdown;
    this.#nameBoxDropdownPopupEl = nameBoxDropdownPopup;
    this.#nameBoxDropdownListEl = nameBoxDropdownList;
    this.#cancelButtonEl = cancelButton;
    this.#commitButtonEl = commitButton;
    this.#fxButtonEl = fxButton;
    this.#expandButtonEl = expandButton;
    this.#addressEl = address;
    this.#nameBoxErrorEl = nameBoxError;
    this.#highlightEl = highlight;
    this.#hintEl = hint;
    this.#errorButton = errorButton;
    this.#errorPanel = errorPanel;
    this.#errorTitleEl = errorTitle;
    this.#errorDescEl = errorDesc;
    this.#errorSuggestionsEl = errorSuggestions;
    this.#errorFixAiButton = errorFixAiButton;
    this.#errorShowRangesButton = errorShowRangesButton;
    this.#errorCloseButton = errorCloseButton;
    this.#functionPickerEl = functionPicker;
    this.#functionPickerInputEl = functionPickerInput;
    this.#functionPickerListEl = functionPickerList;
    this.#nameBoxErrorId = nameBoxErrorId;

    address.addEventListener(
      "focus",
      () => {
        address.select();
      },
      { signal: this.#domAbort.signal },
    );

    address.addEventListener(
      "input",
      () => {
        if (!this.#isNameBoxInvalid) return;
        this.#clearNameBoxError();
      },
      { signal: this.#domAbort.signal },
    );

    address.addEventListener(
      "blur",
      () => {
        // Don't leave the Name Box in an "invalid" visual state if the user abandons the entry.
        if (!this.#isNameBoxInvalid) return;
        this.#clearNameBoxError();
        address.value = this.#nameBoxValue;
      },
      { signal: this.#domAbort.signal },
    );

    nameBoxDropdown.addEventListener(
      "click",
      () => {
      if (this.#nameBoxDropdownProvider) {
        if (this.#isNameBoxDropdownOpen) {
          this.#closeNameBoxDropdown({ restoreAddress: true, reason: "toggle" });
          try {
            this.#addressEl.focus({ preventScroll: true });
          } catch {
            this.#addressEl.focus();
          }
          this.#addressEl.select();
        } else {
          this.#openNameBoxDropdown();
        }
        return;
      }

      if (this.#callbacks.getNameBoxMenuItems) {
        this.#toggleNameBoxMenu();
        return;
      }
      if (this.#callbacks.onOpenNameBoxMenu) {
        Promise.resolve(this.#callbacks.onOpenNameBoxMenu())
          .catch((err) => {
            console.error("Failed to open name box menu:", err);
          });
        return;
      }

      // Fallback affordance: focus the address input so keyboard "Go To" still feels natural.
      address.focus();
      },
      { signal: this.#domAbort.signal },
    );

    address.addEventListener(
      "keydown",
      (e) => {
      if (
        (this.#isNameBoxComposing || e.isComposing) &&
        (e.key === "Enter" || e.key === "Escape" || e.key === "ArrowDown" || e.key === "ArrowUp" || e.key === "F4")
      ) {
        return;
      }

      const wantsMenuKey =
        (e.key === "ArrowDown" && e.altKey && !e.ctrlKey && !e.metaKey) ||
        (e.key === "F4" && !e.altKey && !e.ctrlKey && !e.metaKey);

      if (this.#nameBoxDropdownProvider) {
        if (this.#isNameBoxDropdownOpen) {
          if (wantsMenuKey) {
            e.preventDefault();
            this.#closeNameBoxDropdown({ restoreAddress: true, reason: "toggle" });
            return;
          }
          if (e.key === "ArrowDown") {
            e.preventDefault();
            this.#moveNameBoxDropdownSelection(1);
            return;
          }
          if (e.key === "ArrowUp") {
            e.preventDefault();
            this.#moveNameBoxDropdownSelection(-1);
            return;
          }
          if (e.key === "Enter") {
            e.preventDefault();
            const active = this.#nameBoxDropdownFilteredItems[this.#nameBoxDropdownActiveIndex] ?? null;
            if (active) {
              this.#selectNameBoxDropdownItem(active);
              return;
            }
            // Fall back to the standard Go To behavior if filtering produced no matches.
            this.#closeNameBoxDropdown({ restoreAddress: false, reason: "commit" });
            const ref = address.value.trim();
            const handler = this.#callbacks.onGoTo;
            if (!handler) {
              address.blur();
              return;
            }

            let ok = false;
            try {
              ok = handler(ref) === true;
            } catch {
              ok = false;
            }

            if (!ok) {
              this.#setNameBoxError("Invalid reference");
              try {
                address.focus({ preventScroll: true });
              } catch {
                address.focus();
              }
              address.select();
              return;
            }

            this.#clearNameBoxError();
            // Blur after navigating so follow-up renders can update the value. Since navigation
            // happens synchronously inside `onGoTo` (SpreadsheetApp immediately calls `setActiveCell`
            // while the input is still focused), also apply the latest `#nameBoxValue` so the Name Box
            // reflects the new selection immediately after focus leaves the input.
            address.blur();
            address.value = this.#nameBoxValue;
            return;
          }
          if (e.key === "Escape") {
            e.preventDefault();
            this.#closeNameBoxDropdown({ restoreAddress: true, reason: "escape" });
            // If the user was in an invalid-reference state before opening the dropdown,
            // treat Escape as "cancel the invalid entry" (Excel-like behavior).
            if (this.#isNameBoxInvalid) {
              this.#clearNameBoxError();
              address.value = this.#nameBoxValue;
            }
            return;
          }
          if (e.key === "Tab") {
            // Allow standard focus traversal (Tab/Shift+Tab) while dismissing the dropdown.
            this.#closeNameBoxDropdown({ restoreAddress: true, reason: "outside" });
            return;
          }
        }

        if (wantsMenuKey) {
          e.preventDefault();
          this.#openNameBoxDropdown();
          return;
        }
      } else if (wantsMenuKey) {
        // Excel-style name box dropdown affordance.
        e.preventDefault();
        if (this.#callbacks.getNameBoxMenuItems) {
          this.#toggleNameBoxMenu();
        } else if (this.#callbacks.onOpenNameBoxMenu) {
          Promise.resolve(this.#callbacks.onOpenNameBoxMenu()).catch((err) => {
            console.error("Failed to open name box menu:", err);
          });
        } else {
          this.#toggleNameBoxMenu();
        }
        return;
      }

      if (e.key === "Enter") {
        e.preventDefault();
        const ref = address.value.trim();
        const handler = this.#callbacks.onGoTo;
        if (!handler) {
          // Fallback: treat Enter as a no-op navigation and allow normal blur behavior.
          address.blur();
          return;
        }

        let ok = false;
        try {
          ok = handler(ref) === true;
        } catch {
          ok = false;
        }

        if (!ok) {
          this.#setNameBoxError("Invalid reference");
          // Keep focus in the input so the user can correct the reference.
          try {
            address.focus({ preventScroll: true });
          } catch {
            address.focus();
          }
          address.select();
          return;
        }

        this.#clearNameBoxError();
        // Blur after navigating so follow-up renders can update the value.
        address.blur();
        address.value = this.#nameBoxValue;
        return;
      }

      if (e.key === "Escape") {
        e.preventDefault();
        this.#clearNameBoxError();
        address.value = this.#nameBoxValue;
        address.blur();
      }
      },
      { signal: this.#domAbort.signal },
    );

    address.addEventListener(
      "compositionstart",
      () => {
        this.#isNameBoxComposing = true;
      },
      { signal: this.#domAbort.signal },
    );
    address.addEventListener(
      "compositionend",
      () => {
        this.#isNameBoxComposing = false;
      },
      { signal: this.#domAbort.signal },
    );
    address.addEventListener(
      "blur",
      () => {
        this.#isNameBoxComposing = false;
      },
      { signal: this.#domAbort.signal },
    );

    address.addEventListener(
      "input",
      () => {
        if (!this.#isNameBoxDropdownOpen) return;
        this.#updateNameBoxDropdownFilter(address.value);
      },
      { signal: this.#domAbort.signal },
    );

    textarea.addEventListener("focus", () => this.#beginEditFromFocus(), { signal: this.#domAbort.signal });
    textarea.addEventListener("input", () => this.#onInput(), { signal: this.#domAbort.signal });
    textarea.addEventListener("mousedown", (e) => this.#onTextareaMouseDown(e), { signal: this.#domAbort.signal });
    textarea.addEventListener("click", () => this.#onTextareaClick(), { signal: this.#domAbort.signal });
    textarea.addEventListener("keyup", () => this.#onSelectionChange(), { signal: this.#domAbort.signal });
    textarea.addEventListener("select", () => this.#onSelectionChange(), { signal: this.#domAbort.signal });
    textarea.addEventListener("scroll", () => this.#syncScroll(), { signal: this.#domAbort.signal });
    textarea.addEventListener("keydown", (e) => this.#onKeyDown(e), { signal: this.#domAbort.signal });
    textarea.addEventListener(
      "compositionstart",
      () => {
        this.#isComposing = true;
      },
      { signal: this.#domAbort.signal },
    );
    textarea.addEventListener(
      "compositionend",
      () => {
        this.#isComposing = false;
      },
      { signal: this.#domAbort.signal },
    );
    textarea.addEventListener(
      "blur",
      () => {
        this.#isComposing = false;
      },
      { signal: this.#domAbort.signal },
    );

    // Non-AI function autocomplete dropdown (Excel-like).
    // Mount after registering FormulaBarView's own listeners so focus/input updates keep the model in sync first.
    this.#functionAutocomplete = new FormulaBarFunctionAutocompleteController({ formulaBar: this, anchor: editor });

    // When not editing, allow hover previews using the highlighted spans.
    highlight.addEventListener("mousemove", (e) => this.#onHighlightHover(e), { signal: this.#domAbort.signal });
    highlight.addEventListener("mouseleave", () => this.#clearHoverOverride(), { signal: this.#domAbort.signal });
    highlight.addEventListener(
      "mousedown",
      (e) => {
        // Prevent selecting text in <pre> and instead focus the textarea.
        e.preventDefault();
        this.focus({ cursor: "end" });
      },
      { signal: this.#domAbort.signal },
    );

    errorButton.addEventListener(
      "click",
      () => {
        if (!this.root.classList.contains("formula-bar--has-error")) return;
        this.#setErrorPanelOpen(!this.#isErrorPanelOpen);
      },
      { signal: this.#domAbort.signal },
    );

    errorCloseButton.addEventListener("click", () => this.#setErrorPanelOpen(false, { restoreFocus: true }), {
      signal: this.#domAbort.signal,
    });
    errorPanel.addEventListener("keydown", (e) => this.#onErrorPanelKeyDown(e), { signal: this.#domAbort.signal });
    errorFixAiButton.addEventListener("click", () => this.#fixFormulaErrorWithAi(), { signal: this.#domAbort.signal });
    errorShowRangesButton.addEventListener("click", () => this.#toggleErrorReferenceHighlights(), { signal: this.#domAbort.signal });

    cancelButton.addEventListener("click", () => this.#cancel(), { signal: this.#domAbort.signal });
    commitButton.addEventListener("click", () => this.#commit({ reason: "command", shift: false }), { signal: this.#domAbort.signal });
    fxButton.addEventListener("click", () => this.#focusFx(), { signal: this.#domAbort.signal });
    fxButton.addEventListener(
      "mousedown",
      (e) => {
        // Preserve the caret/selection in the textarea when clicking the fx button.
        e.preventDefault();
      },
      { signal: this.#domAbort.signal },
    );

    expandButton.addEventListener("click", () => this.#toggleExpanded(), { signal: this.#domAbort.signal });
    expandButton.addEventListener(
      "mousedown",
      (e) => {
        // Preserve the caret/selection in the textarea when clicking the toggle button.
        e.preventDefault();
      },
      { signal: this.#domAbort.signal },
    );

    functionPickerInput.addEventListener("input", () => this.#onFunctionPickerInput(), { signal: this.#domAbort.signal });
    const pickerKeyDown = (e: KeyboardEvent) => this.#onFunctionPickerKeyDown(e);
    functionPickerInput.addEventListener("keydown", pickerKeyDown, { signal: this.#domAbort.signal });
    functionPickerList.addEventListener("keydown", pickerKeyDown, { signal: this.#domAbort.signal });
    functionPickerInput.addEventListener(
      "compositionstart",
      () => {
        this.#isFunctionPickerComposing = true;
      },
      { signal: this.#domAbort.signal },
    );
    functionPickerInput.addEventListener(
      "compositionend",
      () => {
        this.#isFunctionPickerComposing = false;
      },
      { signal: this.#domAbort.signal },
    );
    functionPickerInput.addEventListener(
      "blur",
      () => {
        this.#isFunctionPickerComposing = false;
      },
      { signal: this.#domAbort.signal },
    );

    this.#syncExpandedUi();

    // Initial render.
    this.model.setActiveCell({ address: "A1", input: "", value: "" });
    this.#render({ preserveTextareaValue: false });
  }

  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyed = true;

    // Cancel any scheduled work first so it can't re-attach listeners or update DOM after teardown.
    this.#cancelPendingRender();
    this.#cancelPendingTooling();
    this.#clearArgumentPreviewState();

    // Close transient UI surfaces (these also detach any global listeners they own).
    try {
      this.#closeFunctionPicker({ restoreFocus: false });
    } catch {
      // ignore
    }
    try {
      this.#closeNameBoxDropdown({ restoreAddress: true, reason: "toggle" });
    } catch {
      // ignore
    }
    try {
      this.#nameBoxMenu?.destroy();
    } catch {
      // ignore
    }
    this.#nameBoxMenu = null;
    this.#nameBoxMenuEscapeListener = null;

    // Remove autocomplete dropdown + its listeners.
    try {
      this.#functionAutocomplete.destroy();
    } catch {
      // ignore
    }

    // Detach all DOM event handlers registered with `#domAbort`.
    try {
      this.#domAbort.abort();
    } catch {
      // ignore
    }

    // If the formula bar host element is reused for another workbook, ensure it starts empty.
    // (FormulaBarView does not clear existing children on construction.)
    try {
      this.root.replaceChildren();
    } catch {
      // ignore
    }

    try {
      this.root.classList.remove(
        "formula-bar",
        "formula-bar--expanded",
        "formula-bar--function-autocomplete-open",
        "formula-bar--has-error",
        "formula-bar--read-only",
        "formula-bar--editing",
        "formula-bar--error-panel-open",
      );
    } catch {
      // ignore
    }
  }

  #toggleNameBoxMenu(): void {
    const menu = (this.#nameBoxMenu ??= new ContextMenu({
      testId: "name-box-menu",
      onClose: () => {
        this.#nameBoxDropdownEl.setAttribute("aria-expanded", "false");
        this.#addressEl.setAttribute("aria-expanded", "false");
        if (this.#nameBoxMenuEscapeListener) {
          window.removeEventListener("keydown", this.#nameBoxMenuEscapeListener, true);
          this.#nameBoxMenuEscapeListener = null;
        }

        if (this.#restoreNameBoxFocusOnMenuClose) {
          this.#restoreNameBoxFocusOnMenuClose = false;
          try {
            this.#addressEl.focus({ preventScroll: true });
          } catch {
            this.#addressEl.focus();
          }
          this.#addressEl.select();
        }
      },
    }));

    if (menu.isOpen()) {
      this.#restoreNameBoxFocusOnMenuClose = true;
      menu.close();
      return;
    }

    // Track Esc-driven closes so we can restore focus without stealing it on outside clicks.
    this.#restoreNameBoxFocusOnMenuClose = false;
    if (this.#nameBoxMenuEscapeListener) {
      window.removeEventListener("keydown", this.#nameBoxMenuEscapeListener, true);
      this.#nameBoxMenuEscapeListener = null;
    }
    this.#nameBoxMenuEscapeListener = (e: KeyboardEvent) => {
      const isEscape =
        e.key === "Escape" || e.key === "Esc" || e.code === "Escape" || (e as unknown as { keyCode?: number }).keyCode === 27;
      if (!isEscape) return;
      if (!menu.isOpen()) return;
      this.#restoreNameBoxFocusOnMenuClose = true;
    };
    window.addEventListener("keydown", this.#nameBoxMenuEscapeListener, { capture: true, signal: this.#domAbort.signal });

    const rawItems = this.#callbacks.getNameBoxMenuItems?.() ?? [];
    const items: ContextMenuItem[] = [];

    for (const item of rawItems) {
      const label = String(item?.label ?? "").trim();
      if (!label) continue;
      const enabled = item.enabled ?? true;
      const reference = typeof item.reference === "string" ? item.reference.trim() : item.reference ?? null;

      items.push({
        type: "item",
        label,
        enabled,
        onSelect: () => {
          // Selecting any entry is a corrective action; clear prior invalid input feedback.
          this.#clearNameBoxError();
          if (reference) {
            const handler = this.#callbacks.onGoTo;
            if (!handler) return;

            let ok = false;
            try {
              ok = handler(reference) === true;
            } catch {
              ok = false;
            }

            if (!ok) {
              this.#setNameBoxError("Invalid reference");
              try {
                this.#addressEl.focus({ preventScroll: true });
              } catch {
                this.#addressEl.focus();
              }
              this.#addressEl.select();
              return;
            }

            // Successful navigation should behave like pressing Enter in the name box:
            // clear any error feedback and allow follow-up renders to update the displayed address.
            this.#clearNameBoxError();
            this.#addressEl.blur();
            this.#addressEl.value = this.#nameBoxValue;
            return;
          }

          // Fallback: populate + select the text so the user can edit/confirm.
          this.#addressEl.value = label;
          try {
            this.#addressEl.focus({ preventScroll: true });
          } catch {
            this.#addressEl.focus();
          }
          this.#addressEl.select();
        },
      });
    }

    if (items.length === 0) {
      items.push({
        type: "item",
        label: "No named ranges",
        enabled: false,
        onSelect: () => {},
      });
    }

    const rect = this.#nameBoxDropdownEl.getBoundingClientRect();
    this.#nameBoxDropdownEl.setAttribute("aria-expanded", "true");
    this.#addressEl.setAttribute("aria-expanded", "true");
    menu.open({ x: rect.left, y: rect.bottom, items });
  }

  setArgumentPreviewProvider(provider: ((expr: string) => unknown | Promise<unknown>) | null): void {
    this.#argumentPreviewProvider = provider;
    this.#clearArgumentPreviewState();
    this.#render({ preserveTextareaValue: true });
  }

  setReadOnly(readOnly: boolean, opts: { role?: string | null } = {}): void {
    const next = Boolean(readOnly);
    if (next === this.#readOnly && this.textarea.readOnly === next) return;
    this.#readOnly = next;
    this.root.classList.toggle("formula-bar--read-only", next);
    // `readOnly` (not `disabled`) keeps the textarea focusable/selectable for copy.
    this.textarea.readOnly = next;
    this.textarea.setAttribute("aria-readonly", next ? "true" : "false");
    const suffix = opts.role ? ` (${String(opts.role)})` : "";
    this.textarea.title = next ? `Read-only${suffix}` : "";

    // If permissions flipped while editing, exit edit mode so we never show commit/cancel
    // controls that would no-op.
    if (next && this.model.isEditing) {
      this.#functionAutocomplete.close();
      this.#closeFunctionPicker({ restoreFocus: false });
      this.model.cancel();
      this.#hoverOverride = null;
      this.#hoverOverrideText = null;
      this.#mouseDownSelectedReferenceIndex = null;
      this.#selectedReferenceIndex = null;
    }

    this.#render({ preserveTextareaValue: false });
    this.#emitOverlays();
  }

  #setNameBoxError(message = "Invalid reference"): void {
    this.#isNameBoxInvalid = true;
    this.#nameBoxEl.classList.add("formula-bar-name-box--invalid");
    this.#addressEl.setAttribute("aria-invalid", "true");

    // Provide a lightweight Excel-like tooltip under the input.
    this.#nameBoxErrorEl.textContent = message;
    this.#nameBoxErrorEl.hidden = false;

    // Accessibility: associate the tooltip with the input.
    const describedBy = this.#addressEl.getAttribute("aria-describedby") ?? "";
    const tokens = describedBy
      .split(/\s+/)
      .map((t) => t.trim())
      .filter(Boolean);
    if (!tokens.includes(this.#nameBoxErrorId)) {
      tokens.push(this.#nameBoxErrorId);
      this.#addressEl.setAttribute("aria-describedby", tokens.join(" "));
    }
  }

  #clearNameBoxError(): void {
    if (!this.#isNameBoxInvalid) return;
    this.#isNameBoxInvalid = false;
    this.#nameBoxEl.classList.remove("formula-bar-name-box--invalid");
    this.#addressEl.removeAttribute("aria-invalid");
    this.#nameBoxErrorEl.hidden = true;

    const describedBy = this.#addressEl.getAttribute("aria-describedby") ?? "";
    const next = describedBy
      .split(/\s+/)
      .map((t) => t.trim())
      .filter((t) => t && t !== this.#nameBoxErrorId)
      .join(" ");
    if (next) {
      this.#addressEl.setAttribute("aria-describedby", next);
    } else {
      this.#addressEl.removeAttribute("aria-describedby");
    }
  }

  setAiSuggestion(suggestion: string | { text: string; preview?: unknown } | null): void {
    this.model.setAiSuggestion(suggestion);
    this.#render({ preserveTextareaValue: true });
  }

  focus(opts: { cursor?: "end" | "all" } = {}): void {
    // Ensure the textarea is visible so `.focus()` works even when the formula bar is not currently editing.
    // `#render()` keeps this class in sync with `model.isEditing`, but `focus()` is called while
    // still in view mode, so we need to allow the textarea to become focusable first.
    this.root.classList.add("formula-bar--editing");
    // Prevent browser focus handling from scrolling the desktop shell horizontally.
    // (The app uses its own scroll containers; window scrolling is accidental and
    // breaks pointer-coordinate based interactions like range-drag insertion.)
    try {
      this.textarea.focus({ preventScroll: true });
    } catch {
      // Older browsers may not support FocusOptions; fall back to default focus behavior.
      this.textarea.focus();
    }
    if (opts.cursor === "all") {
      this.textarea.setSelectionRange(0, this.textarea.value.length);
    } else if (opts.cursor === "end") {
      const end = this.textarea.value.length;
      this.textarea.setSelectionRange(end, end);
    }
    // Programmatic focus can be invoked while the textarea is already focused (e.g. after a commit/cancel
    // that doesn't fully blur in all environments). Ensure we still transition into edit mode, mirroring
    // the textarea focus listener.
    this.#beginEditFromFocus();
    this.#onSelectionChange();
  }

  setActiveCell(info: { address: string; input: string; value: unknown; nameBox?: string }): void {
    const { nameBox, ...activeCell } = info;
    this.#nameBoxValue = nameBox ?? activeCell.address;

    // Keep the Name Box display in sync with selection changes even while editing
    // (but never clobber the user's in-progress typing in the Name Box itself).
    if (document.activeElement !== this.#addressEl && this.#addressEl.value !== this.#nameBoxValue) {
      this.#addressEl.value = this.#nameBoxValue;
    }

    if (this.model.isEditing) return;
    // If the user was hovering a reference span in view mode, switching the active cell can
    // replace the highlighted markup without triggering a `mouseleave` event. Capture whether we
    // had any hover-derived overlays so we can explicitly clear them after re-rendering.
    const hadHoverOverride = this.#hoverOverride != null || this.#hoverOverrideText != null;
    this.model.setActiveCell(activeCell);
    this.#hoverOverride = null;
    this.#hoverOverrideText = null;
    this.#mouseDownSelectedReferenceIndex = null;
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: false });
    if (hadHoverOverride) {
      this.#emitOverlays();
    }
  }

  isEditing(): boolean {
    return this.model.isEditing;
  }

  /**
   * Best-effort formula locale ID used for editor tooling and locale-aware UI affordances
   * (e.g. localized function-name autocompletion).
   */
  currentLocaleId(): string {
    const raw =
      this.#tooling?.getLocaleId?.() ??
      (typeof document !== "undefined" ? document.documentElement?.lang : "") ??
      "en-US";
    const trimmed = String(raw ?? "").trim();
    return trimmed || "en-US";
  }

  commitEdit(reason: FormulaBarCommitReason = "command", shift = false): void {
    this.#commit({ reason, shift });
  }

  cancelEdit(): void {
    this.#cancel();
  }

  isFormulaEditing(): boolean {
    return this.model.isEditing && isFormulaText(this.model.draft);
  }

  beginRangeSelection(range: RangeAddress, sheetId?: string): void {
    if (this.#readOnly) return;
    this.#functionAutocomplete.close();
    this.model.beginEdit();
    this.model.beginRangeSelection(range, sheetId);
    this.#mouseDownSelectedReferenceIndex = null;
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: false });
    this.#setTextareaSelectionFromModel();
    this.#emitOverlays();
    this.#scheduleEngineTooling();
  }

  updateRangeSelection(range: RangeAddress, sheetId?: string): void {
    if (this.#readOnly) return;
    this.#functionAutocomplete.close();
    this.model.updateRangeSelection(range, sheetId);
    this.#mouseDownSelectedReferenceIndex = null;
    this.#selectedReferenceIndex = null;
    // Range selection can update rapidly while the user drags; avoid forcing a full
    // highlight rebuild on every event. Update the textarea value/selection immediately
    // so commits reflect the latest range, and coalesce the expensive highlight render
    // to the next animation frame.
    this.textarea.value = this.model.draft;
    this.#setTextareaSelectionFromModel();
    this.#requestRender({ preserveTextareaValue: true });
    this.#emitOverlays();
    this.#scheduleEngineTooling();
  }

  endRangeSelection(): void {
    this.model.endRangeSelection();
  }

  #beginEditFromFocus(): void {
    if (this.#readOnly) return;
    if (this.model.isEditing) return;
    this.#errorPanelReferenceHighlights = null;
    this.model.beginEdit();
    this.#callbacks.onBeginEdit?.(this.model.activeCell.address);
    // Best-effort: if we're editing a formula, start loading signature metadata so function
    // hints can show argument names for the full catalog without blocking initial render.
    if (isFormulaText(this.model.draft) && !isFunctionSignatureCatalogReady()) {
      void preloadFunctionSignatureCatalog();
    }
    // Hover overrides are a view-mode affordance and should not leak into editing behavior.
    this.#hoverOverride = null;
    this.#hoverOverrideText = null;
    this.#mouseDownSelectedReferenceIndex = null;
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: true });
    this.#emitOverlays();
    this.#scheduleEngineTooling();
  }

  #onInput(): void {
    if (!this.model.isEditing) return;

    const value = this.textarea.value;
    const start = this.textarea.selectionStart ?? value.length;
    const end = this.textarea.selectionEnd ?? value.length;

    this.model.updateDraft(value, start, end);
    // If the user just started typing a formula, begin loading signature metadata in the background.
    if (isFormulaText(value) && !isFunctionSignatureCatalogReady()) {
      void preloadFunctionSignatureCatalog();
    }
    this.#selectedReferenceIndex = this.#inferSelectedReferenceIndex(start, end);
    this.#requestRender({ preserveTextareaValue: true });
    this.#emitOverlays();
    this.#scheduleEngineTooling();
  }

  #onSelectionChange(): void {
    if (!this.model.isEditing) return;

    // Selection/cursor updates are common and usually do not change the underlying text.
    // Avoid reading `textarea.value` (and comparing potentially-large strings) in this hot path.
    const draftLen = this.model.draft.length;
    const start = this.textarea.selectionStart ?? draftLen;
    const end = this.textarea.selectionEnd ?? draftLen;

    if (this.model.cursorStart === start && this.model.cursorEnd === end) return;

    this.model.updateDraft(this.model.draft, start, end);
    this.#selectedReferenceIndex = this.#inferSelectedReferenceIndex(start, end);
    this.#requestRender({ preserveTextareaValue: true });
    this.#emitOverlays();
    this.#scheduleEngineTooling();
  }

  #onTextareaMouseDown(e: MouseEvent): void {
    if (!this.model.isEditing) return;
    // Only track click-to-select toggle state for primary-button interactions.
    // (Right-click/context menu can fire mousedown without a corresponding click.)
    if (e.button !== 0) {
      this.#mouseDownSelectedReferenceIndex = null;
      return;
    }

    // Clicking a textarea can collapse an existing selection *before* the `click` event
    // fires (often emitting a `select` event in between). Capture whether a full
    // reference token was selected at pointer-down time so `#onTextareaClick()` can
    // reliably implement Excel-style click-to-select / click-again-to-edit toggling.
    const start = this.textarea.selectionStart ?? this.textarea.value.length;
    const end = this.textarea.selectionEnd ?? this.textarea.value.length;
    this.model.updateDraft(this.textarea.value, start, end);
    this.#mouseDownSelectedReferenceIndex = this.#inferSelectedReferenceIndex(start, end);
  }

  #onTextareaClick(): void {
    if (!this.model.isEditing) return;

    const prevSelectedReferenceIndex = this.#mouseDownSelectedReferenceIndex ?? this.#selectedReferenceIndex;
    this.#mouseDownSelectedReferenceIndex = null;
    const value = this.textarea.value;
    const start = this.textarea.selectionStart ?? value.length;
    const end = this.textarea.selectionEnd ?? value.length;
    this.model.updateDraft(value, start, end);
    this.#selectedReferenceIndex = this.#inferSelectedReferenceIndex(start, end);

    const isFormulaEditing = isFormulaText(this.model.draft);
    if (isFormulaEditing && start === end) {
      const activeIndex = this.model.activeReferenceIndex();
      const active = activeIndex == null ? null : this.model.coloredReferences()[activeIndex] ?? null;

      if (active) {
        // Excel UX: clicking a reference selects the entire reference token so
        // range dragging replaces it. A subsequent click on the same reference
        // toggles back to a caret, allowing manual edits within the reference.
        if (prevSelectedReferenceIndex === activeIndex) {
          this.#selectedReferenceIndex = null;
        } else {
          this.textarea.setSelectionRange(active.start, active.end);
          this.model.updateDraft(this.textarea.value, active.start, active.end);
          this.#selectedReferenceIndex = activeIndex;
        }
      }
    }

    this.#requestRender({ preserveTextareaValue: true });
    this.#emitOverlays();
    this.#scheduleEngineTooling();
  }

  #requestRender(opts: { preserveTextareaValue: boolean }): void {
    // Merge pending render options; if any caller needs to overwrite the textarea
    // value, the combined render must also overwrite it.
    if (this.#pendingRender) {
      this.#pendingRender.preserveTextareaValue = this.#pendingRender.preserveTextareaValue && opts.preserveTextareaValue;
    } else {
      this.#pendingRender = opts;
    }

    // Coalesce multiple rapid input/keyup/select events into a single render per frame.
    if (this.#scheduledRender) return;

    const flush = (): void => {
      // Clear the scheduled handle before rendering so a render can schedule another frame.
      this.#scheduledRender = null;
      const pending = this.#pendingRender ?? { preserveTextareaValue: true };
      this.#pendingRender = null;
      this.#render(pending);
    };

    if (typeof requestAnimationFrame === "function") {
      const id = requestAnimationFrame(() => flush());
      this.#scheduledRender = { id, kind: "raf" };
    } else {
      const id = setTimeout(() => flush(), 0);
      this.#scheduledRender = { id, kind: "timeout" };
    }
  }

  #cancelPendingRender(): void {
    if (!this.#scheduledRender) return;
    const { id, kind } = this.#scheduledRender;
    if (kind === "raf") {
      // `cancelAnimationFrame` isn't present in all JS environments; fall back safely.
      try {
        cancelAnimationFrame(id);
      } catch {
        // ignore
      }
    } else {
      clearTimeout(id);
    }
    this.#scheduledRender = null;
    this.#pendingRender = null;
  }

  #scheduleEngineTooling(): void {
    // Only run editor-tooling calls while editing; the formula bar view mode already
    // uses stable highlights and we want to avoid late async updates after commit/cancel.
    if (!this.model.isEditing) {
      this.#cancelPendingTooling();
      return;
    }

    // Only ask the engine to lex/parse when the draft is actually a formula.
    // This avoids surfacing parse errors while editing plain text values.
    const draft = this.model.draft;
    if (!isFormulaText(draft)) {
      this.#cancelPendingTooling();
      return;
    }

    const engine = this.#tooling?.getWasmEngine?.() ?? null;
    if (!engine) return;

    const rawLocaleId =
      this.#tooling?.getLocaleId?.() ??
      (typeof document !== "undefined" ? document.documentElement?.lang : "") ??
      "en-US";
    // `parseFormulaPartial` expects a supported formula locale id. Normalize variants and fall
    // back to `en-US` when the host locale is unsupported so engine tooling stays available.
    const localeId = normalizeFormulaLocaleId(rawLocaleId) ?? "en-US";
    const referenceStyle = this.#tooling?.referenceStyle ?? "A1";

    const cursor = this.model.cursorStart;
    const requestId = ++this.#toolingRequestId;

    if (this.#toolingPending) {
      this.#toolingPending.requestId = requestId;
      this.#toolingPending.draft = draft;
      this.#toolingPending.cursor = cursor;
      this.#toolingPending.localeId = localeId || "en-US";
      this.#toolingPending.referenceStyle = referenceStyle;
    } else {
      this.#toolingPending = { requestId, draft, cursor, localeId: localeId || "en-US", referenceStyle };
    }

    // Coalesce multiple rapid edits/selection changes into one tooling request per frame.
    if (this.#toolingScheduled) return;

    const flush = (): void => {
      this.#toolingScheduled = null;
      const pending = this.#toolingPending;
      this.#toolingPending = null;
      if (!pending) return;
      void this.#runEngineTooling(pending);
    };

    if (typeof requestAnimationFrame === "function") {
      const id = requestAnimationFrame(() => flush());
      this.#toolingScheduled = { id, kind: "raf" };
    } else {
      const id = setTimeout(() => flush(), 0);
      this.#toolingScheduled = { id, kind: "timeout" };
    }
  }

  async #runEngineTooling(pending: {
    requestId: number;
    draft: string;
    cursor: number;
    localeId: string;
    referenceStyle: NonNullable<FormulaParseOptions["referenceStyle"]>;
  }): Promise<void> {
    // It's possible for a scheduled tooling flush to run after commit/cancel; bail early
    // before invoking any async engine work.
    if (!this.model.isEditing) return;
    if (this.model.draft !== pending.draft) return;
    const engine = this.#tooling?.getWasmEngine?.() ?? null;
    if (!engine) return;
    if (!isFormulaText(pending.draft)) return;

    try {
      const options: FormulaParseOptions = { localeId: pending.localeId, referenceStyle: pending.referenceStyle };
      // Abort any in-flight tooling request so rapid typing doesn't queue up work in the worker.
      this.#toolingAbort?.abort();
      const abort = new AbortController();
      this.#toolingAbort = abort;
      const rpcOptions = { signal: abort.signal };

      const cached = this.#toolingLexCache;
      const cacheHit =
        cached != null &&
        cached.draft === pending.draft &&
        cached.localeId === pending.localeId &&
        cached.referenceStyle === pending.referenceStyle;

      let lexResult: Awaited<ReturnType<EngineClient["lexFormulaPartial"]>>;
      let parseResult: Awaited<ReturnType<EngineClient["parseFormulaPartial"]>>;

      if (cacheHit) {
        // When only the caret moves, reuse the cached lexer result and avoid the extra Promise.all/array
        // allocations in this hot path.
        lexResult = cached.lexResult;
        parseResult = await engine.parseFormulaPartial(pending.draft, pending.cursor, options, rpcOptions);
      } else {
        [lexResult, parseResult] = await Promise.all([
          engine.lexFormulaPartial(pending.draft, options, rpcOptions),
          engine.parseFormulaPartial(pending.draft, pending.cursor, options, rpcOptions),
        ]);

        this.#toolingLexCache = {
          draft: pending.draft,
          localeId: pending.localeId,
          referenceStyle: pending.referenceStyle,
          lexResult,
        };
      }

      // Ignore stale/out-of-order results.
      if (pending.requestId !== this.#toolingRequestId) return;
      if (!this.model.isEditing) return;
      if (this.model.draft !== pending.draft) return;

      this.model.applyEngineToolingResult({ formula: pending.draft, localeId: pending.localeId, lexResult, parseResult });
      this.#requestRender({ preserveTextareaValue: true });
    } catch {
      // Best-effort: if the engine worker is unavailable/uninitialized, keep the local
      // tokenizer/highlighter without surfacing errors to the user.
    }
  }

  #cancelPendingTooling(): void {
    this.#toolingAbort?.abort();
    this.#toolingAbort = null;
    if (this.#toolingScheduled) {
      const { id, kind } = this.#toolingScheduled;
      if (kind === "raf") {
        try {
          cancelAnimationFrame(id);
        } catch {
          // ignore
        }
      } else {
        clearTimeout(id);
      }
    }
    this.#toolingScheduled = null;
    this.#toolingPending = null;
    // Bump request id so any in-flight engine responses are considered stale.
    this.#toolingRequestId += 1;
  }

  #onKeyDown(e: KeyboardEvent): void {
    if (!this.model.isEditing) return;

    if (this.#isComposing || e.isComposing) {
      // While IME composition is active, avoid interpreting navigation/commit keys.
      // However, we still need to prevent browser focus traversal on Tab so composition
      // isn't interrupted by moving focus away from the formula bar.
      if (e.key === "Tab") {
        e.preventDefault();
        e.stopPropagation();
        return;
      }

      if (e.key === "Enter" || e.key === "Escape" || e.key === "F4" || e.key === "ArrowDown" || e.key === "ArrowUp") {
        return;
      }
    }

    if (this.#functionAutocomplete.handleKeyDown(e)) return;

    if (e.key === "F4" && !e.altKey && !e.ctrlKey && !e.metaKey && isFormulaText(this.model.draft)) {
      e.preventDefault();

      const prevText = this.textarea.value;
      const cursorStart = this.textarea.selectionStart ?? prevText.length;
      const cursorEnd = this.textarea.selectionEnd ?? prevText.length;

      // Ensure model-derived reference metadata is current for the F4 operation
      // (the selection may have changed without triggering our keyup/select listeners yet).
      if (
        this.model.draft !== prevText ||
        this.model.cursorStart !== cursorStart ||
        this.model.cursorEnd !== cursorEnd
      ) {
        this.model.updateDraft(prevText, cursorStart, cursorEnd);
      }

      const toggled = toggleA1AbsoluteAtCursor(prevText, cursorStart, cursorEnd);
      if (!toggled) return;

      this.textarea.value = toggled.text;
      this.textarea.setSelectionRange(toggled.cursorStart, toggled.cursorEnd);
      this.model.updateDraft(toggled.text, toggled.cursorStart, toggled.cursorEnd);
      this.#selectedReferenceIndex = this.#inferSelectedReferenceIndex(toggled.cursorStart, toggled.cursorEnd);
      this.#render({ preserveTextareaValue: true });
      this.#emitOverlays();
      this.#scheduleEngineTooling();
      return;
    }

    if (e.key === "Tab") {
      // Excel-like behavior: Tab/Shift+Tab commits the edit (and the app navigates selection).
      // Exception: plain Tab accepts an AI suggestion if one is available.
      //
      // Never allow default browser focus traversal while editing. (During IME composition we
      // also prevent focus traversal, but avoid committing so the IME can finish composing.)
      if (!e.shiftKey) {
        const accepted = this.model.acceptAiSuggestion();
        if (accepted) {
          e.preventDefault();
          this.#selectedReferenceIndex = null;
          this.#render({ preserveTextareaValue: false });
          this.#setTextareaSelectionFromModel();
          this.#emitOverlays();
          this.#scheduleEngineTooling();
          return;
        }
      }

      e.preventDefault();
      this.#commit({ reason: "tab", shift: e.shiftKey });
      return;
    }

    if (e.key === "Escape") {
      e.preventDefault();
      this.#cancel();
      return;
    }

    // Excel behavior: Enter commits, Alt+Enter inserts newline.
    if (e.key === "Enter" && e.altKey) {
      e.preventDefault();

      const prevText = this.textarea.value;
      const cursorStart = this.textarea.selectionStart ?? prevText.length;
      const cursorEnd = this.textarea.selectionEnd ?? prevText.length;

      const indentation = computeFormulaIndentation(prevText, cursorStart);
      const insertion = `\n${indentation}`;

      this.textarea.value = prevText.slice(0, cursorStart) + insertion + prevText.slice(cursorEnd);

      const nextCursor = cursorStart + insertion.length;
      this.textarea.setSelectionRange(nextCursor, nextCursor);

       // Reuse the standard input path to keep the model + highlight in sync.
       this.#onInput();
       return;
     }

    if (e.key === "Enter" && !e.altKey) {
      e.preventDefault();
      this.#commit({ reason: "enter", shift: e.shiftKey });
      return;
    }
  }

  #cancel(): void {
    if (!this.model.isEditing) return;
    this.#functionAutocomplete.close();
    this.#closeFunctionPicker({ restoreFocus: false });
    this.textarea.blur();
    this.model.cancel();
    this.#cancelPendingTooling();
    this.#hoverOverride = null;
    this.#hoverOverrideText = null;
    this.#selectedReferenceIndex = null;
    this.#mouseDownSelectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: false });
    try {
      this.#callbacks.onCancel?.();
    } catch (err) {
      // Embedding callback errors should not break editing UX (or surface as unhandled
      // window errors in tests). Log for visibility and keep the view consistent.
      console.error("FormulaBarView.onCancel threw", err);
    } finally {
      // Even if the embedding callback throws, ensure we still clear hover/range overlays.
      this.#emitOverlays();
    }
  }

  #commit(commit: FormulaBarCommit): void {
    if (!this.model.isEditing) return;
    this.#functionAutocomplete.close();
    this.#closeFunctionPicker({ restoreFocus: false });
    this.textarea.blur();
    const committed = this.model.commit();
    this.#cancelPendingTooling();
    this.#hoverOverride = null;
    this.#hoverOverrideText = null;
    this.#selectedReferenceIndex = null;
    this.#mouseDownSelectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: false });
    try {
      this.#callbacks.onCommit(committed, commit);
    } catch (err) {
      // Embedding callback errors should not break editing UX (or surface as unhandled
      // window errors in tests). Log for visibility and keep the view consistent.
      console.error("FormulaBarView.onCommit threw", err);
    } finally {
      // Even if the embedding callback throws, ensure we still clear hover/range overlays.
      this.#emitOverlays();
    }
  }

  #focusFx(): void {
    // If the formula bar isn't mounted, avoid stealing focus (and avoid creating global pickers).
    if (!this.root.isConnected) return;

    // Avoid overlapping UI affordances (typing autocomplete vs. explicit fx picker).
    this.#functionAutocomplete.close();

    // Excel-style: clicking fx focuses the formula input and commonly starts a formula.
    if (this.model.isEditing) this.focus();
    else this.focus({ cursor: "end" });

    if (!this.model.isEditing) return;

    this.#openFunctionPicker();
  }

  #openFunctionPicker(): void {
    if (this.#functionPickerOpen) {
      this.#functionPickerInputEl.focus();
      this.#functionPickerInputEl.select();
      return;
    }
    if (!this.root.isConnected) return;
    if (!this.model.isEditing) return;

    const start = this.textarea.selectionStart ?? this.textarea.value.length;
    const end = this.textarea.selectionEnd ?? this.textarea.value.length;
    this.#functionPickerAnchorSelection = { start, end };

    this.#functionPickerOpen = true;
    this.#functionPickerEl.hidden = false;
    this.#fxButtonEl.setAttribute("aria-expanded", "true");
    this.#functionPickerInputEl.setAttribute("aria-expanded", "true");
    this.#isFunctionPickerComposing = false;
    this.#functionPickerInputEl.value = "";
    this.#functionPickerSelectedIndex = 0;

    // Best-effort: load signature metadata in the background so the picker can show richer hints.
    void preloadFunctionSignatureCatalog()
      .then(() => {
        if (!this.#functionPickerOpen) return;
        this.#renderFunctionPickerResults();
      })
      .catch(() => {
        // Best-effort: ignore catalog preload failures.
      });

    this.#positionFunctionPicker();
    this.#renderFunctionPickerResults();

    document.addEventListener("mousedown", this.#functionPickerDocMouseDown, { capture: true, signal: this.#domAbort.signal });

    this.#functionPickerInputEl.focus();
    this.#functionPickerInputEl.select();
  }

  #closeFunctionPicker(opts: { restoreFocus: boolean } = { restoreFocus: true }): void {
    if (!this.#functionPickerOpen) return;
    this.#functionPickerOpen = false;
    this.#functionPickerEl.hidden = true;
    this.#fxButtonEl.setAttribute("aria-expanded", "false");
    this.#functionPickerInputEl.setAttribute("aria-expanded", "false");
    this.#functionPickerInputEl.removeAttribute("aria-activedescendant");
    this.#isFunctionPickerComposing = false;
    this.#functionPickerItems = [];
    this.#functionPickerItemEls = [];
    this.#functionPickerSelectedIndex = 0;
    const anchor = this.#functionPickerAnchorSelection;
    this.#functionPickerAnchorSelection = null;

    document.removeEventListener("mousedown", this.#functionPickerDocMouseDown, true);

    if (!opts.restoreFocus) return;
    if (!this.root.isConnected) return;

    try {
      this.textarea.focus({ preventScroll: true });
    } catch {
      this.textarea.focus();
    }

    if (anchor) {
      this.textarea.setSelectionRange(anchor.start, anchor.end);
      this.model.updateDraft(this.textarea.value, anchor.start, anchor.end);
      this.#selectedReferenceIndex = this.#inferSelectedReferenceIndex(anchor.start, anchor.end);
      this.#render({ preserveTextareaValue: true });
      this.#emitOverlays();
      this.#scheduleEngineTooling();
    }
  }

  #onFunctionPickerDocMouseDown(e: MouseEvent): void {
    if (!this.#functionPickerOpen) return;
    const target = e.target as Node | null;
    if (!target) return;
    if (this.#functionPickerEl.contains(target)) return;
    if (this.#fxButtonEl.contains(target)) return;
    // Clicking outside should close the popover without stealing focus from the clicked surface.
    this.#closeFunctionPicker({ restoreFocus: false });
  }

  #positionFunctionPicker(): void {
    // Anchor below the fx button by default.
    const rootRect = this.root.getBoundingClientRect();
    const fxRect = this.#fxButtonEl.getBoundingClientRect();
    const top = fxRect.bottom - rootRect.top + 6;
    const left = fxRect.left - rootRect.left;
    this.#functionPickerEl.style.top = `${Math.max(0, Math.round(top))}px`;
    this.#functionPickerEl.style.left = `${Math.max(0, Math.round(left))}px`;
  }

  #onFunctionPickerInput(): void {
    if (!this.#functionPickerOpen) return;
    this.#functionPickerSelectedIndex = 0;
    this.#renderFunctionPickerResults();
  }

  #onFunctionPickerKeyDown(e: KeyboardEvent): void {
    if (!this.#functionPickerOpen) return;

    if (
      (this.#isFunctionPickerComposing || e.isComposing) &&
      (e.key === "Enter" || e.key === "Escape" || e.key === "ArrowDown" || e.key === "ArrowUp")
    ) {
      return;
    }

    if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      this.#closeFunctionPicker({ restoreFocus: true });
      return;
    }

    if (e.key === "ArrowDown") {
      e.preventDefault();
      e.stopPropagation();
      this.#updateFunctionPickerSelection(this.#functionPickerSelectedIndex + 1);
      return;
    }

    if (e.key === "ArrowUp") {
      e.preventDefault();
      e.stopPropagation();
      this.#updateFunctionPickerSelection(this.#functionPickerSelectedIndex - 1);
      return;
    }

    if (e.key === "Enter") {
      e.preventDefault();
      e.stopPropagation();
      this.#selectFunctionPickerItem(this.#functionPickerSelectedIndex);
    }
  }

  #updateFunctionPickerSelection(nextIndex: number): void {
    if (this.#functionPickerItems.length === 0) {
      this.#functionPickerSelectedIndex = 0;
      this.#functionPickerInputEl.removeAttribute("aria-activedescendant");
      return;
    }

    const clamped = Math.max(0, Math.min(nextIndex, this.#functionPickerItems.length - 1));
    const prev = this.#functionPickerSelectedIndex;
    this.#functionPickerSelectedIndex = clamped;

    const prevEl = this.#functionPickerItemEls[prev];
    if (prevEl) prevEl.setAttribute("aria-selected", "false");

    const nextEl = this.#functionPickerItemEls[clamped];
    if (nextEl) {
      nextEl.setAttribute("aria-selected", "true");
      if (nextEl.id) {
        this.#functionPickerInputEl.setAttribute("aria-activedescendant", nextEl.id);
      } else {
        this.#functionPickerInputEl.removeAttribute("aria-activedescendant");
      }
      if (typeof nextEl.scrollIntoView === "function") nextEl.scrollIntoView({ block: "nearest" });
    } else {
      this.#functionPickerInputEl.removeAttribute("aria-activedescendant");
    }
  }

  #selectFunctionPickerItem(index: number): void {
    const item = this.#functionPickerItems[index];
    if (!item) return;
    const anchor = this.#functionPickerAnchorSelection;
    if (!anchor) return;

    this.#closeFunctionPicker({ restoreFocus: false });
    this.#insertFunctionAtSelection(item.name, anchor);
  }

  #insertFunctionAtSelection(name: string, selection: { start: number; end: number }): void {
    if (!this.root.isConnected) return;
    if (!this.model.isEditing) return;

    const prevText = this.textarea.value;
    const isEmpty = prevText.trim() === "";
    const start = Math.max(0, Math.min(selection.start, prevText.length));
    const end = Math.max(0, Math.min(selection.end, prevText.length));

    const insert = `${name}()`;
    // If the user was editing an empty cell, selecting a function should insert the
    // leading "=" (Excel behavior) so the result is a valid formula.
    const nextText = isEmpty ? `=${insert}` : prevText.slice(0, start) + insert + prevText.slice(end);
    // Place the caret inside the parentheses so users can immediately type arguments.
    const cursor = isEmpty
      ? Math.max(0, nextText.length - 1)
      : Math.max(0, Math.min(start + insert.length - 1, nextText.length));

    this.textarea.value = nextText;
    try {
      this.textarea.focus({ preventScroll: true });
    } catch {
      this.textarea.focus();
    }
    this.textarea.setSelectionRange(cursor, cursor);
    this.model.updateDraft(nextText, cursor, cursor);
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: true });
    this.#emitOverlays();
    this.#scheduleEngineTooling();
  }

  #renderFunctionPickerResults(): void {
    const limit = 50;
    const query = this.#functionPickerInputEl.value;
    const localeId = this.currentLocaleId();
    const items: FunctionPickerItem[] = buildFunctionPickerItems(query, limit, localeId);

    this.#functionPickerItems = items;
    this.#functionPickerItemEls = renderFunctionPickerList({
      listEl: this.#functionPickerListEl,
      query,
      items,
      selectedIndex: this.#functionPickerSelectedIndex,
      onSelect: (index) => this.#selectFunctionPickerItem(index),
    });
    const optionPrefix = `${this.#functionPickerListEl.id || "formula-function-picker-list"}-option-`;
    for (let i = 0; i < this.#functionPickerItemEls.length; i += 1) {
      this.#functionPickerItemEls[i]!.id = `${optionPrefix}${i}`;
    }

    // Ensure selection is valid after query changes.
    this.#updateFunctionPickerSelection(this.#functionPickerSelectedIndex);
  }

  #render(opts: { preserveTextareaValue: boolean }): void {
    // If we're rendering synchronously (e.g. commit/cancel/AI suggestion), cancel any
    // pending scheduled render so we don't immediately re-render the same state.
    this.#cancelPendingRender();

    if (!this.model.isEditing) {
      this.#selectedReferenceIndex = null;
    }

    if (document.activeElement !== this.#addressEl && this.#addressEl.value !== this.#nameBoxValue) {
      this.#addressEl.value = this.#nameBoxValue;
    }

    if (!opts.preserveTextareaValue) {
      this.textarea.value = this.model.draft;
    }

    const showEditingActions = this.model.isEditing;
    if (this.#lastShowEditingActions !== showEditingActions) {
      this.#lastShowEditingActions = showEditingActions;
      this.#cancelButtonEl.hidden = !showEditingActions;
      this.#cancelButtonEl.disabled = !showEditingActions;
      this.#commitButtonEl.hidden = !showEditingActions;
      this.#commitButtonEl.disabled = !showEditingActions;
    }

    const cursor = this.model.cursorStart;
    const ghost = this.model.isEditing ? this.model.aiGhostText() : "";
    const previewRaw = this.model.isEditing ? this.model.aiSuggestionPreview() : null;
    const previewText = ghost && previewRaw != null ? formatPreview(previewRaw) : "";
    const draft = this.model.draft;
    const draftVersion = this.model.draftVersion;

    const isFormulaEditing = this.model.isEditing && isFormulaText(draft);
    const coloredReferences = isFormulaEditing ? this.model.coloredReferences() : EMPTY_COLORED_REFERENCES;
    const activeReferenceIndex = isFormulaEditing ? this.model.activeReferenceIndex() : null;
    const highlightedSpans = this.model.highlightedSpans();

    const canFastUpdateActiveReference =
      isFormulaEditing &&
      !ghost &&
      this.#lastHighlightDraft === draft &&
      this.#lastHighlightIsFormulaEditing &&
      !this.#lastHighlightHadGhost &&
      this.#lastHighlightSpans === highlightedSpans &&
      this.#lastColoredReferences === coloredReferences;

    const canSkipHighlightWithGhost =
      Boolean(ghost) &&
      this.#lastHighlightDraft === draft &&
      this.#lastHighlightIsFormulaEditing === isFormulaEditing &&
      this.#lastHighlightHadGhost &&
      this.#lastHighlightCursor === cursor &&
      this.#lastHighlightGhost === ghost &&
      this.#lastHighlightPreviewText === previewText &&
      this.#lastHighlightSpans === highlightedSpans &&
      this.#lastColoredReferences === coloredReferences &&
      this.#lastActiveReferenceIndex === activeReferenceIndex;

    const canSkipHighlight =
      !ghost &&
      !isFormulaEditing &&
      this.#lastHighlightDraft === draft &&
      !this.#lastHighlightIsFormulaEditing &&
      !this.#lastHighlightHadGhost &&
      this.#lastHighlightSpans === highlightedSpans;

    if (canSkipHighlightWithGhost) {
      // When an AI suggestion is visible, other parts of the formula bar can still re-render
      // (e.g. engine tooling/hint updates). If the draft/cursor/ghost state is unchanged,
      // skip rebuilding the highlight HTML string for long formulas.
      this.#lastHighlightDraft = draft;
      this.#lastHighlightIsFormulaEditing = isFormulaEditing;
      this.#lastHighlightHadGhost = true;
      this.#lastHighlightCursor = cursor;
      this.#lastHighlightGhost = ghost;
      this.#lastHighlightPreviewText = previewText;
      this.#lastActiveReferenceIndex = activeReferenceIndex;
      this.#lastHighlightSpans = highlightedSpans;
      this.#lastColoredReferences = coloredReferences;
    } else if (canSkipHighlight) {
      // No cursor-dependent styling in view/plain-text mode; avoid rebuilding the highlight HTML
      // string when the draft/tokenization state is unchanged (common on cursor moves).
      this.#lastHighlightDraft = draft;
      this.#lastHighlightIsFormulaEditing = false;
      this.#lastHighlightHadGhost = false;
      this.#lastHighlightCursor = null;
      this.#lastHighlightGhost = null;
      this.#lastHighlightPreviewText = null;
      this.#lastActiveReferenceIndex = null;
      this.#lastHighlightSpans = highlightedSpans;
      this.#lastColoredReferences = coloredReferences;
    } else if (canFastUpdateActiveReference) {
      if (this.#lastActiveReferenceIndex !== activeReferenceIndex) {
        const prev = this.#lastActiveReferenceIndex;
        const next = activeReferenceIndex;
        if (prev != null) {
          this.#referenceElementsForIndex(prev).forEach((el) => el.classList.remove("formula-bar-reference--active"));
        }
        if (next != null) {
          this.#referenceElementsForIndex(next).forEach((el) => el.classList.add("formula-bar-reference--active"));
        }
        this.#lastActiveReferenceIndex = next;
        // We updated class attributes without rebuilding the HTML string; invalidate the
        // string cache so future full renders don't compare against a stale snapshot.
        this.#lastHighlightHtml = null;
      }
      this.#lastHighlightDraft = draft;
      this.#lastHighlightIsFormulaEditing = true;
      this.#lastHighlightHadGhost = false;
      this.#lastHighlightCursor = null;
      this.#lastHighlightGhost = null;
      this.#lastHighlightPreviewText = null;
      this.#lastHighlightSpans = highlightedSpans;
      this.#lastColoredReferences = coloredReferences;
    } else {
      // `highlightedSpans` can be large for long formulas; prefer fixed-length arrays where possible
      // to avoid repeated `push` growth.
      const highlightParts: string[] = ghost
        ? new Array<string>(highlightedSpans.length + 3)
        : new Array<string>(highlightedSpans.length);
      // Escaping every token individually incurs a lot of overhead on long formulas. If the
      // underlying draft contains no HTML-significant characters, we can safely emit token
      // text as-is.
      const needsEscapeDraft = ESCAPE_HTML_TEST_RE.test(draft);

      const refs = coloredReferences;
      const canIdentifierBeReference = Boolean(this.model.extractFormulaReferencesOptions()?.resolveName);
      let refCursor = 0;

      const findContainingRef = (start: number, end: number) => {
        // `coloredReferences` are ordered by appearance in the formula string, so we can
        // walk them alongside the highlight spans (also ordered by start offset).
        while (refCursor < refs.length && refs[refCursor]!.end <= start) refCursor += 1;
        const ref = refs[refCursor];
        if (!ref) return null;
        if (ref.start <= start && end <= ref.end) return ref;
        return null;
      };

      const spansHaveKind =
        highlightedSpans.length > 0 && (highlightedSpans[0] as unknown as { kind?: unknown }).kind != null;

      const renderSpan = (span: { start: number; end: number; className?: string }, text: string, kind: string): string => {
        const extraClass = span.className;
        // Whitespace spans don't receive styling and are never hover targets; avoid wrapping them
        // in <span> tags to keep long, formatted formulas lighter to render.
        if (kind === "whitespace" && !extraClass) {
          return text;
        }

        // If the draft contains any HTML-significant characters, escape only the tokens that
        // can actually include them. This avoids paying the regex test cost on every token when
        // the formula includes a single `&` operator or comparison like "<".
        let content = text;
        if (needsEscapeDraft) {
          // Only a few token kinds can plausibly contain `<`, `>`, or `&`:
          // - operators: `&`, `<`, `>`, `<=`, `<>`, ...
          // - strings: user-entered string literals can contain them
          // - references: quoted sheet names / workbook prefixes can contain them (e.g. `'A&B'!A1`)
          // - unknown: defensive fallback
          if (kind === "operator") {
            if (text === "&") content = "&amp;";
            else if (text === "<") content = "&lt;";
            else if (text === ">") content = "&gt;";
            else if (text.indexOf("&") !== -1 || text.indexOf("<") !== -1 || text.indexOf(">") !== -1) {
              content = escapeHtml(text);
            }
          } else if (kind === "string" || kind === "reference" || kind === "unknown") {
            if (text.indexOf("&") !== -1 || text.indexOf("<") !== -1 || text.indexOf(">") !== -1) {
              content = escapeHtml(text);
            }
          }
        }

        // Identifier spans are only needed:
        // - in view mode when we have a name resolver (so hover previews can resolve named ranges), or
        // - in edit mode when identifier tokens can map to extracted references (named ranges).
        // Otherwise they are unstyled and don't participate in reference highlighting.
        if (!extraClass && (kind === "unknown" || (kind === "identifier" && !canIdentifierBeReference))) {
          return content;
        }

        if (!isFormulaEditing) {
          const classAttr = extraClass ? ` class="${extraClass}"` : "";
          return `<span data-kind="${kind}"${classAttr}>${content}</span>`;
        }

        if (kind === "error") {
          const classAttr = extraClass ? ` class="${extraClass}"` : "";
          return `<span data-kind="${kind}"${classAttr}>${content}</span>`;
        }

        // Only `reference` and `identifier` spans can correspond to extracted references
        // (A1 refs / structured refs are tokenized as `reference`, named ranges as `identifier`).
        // Avoid the per-token reference containment checks for everything else.
        if (kind !== "reference" && !(kind === "identifier" && canIdentifierBeReference)) {
          const classAttr = extraClass ? ` class="${extraClass}"` : "";
          return `<span data-kind="${kind}"${classAttr}>${content}</span>`;
        }

        const containing = findContainingRef(span.start, span.end);
        if (!containing) {
          if (kind === "identifier" && !extraClass) {
            return content;
          }
          const classAttr = extraClass ? ` class="${extraClass}"` : "";
          return `<span data-kind="${kind}"${classAttr}>${content}</span>`;
        }

        const isActive = activeReferenceIndex === containing.index;
        const baseClass = isActive ? "formula-bar-reference formula-bar-reference--active" : "formula-bar-reference";
        const classAttr = extraClass ? ` class="${baseClass} ${extraClass}"` : ` class="${baseClass}"`;
        return `<span data-kind="${kind}" data-ref-index="${containing.index}"${classAttr} style="color: ${containing.color};">${content}</span>`;
      };

      if (!ghost) {
        if (spansHaveKind) {
          for (let i = 0; i < highlightedSpans.length; i += 1) {
            const span = highlightedSpans[i] as unknown as { kind: string; text: string; start: number; end: number; className?: string };
            highlightParts[i] = renderSpan(span, span.text, span.kind);
          }
        } else {
          for (let i = 0; i < highlightedSpans.length; i += 1) {
            const span = highlightedSpans[i] as unknown as { type: string; text: string; start: number; end: number; className?: string };
            highlightParts[i] = renderSpan(span, span.text, span.type);
          }
        }
      } else {
        let outIdx = 0;
        let ghostInserted = false;
        let previewInserted = false;
        const ghostContent = ESCAPE_HTML_TEST_RE.test(ghost) ? escapeHtml(ghost) : ghost;
        const ghostHtml = `<span class="formula-bar-ghost">${ghostContent}</span>`;
        const previewContent = previewText && ESCAPE_HTML_TEST_RE.test(previewText) ? escapeHtml(previewText) : previewText;
        const previewHtml = previewText ? `<span class="formula-bar-preview">${previewContent}</span>` : "";

        if (spansHaveKind) {
          for (let i = 0; i < highlightedSpans.length; i += 1) {
            const span = highlightedSpans[i] as unknown as { kind: string; text: string; start: number; end: number; className?: string };
            const kind = span.kind;
            if (!ghostInserted && cursor <= span.start) {
              highlightParts[outIdx++] = ghostHtml;
              if (previewHtml && !previewInserted) {
                highlightParts[outIdx++] = previewHtml;
                previewInserted = true;
              }
              ghostInserted = true;
            }

            if (!ghostInserted && cursor > span.start && cursor < span.end) {
              const split = cursor - span.start;
              const before = span.text.slice(0, split);
              const after = span.text.slice(split);
              if (before) {
                highlightParts[outIdx++] = renderSpan(span, before, kind);
              }
              highlightParts[outIdx++] = ghostHtml;
              if (previewHtml && !previewInserted) {
                highlightParts[outIdx++] = previewHtml;
                previewInserted = true;
              }
              ghostInserted = true;
              if (after) {
                highlightParts[outIdx++] = renderSpan(span, after, kind);
              }
              continue;
            }

            highlightParts[outIdx++] = renderSpan(span, span.text, kind);
          }
        } else {
          for (let i = 0; i < highlightedSpans.length; i += 1) {
            const span = highlightedSpans[i] as unknown as { type: string; text: string; start: number; end: number; className?: string };
            const kind = span.type;
            if (!ghostInserted && cursor <= span.start) {
              highlightParts[outIdx++] = ghostHtml;
              if (previewHtml && !previewInserted) {
                highlightParts[outIdx++] = previewHtml;
                previewInserted = true;
              }
              ghostInserted = true;
            }

            if (!ghostInserted && cursor > span.start && cursor < span.end) {
              const split = cursor - span.start;
              const before = span.text.slice(0, split);
              const after = span.text.slice(split);
              if (before) {
                highlightParts[outIdx++] = renderSpan(span, before, kind);
              }
              highlightParts[outIdx++] = ghostHtml;
              if (previewHtml && !previewInserted) {
                highlightParts[outIdx++] = previewHtml;
                previewInserted = true;
              }
              ghostInserted = true;
              if (after) {
                highlightParts[outIdx++] = renderSpan(span, after, kind);
              }
              continue;
            }

            highlightParts[outIdx++] = renderSpan(span, span.text, kind);
          }
        }

        if (!ghostInserted) {
          highlightParts[outIdx++] = ghostHtml;
          if (previewHtml && !previewInserted) {
            highlightParts[outIdx++] = previewHtml;
            previewInserted = true;
          }
        }

        // Trim unused preallocated capacity so `.join("")` doesn't iterate over the full array length.
        highlightParts.length = outIdx;
      }

      const highlightHtml = highlightParts.join("");

      // Avoid forcing a full DOM re-parse/layout if the highlight HTML is unchanged.
      // Also keep cached reference-span node lists in sync with the current DOM.
      const highlightChanged = highlightHtml !== this.#lastHighlightHtml;
      if (highlightChanged) {
        this.#highlightEl.innerHTML = highlightHtml;
        this.#referenceElsByIndex = null;
      }
      if (!isFormulaEditing || coloredReferences.length === 0) {
        this.#referenceElsByIndex = null;
      }

      this.#lastHighlightHtml = highlightHtml;
      this.#lastHighlightDraft = draft;
      this.#lastHighlightIsFormulaEditing = isFormulaEditing;
      this.#lastHighlightHadGhost = Boolean(ghost);
      this.#lastHighlightCursor = ghost ? cursor : null;
      this.#lastHighlightGhost = ghost ? ghost : null;
      this.#lastHighlightPreviewText = ghost ? previewText : null;
      this.#lastActiveReferenceIndex = activeReferenceIndex;
      this.#lastHighlightSpans = highlightedSpans;
      this.#lastColoredReferences = coloredReferences;
    }

    // Toggle editing UI state (textarea visibility, hover hit-testing, etc.) through CSS classes.
    const isEditing = this.model.isEditing;
    if (this.#lastRootIsEditingClass !== isEditing) {
      this.#lastRootIsEditingClass = isEditing;
      this.root.classList.toggle("formula-bar--editing", isEditing);
    }

    const syntaxError = isFormulaEditing ? this.model.syntaxError() : null;
    const hasSyntaxError = Boolean(syntaxError);
    if (this.#lastHintHasSyntaxErrorClass !== hasSyntaxError) {
      this.#lastHintHasSyntaxErrorClass = hasSyntaxError;
      this.#hintEl.classList.toggle("formula-bar-hint--syntax-error", hasSyntaxError);
    }
    const hint = isFormulaEditing ? this.model.functionHint() : null;

    // Keep argument preview state up to date, but avoid re-rendering the entire hint panel unless
    // the visible hint content actually changed (cursor moves within the same argument are common
    // on long formulas).
    let wantsArgPreview = false;
    let activeArgForPreview: ReturnType<FormulaBarModel["activeArgumentSpan"]> | null = null;
    let argPreviewKey: string | null = null;

    if (!hint || hasSyntaxError) {
      this.#clearArgumentPreviewState();
    } else {
      const provider = this.#argumentPreviewProvider;
      let activeArg = typeof provider === "function" ? this.model.activeArgumentSpan() : null;

      // Keep the argument preview in sync with the hint behavior when the caret is
      // positioned just after a closing paren, e.g. `=ROUND(1,2)|`. In that case
      // `activeArgumentSpan()` returns null because the tokenizer-based parser has
      // already consumed the closing `)`, but the hint panel still treats the last
      // argument as active (Excel UX).
      if (!activeArg && typeof provider === "function" && this.model.cursorStart === this.model.cursorEnd && this.model.cursorStart > 0) {
        let scan = this.model.cursorStart - 1;
        while (scan >= 0 && isWhitespaceChar(draft[scan] ?? "")) scan -= 1;
        if (scan >= 0 && draft[scan] === ")") {
          activeArg = this.model.activeArgumentSpan(scan);
        }
      }

      if (activeArg && typeof provider === "function" && typeof activeArg.argText === "string" && activeArg.argText !== "") {
        wantsArgPreview = true;
        activeArgForPreview = activeArg;
        // Key preview state off the draft version + argument identity. `draftVersion` is bumped on any
        // draft text change, so cursor moves within the same argument do not allocate/copy the full
        // argument text into a key string.
        argPreviewKey = `${draftVersion}|${activeArg.fnName}|${activeArg.argIndex}|${activeArg.span.start}:${activeArg.span.end}`;
        if (this.#argumentPreviewKey !== argPreviewKey) {
          this.#argumentPreviewKey = argPreviewKey;
          this.#argumentPreviewDisplayKey = argPreviewKey;
          this.#argumentPreviewDisplayExpr = formatArgumentPreviewExpression(activeArg.argText);
          this.#argumentPreviewValue = null;
          this.#argumentPreviewDisplayValue = null;
          this.#argumentPreviewPending = true;
          this.#scheduleArgumentPreviewEvaluation(activeArg, argPreviewKey);
        }
      } else {
        this.#clearArgumentPreviewState();
      }
    }

    const syntaxKey = syntaxError ? `${syntaxError.span?.start ?? ""}:${syntaxError.span?.end ?? ""}:${syntaxError.message}` : "";
    const nextHintIsFormulaEditing = isFormulaEditing;
    const nextHint = hint;
    const nextArgPreviewKey = wantsArgPreview ? argPreviewKey : null;
    const nextArgPreviewRhs = wantsArgPreview
      ? this.#argumentPreviewPending
        ? "…"
        : (this.#argumentPreviewDisplayValue ?? formatArgumentPreviewValue(this.#argumentPreviewValue))
      : null;

    const hintChanged =
      nextHintIsFormulaEditing !== this.#lastHintIsFormulaEditing ||
      syntaxKey !== this.#lastHintSyntaxKey ||
      nextHint !== this.#lastHint ||
      nextArgPreviewKey !== this.#lastHintArgPreviewKey ||
      nextArgPreviewRhs !== this.#lastHintArgPreviewRhs;

    if (hintChanged) {
      this.#lastHintIsFormulaEditing = nextHintIsFormulaEditing;
      this.#lastHintSyntaxKey = syntaxKey;
      this.#lastHint = nextHint;
      this.#lastHintArgPreviewKey = nextArgPreviewKey;
      this.#lastHintArgPreviewRhs = nextArgPreviewRhs;
      this.#hintEl.replaceChildren();

      if (syntaxError) {
        const message = document.createElement("div");
        message.className = "formula-bar-hint-error";
        message.textContent = syntaxError.message;
        this.#hintEl.appendChild(message);
      }

      if (hint) {
        const panel = document.createElement("div");
        panel.className = "formula-bar-hint-panel";

        const title = document.createElement("div");
        title.className = "formula-bar-hint-title";
        title.textContent = "PARAMETERS";

        const body = document.createElement("div");
        body.className = "formula-bar-hint-body";

        const signature = document.createElement("span");
        signature.className = "formula-bar-hint-signature";

        for (const part of hint.parts) {
          const token = document.createElement("span");
          token.className = `formula-bar-hint-token formula-bar-hint-token--${part.kind}`;
          token.dataset.kind = part.kind;
          token.textContent = part.text;
          signature.appendChild(token);
        }

        body.appendChild(signature);

        const summary = hint.signature.summary?.trim?.() ?? "";
        if (summary) {
          const sep = document.createElement("span");
          sep.className = "formula-bar-hint-summary-separator";
          sep.textContent = " — ";

          const summaryEl = document.createElement("span");
          summaryEl.className = "formula-bar-hint-summary";
          summaryEl.textContent = summary;

          body.appendChild(sep);
          body.appendChild(summaryEl);
        }

        if (wantsArgPreview && activeArgForPreview) {
          const previewEl = document.createElement("div");
          previewEl.className = "formula-bar-hint-arg-preview";
          previewEl.dataset.testid = "formula-hint-arg-preview";
          previewEl.dataset.argStart = String(activeArgForPreview.span.start);
          previewEl.dataset.argEnd = String(activeArgForPreview.span.end);

          const rhs = nextArgPreviewRhs ?? "";
          const displayArgText =
            this.#argumentPreviewDisplayKey === argPreviewKey
              ? (this.#argumentPreviewDisplayExpr ?? "")
              : formatArgumentPreviewExpression(activeArgForPreview.argText);
          previewEl.textContent = `↳ ${displayArgText}  →  ${rhs}`;
          body.appendChild(previewEl);
        } else {
          this.#clearArgumentPreviewState();
        }
        panel.appendChild(title);
        panel.appendChild(body);
        this.#hintEl.appendChild(panel);
      }
    }

    const explanation = this.model.errorExplanation();
    const address = this.model.activeCell.address;
    if (explanation !== this.#lastErrorExplanation || address !== this.#lastErrorExplanationAddress) {
      this.#lastErrorExplanation = explanation;
      this.#lastErrorExplanationAddress = address;
      if (!explanation) {
        this.root.classList.toggle("formula-bar--has-error", false);
        this.#errorButton.hidden = true;
        this.#errorButton.disabled = true;
        this.#errorTitleEl.textContent = "";
        this.#errorDescEl.textContent = "";
        this.#errorSuggestionsEl.replaceChildren();
        this.#setErrorPanelOpen(false, { restoreFocus: false });
      } else {
        this.root.classList.toggle("formula-bar--has-error", true);
        this.#errorButton.hidden = false;
        this.#errorButton.disabled = false;
        this.#errorTitleEl.textContent = `${explanation.code} (${address}): ${explanation.title}`;
        this.#errorDescEl.textContent = explanation.description;
        this.#errorSuggestionsEl.replaceChildren(
          ...explanation.suggestions.map((s) => {
            const li = document.createElement("li");
            li.textContent = s;
            return li;
          })
        );
      }
    }

    this.#syncErrorPanelActions(explanation);

    this.#syncScroll();
    this.#adjustHeight();
  }

  #referenceElementsForIndex(idx: number): HTMLElement[] {
    if (this.#referenceElsByIndex == null) {
      // When cursoring through a formula with many references, `querySelectorAll` per-ref-index
      // can become noticeable. Build a full index->elements map once per highlight DOM update.
      const buckets: Array<HTMLElement[] | undefined> = [];
      const els = this.#highlightEl.querySelectorAll<HTMLElement>("[data-ref-index]");
      for (const el of els) {
        const raw = el.dataset.refIndex;
        if (!raw) continue;
        const parsed = Number(raw);
        if (!Number.isInteger(parsed) || parsed < 0) continue;
        const bucket = buckets[parsed];
        if (bucket) bucket.push(el);
        else buckets[parsed] = [el];
      }
      this.#referenceElsByIndex = buckets;
    }
    return this.#referenceElsByIndex[idx] ?? EMPTY_REFERENCE_ELS;
  }

  #clearArgumentPreviewState(): void {
    if (
      this.#argumentPreviewKey === null &&
      this.#argumentPreviewDisplayKey === null &&
      this.#argumentPreviewDisplayExpr === null &&
      this.#argumentPreviewValue === null &&
      this.#argumentPreviewDisplayValue === null &&
      this.#argumentPreviewTimer == null &&
      !this.#argumentPreviewPending
    ) {
      // Already cleared; avoid bumping the request id on every render when no hint/preview is shown.
      return;
    }
    this.#argumentPreviewKey = null;
    this.#argumentPreviewDisplayKey = null;
    this.#argumentPreviewDisplayExpr = null;
    this.#argumentPreviewValue = null;
    this.#argumentPreviewDisplayValue = null;
    this.#argumentPreviewPending = false;
    this.#argumentPreviewRequestId += 1;
    if (this.#argumentPreviewTimer != null) {
      clearTimeout(this.#argumentPreviewTimer);
      this.#argumentPreviewTimer = null;
    }
  }

  #scheduleArgumentPreviewEvaluation(activeArg: ReturnType<FormulaBarModel["activeArgumentSpan"]>, key: string): void {
    if (!activeArg) return;
    const provider = this.#argumentPreviewProvider;
    if (typeof provider !== "function") return;

    // Cancel any pending timer before scheduling a new evaluation. This keeps typing responsive
    // and avoids doing work for stale cursor positions.
    if (this.#argumentPreviewTimer != null) {
      clearTimeout(this.#argumentPreviewTimer);
      this.#argumentPreviewTimer = null;
    }

    const requestId = ++this.#argumentPreviewRequestId;
    const expr = activeArg.argText;

    this.#argumentPreviewTimer = setTimeout(() => {
      this.#argumentPreviewTimer = null;

      // Allow the preview provider to be async, but bound the time we wait for it.
      const timeoutMs = 100;
      let timeoutId: ReturnType<typeof setTimeout> | null = null;
      const timeoutPromise = new Promise<unknown>((resolve) => {
        timeoutId = setTimeout(() => resolve("(preview unavailable)"), timeoutMs);
      });

      Promise.race([Promise.resolve().then(() => provider(expr)), timeoutPromise])
        .then((value) => {
          if (timeoutId != null) clearTimeout(timeoutId);
          if (requestId !== this.#argumentPreviewRequestId) return;
          if (this.#argumentPreviewKey !== key) return;
          this.#argumentPreviewPending = false;
          this.#argumentPreviewValue = value === undefined ? "(preview unavailable)" : value;
          this.#argumentPreviewDisplayValue = formatArgumentPreviewValue(this.#argumentPreviewValue);
          this.#requestRender({ preserveTextareaValue: true });
        })
        .catch(() => {
          if (timeoutId != null) clearTimeout(timeoutId);
          if (requestId !== this.#argumentPreviewRequestId) return;
          if (this.#argumentPreviewKey !== key) return;
          this.#argumentPreviewPending = false;
          this.#argumentPreviewValue = "(preview unavailable)";
          this.#argumentPreviewDisplayValue = formatArgumentPreviewValue(this.#argumentPreviewValue);
          this.#requestRender({ preserveTextareaValue: true });
        });
    }, 0);
  }

  #setErrorPanelOpen(open: boolean, opts: { restoreFocus: boolean } = { restoreFocus: true }): void {
    const wasOpen = this.#isErrorPanelOpen;
    this.#isErrorPanelOpen = open;
    this.root.classList.toggle("formula-bar--error-panel-open", open);
    this.#errorButton.setAttribute("aria-expanded", open ? "true" : "false");
    this.#errorPanel.hidden = !open;

    if (!open) {
      const hadReferenceHighlights = this.#errorPanelReferenceHighlights != null;
      this.#errorPanelReferenceHighlights = null;
      this.#syncErrorPanelActions();
      // Clear view-mode highlights; preserve formula-editing highlights.
      if (hadReferenceHighlights) {
        this.#emitOverlays();
      }
      if (opts.restoreFocus) {
        try {
          this.#errorButton.focus({ preventScroll: true });
        } catch {
          this.#errorButton.focus();
        }
      }
      return;
    }

    this.#syncErrorPanelActions();

    if (!wasOpen) {
      this.#focusFirstErrorPanelControl();
    }
  }

  #adjustHeight(): void {
    const minHeight = FORMULA_BAR_MIN_HEIGHT;
    const maxHeight = this.#isExpanded ? FORMULA_BAR_MAX_HEIGHT_EXPANDED : FORMULA_BAR_MAX_HEIGHT_COLLAPSED;

    // Measuring `scrollHeight` can trigger layout and become noticeable on very long formulas.
    // Skip redundant re-measurements when only the selection/cursor changed.
    const draft = this.model.draft;
    if (
      this.#lastAdjustedHeightDraft === draft &&
      this.#lastAdjustedHeightIsEditing === this.model.isEditing &&
      this.#lastAdjustedHeightIsExpanded === this.#isExpanded
    ) {
      return;
    }

    this.#lastAdjustedHeightDraft = draft;
    this.#lastAdjustedHeightIsEditing = this.model.isEditing;
    this.#lastAdjustedHeightIsExpanded = this.#isExpanded;

    const highlightEl = this.#highlightEl as HTMLElement;

    if (this.model.isEditing) {
      // Reset before measuring.
      this.textarea.style.height = `${minHeight}px`;
      const desired = Math.max(minHeight, Math.min(maxHeight, this.textarea.scrollHeight));
      this.textarea.style.height = `${desired}px`;
      highlightEl.style.height = `${desired}px`;
      return;
    }

    // In view mode, the textarea is hidden, so measure the highlighted <pre>.
    highlightEl.style.height = `${minHeight}px`;
    const desired = Math.max(minHeight, Math.min(maxHeight, highlightEl.scrollHeight));
    highlightEl.style.height = `${desired}px`;
  }

  #toggleExpanded(): void {
    this.#isExpanded = !this.#isExpanded;
    storeFormulaBarExpandedState(this.#isExpanded);
    this.#syncExpandedUi();
    this.#adjustHeight();
  }

  #syncExpandedUi(): void {
    this.root.classList.toggle("formula-bar--expanded", this.#isExpanded);
    this.#expandButtonEl.textContent = this.#isExpanded ? "▴" : "▾";
    const label = this.#isExpanded ? "Collapse formula bar" : "Expand formula bar";
    this.#expandButtonEl.title = label;
    this.#expandButtonEl.setAttribute("aria-label", label);
    this.#expandButtonEl.setAttribute("aria-pressed", this.#isExpanded ? "true" : "false");
  }

  #syncScroll(): void {
    const highlightEl = this.#highlightEl as HTMLElement;
    const nextTop = this.textarea.scrollTop;
    const nextLeft = this.textarea.scrollLeft;
    // Writing scroll positions can trigger additional work in the browser; only sync when
    // the underlying values changed (or when a highlight re-render reset scroll state).
    if (highlightEl.scrollTop !== nextTop) highlightEl.scrollTop = nextTop;
    if (highlightEl.scrollLeft !== nextLeft) highlightEl.scrollLeft = nextLeft;
  }

  #setTextareaSelectionFromModel(): void {
    if (!this.model.isEditing) return;
    const start = this.model.cursorStart;
    const end = this.model.cursorEnd;
    this.textarea.setSelectionRange(start, end);
  }

  #emitOverlays(): void {
    const range = this.#hoverOverride ?? this.model.hoveredReference();
    const refText = this.#hoverOverrideText ?? this.model.hoveredReferenceText();
    const normalizedText = refText ?? null;

    let hoverChanged = false;
    if (!range) {
      hoverChanged = this.#lastEmittedHoverRange !== null;
    } else {
      const prev = this.#lastEmittedHoverRange;
      const startRow = range.start.row;
      const startCol = range.start.col;
      const endRow = range.end.row;
      const endCol = range.end.col;
      if (!prev || prev.startRow !== startRow || prev.startCol !== startCol || prev.endRow !== endRow || prev.endCol !== endCol) {
        hoverChanged = true;
      }
      if (hoverChanged) this.#lastEmittedHoverRange = { startRow, startCol, endRow, endCol };
    }
    if (!range && hoverChanged) this.#lastEmittedHoverRange = null;
    if (this.#lastEmittedHoverText !== normalizedText) hoverChanged = true;

    if (hoverChanged) {
      this.#lastEmittedHoverText = normalizedText;
      this.#callbacks.onHoverRange?.(range);
      this.#callbacks.onHoverRangeWithText?.(range, normalizedText);
    }

    // Reference highlight overlay updates can be costly (e.g. SpreadsheetApp recomputes/filters highlights).
    // Only emit when the underlying highlights actually changed.
    const isFormula = isFormulaText(this.model.draft);
    const nextMode: ReferenceHighlightMode =
      this.model.isEditing && isFormula ? "editing" : this.#errorPanelReferenceHighlights ? "errorPanel" : "none";
    let highlightsChanged = nextMode !== this.#lastEmittedReferenceHighlightsMode;

    if (nextMode === "editing") {
      const colored = this.model.coloredReferences();
      const active = this.model.activeReferenceIndex();
      if (colored !== this.#lastEmittedReferenceHighlightsColoredRefs || active !== this.#lastEmittedReferenceHighlightsActiveIndex) {
        highlightsChanged = true;
      }
      if (highlightsChanged) {
        this.#lastEmittedReferenceHighlightsColoredRefs = colored;
        this.#lastEmittedReferenceHighlightsActiveIndex = active;
        this.#lastEmittedReferenceHighlightsErrorPanel = null;
      }
    } else if (nextMode === "errorPanel") {
      if (this.#errorPanelReferenceHighlights !== this.#lastEmittedReferenceHighlightsErrorPanel) {
        highlightsChanged = true;
      }
      if (highlightsChanged) {
        this.#lastEmittedReferenceHighlightsColoredRefs = null;
        this.#lastEmittedReferenceHighlightsActiveIndex = null;
        this.#lastEmittedReferenceHighlightsErrorPanel = this.#errorPanelReferenceHighlights;
      }
    } else {
      if (highlightsChanged) {
        this.#lastEmittedReferenceHighlightsColoredRefs = null;
        this.#lastEmittedReferenceHighlightsActiveIndex = null;
        this.#lastEmittedReferenceHighlightsErrorPanel = null;
      }
    }

    if (highlightsChanged) {
      this.#lastEmittedReferenceHighlightsMode = nextMode;
      this.#callbacks.onReferenceHighlights?.(this.#currentReferenceHighlights());
    }
  }

  #onHighlightHover(e: MouseEvent): void {
    if (this.model.isEditing) return;
    const target = e.target as HTMLElement | null;
    const span = target?.closest?.("span[data-kind]") as HTMLElement | null;
    if (!span) {
      this.#clearHoverOverride();
      return;
    }
    const kind = span.dataset.kind;
    const text = span.textContent ?? "";
    if (!text) {
      this.#clearHoverOverride();
      return;
    }

    if (kind === "reference") {
      // `#hoverOverride` can legitimately be null (unresolvable reference). Use `undefined` as
      // the "no cache" sentinel so we can reuse cached nulls across mousemove events.
      const cachedRange = this.#hoverOverrideText === text ? this.#hoverOverride : undefined;
      const nextRange = cachedRange === undefined ? this.model.resolveReferenceText(text) : cachedRange;
      const prevRange = this.#lastEmittedHoverRange;
      const sameRange =
        nextRange == null
          ? prevRange == null
          : prevRange != null &&
            prevRange.startRow === nextRange.start.row &&
            prevRange.startCol === nextRange.start.col &&
            prevRange.endRow === nextRange.end.row &&
            prevRange.endCol === nextRange.end.col;

      // `mousemove` can fire repeatedly while still over the same token; avoid
      // re-emitting identical hover previews.
      //
      // Note: some consumers (e.g. SpreadsheetApp's range preview tooltip) may want to
      // refresh derived UI when the underlying document changes even if the hovered
      // reference span is unchanged. Still emit the text-aware callback so consumers
      // can apply their own caching (e.g. by document version).
      if (this.#lastEmittedHoverText === text && sameRange) {
        this.#hoverOverrideText = text;
        this.#hoverOverride = nextRange;
        // Even when the hovered token/range hasn't changed, allow consumers to refresh any
        // value-dependent UI (e.g. SpreadsheetApp's range preview tooltip, which keys off a
        // monotonic document version). Avoid re-emitting the range-outline preview unless the
        // embedding app doesn't provide the richer `onHoverRangeWithText` callback.
        if (typeof this.#callbacks.onHoverRangeWithText === "function") {
          this.#callbacks.onHoverRangeWithText(this.#hoverOverride, this.#hoverOverrideText);
        } else {
          this.#callbacks.onHoverRange?.(this.#hoverOverride);
        }
        return;
      }

      this.#hoverOverrideText = text;
      this.#hoverOverride = nextRange;
      this.#callbacks.onHoverRange?.(this.#hoverOverride);
      this.#callbacks.onHoverRangeWithText?.(this.#hoverOverride, this.#hoverOverrideText);
      // Keep overlay emission caches in sync; hover previews in view mode are emitted directly
      // (not via `#emitOverlays`) to avoid recomputing reference highlight overlays on every mousemove.
      this.#lastEmittedHoverText = this.#hoverOverrideText;
      this.#lastEmittedHoverRange = this.#hoverOverride
        ? {
            startRow: this.#hoverOverride.start.row,
            startCol: this.#hoverOverride.start.col,
            endRow: this.#hoverOverride.end.row,
            endCol: this.#hoverOverride.end.col,
          }
        : null;
      return;
    }

    if (kind === "identifier") {
      const cachedRange = this.#hoverOverrideText === text ? this.#hoverOverride : undefined;
      let nextRange: RangeAddress | null = cachedRange === undefined ? null : cachedRange;
      if (cachedRange === undefined) {
        const resolved = this.model.resolveNameRange(text);
        if (!resolved) {
          this.#clearHoverOverride();
          return;
        }

        nextRange = {
          start: { row: resolved.startRow, col: resolved.startCol },
          end: { row: resolved.endRow, col: resolved.endCol },
        } satisfies RangeAddress;
      }

      const prevRange = this.#lastEmittedHoverRange;
      const sameRange =
        prevRange != null &&
        prevRange.startRow === nextRange.start.row &&
        prevRange.startCol === nextRange.start.col &&
        prevRange.endRow === nextRange.end.row &&
        prevRange.endCol === nextRange.end.col;

       if (this.#lastEmittedHoverText === text && sameRange) {
         this.#hoverOverrideText = text;
         this.#hoverOverride = nextRange;
         if (typeof this.#callbacks.onHoverRangeWithText === "function") {
           this.#callbacks.onHoverRangeWithText(this.#hoverOverride, this.#hoverOverrideText);
         } else {
           this.#callbacks.onHoverRange?.(this.#hoverOverride);
         }
         return;
       }

      this.#hoverOverrideText = text;
      this.#hoverOverride = nextRange;
      this.#callbacks.onHoverRange?.(this.#hoverOverride);
      this.#callbacks.onHoverRangeWithText?.(this.#hoverOverride, this.#hoverOverrideText);
      this.#lastEmittedHoverText = this.#hoverOverrideText;
      this.#lastEmittedHoverRange = this.#hoverOverride
        ? {
            startRow: this.#hoverOverride.start.row,
            startCol: this.#hoverOverride.start.col,
            endRow: this.#hoverOverride.end.row,
            endCol: this.#hoverOverride.end.col,
          }
        : null;
      return;
    }

    this.#clearHoverOverride();
  }

  #clearHoverOverride(): void {
    if (this.#hoverOverride === null && this.#hoverOverrideText === null) return;
    this.#hoverOverride = null;
    this.#hoverOverrideText = null;
    this.#emitOverlays();
  }

  #inferSelectedReferenceIndex(start: number, end: number): number | null {
    if (!this.model.isEditing || !isFormulaText(this.model.draft)) return null;
    if (start === end) return null;
    const refs = this.model.coloredReferences();
    let lo = 0;
    let hi = refs.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const ref = refs[mid]!;
      if (ref.start < start) {
        lo = mid + 1;
        continue;
      }
      if (ref.start > start) {
        hi = mid - 1;
        continue;
      }
      return ref.end === end ? ref.index : null;
    }
    return null;
  }

  #fixFormulaErrorWithAi(): void {
    const explanation = this.model.errorExplanation();
    if (!explanation) return;
    this.#callbacks.onFixFormulaErrorWithAi?.({
      address: this.model.activeCell.address,
      input: this.model.activeCell.input,
      draft: this.model.draft,
      value: this.model.activeCell.value,
      explanation
    });
  }

  #toggleErrorReferenceHighlights(): void {
    if (this.#errorPanelReferenceHighlights) {
      this.#errorPanelReferenceHighlights = null;
    } else {
      this.#errorPanelReferenceHighlights = computeReferenceHighlights(
        this.model.draft,
        this.model.extractFormulaReferencesOptions()
      );
      if (this.#errorPanelReferenceHighlights.length === 0) {
        this.#errorPanelReferenceHighlights = null;
      }
    }

    this.#syncErrorPanelActions();
    this.#emitOverlays();
  }

  #syncErrorPanelActions(explanation?: ReturnType<FormulaBarModel["errorExplanation"]> | null): void {
    const resolved = explanation === undefined ? this.model.errorExplanation() : explanation;
    const canFix = Boolean(resolved) && typeof this.#callbacks.onFixFormulaErrorWithAi === "function";
    const fixDisabled = !canFix;
    if (this.#lastErrorFixAiDisabled !== fixDisabled) {
      this.#lastErrorFixAiDisabled = fixDisabled;
      this.#errorFixAiButton.disabled = fixDisabled;
    }

    const isFormula = isFormulaText(this.model.draft);
    const isShowingRanges = this.#errorPanelReferenceHighlights != null;
    const showRangesDisabled = !isFormula;
    if (this.#lastErrorShowRangesDisabled !== showRangesDisabled) {
      this.#lastErrorShowRangesDisabled = showRangesDisabled;
      this.#errorShowRangesButton.disabled = showRangesDisabled;
    }

    if (this.#lastErrorShowRangesPressed !== isShowingRanges) {
      this.#lastErrorShowRangesPressed = isShowingRanges;
      this.#errorShowRangesButton.setAttribute("aria-pressed", isShowingRanges ? "true" : "false");
    }

    const showRangesText = isShowingRanges ? "Hide referenced ranges" : "Show referenced ranges";
    if (this.#lastErrorShowRangesText !== showRangesText) {
      this.#lastErrorShowRangesText = showRangesText;
      this.#errorShowRangesButton.textContent = showRangesText;
    }
  }

  #currentReferenceHighlights(): FormulaReferenceHighlight[] {
    const isFormula = isFormulaText(this.model.draft);
    if (this.model.isEditing && isFormula) {
      return this.model.referenceHighlights();
    }
    if (this.#errorPanelReferenceHighlights) {
      return this.#errorPanelReferenceHighlights;
    }
    return [];
  }

  #focusFirstErrorPanelControl(): void {
    const candidates: HTMLElement[] = [];
    if (!this.#errorFixAiButton.disabled) candidates.push(this.#errorFixAiButton);
    if (!this.#errorShowRangesButton.disabled) candidates.push(this.#errorShowRangesButton);
    candidates.push(this.#errorCloseButton);

    const target = candidates.find((el) => !el.hidden && !(el as HTMLButtonElement).disabled) ?? null;
    if (!target) return;
    try {
      target.focus({ preventScroll: true });
    } catch {
      target.focus();
    }
  }

  #onErrorPanelKeyDown(e: KeyboardEvent): void {
    if (!this.#isErrorPanelOpen) return;

    if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      this.#setErrorPanelOpen(false, { restoreFocus: true });
      return;
    }

    if (e.key !== "Tab") return;

    const focusable = [this.#errorFixAiButton, this.#errorShowRangesButton, this.#errorCloseButton].filter(
      (el) => !el.hidden && !el.disabled
    );
    if (focusable.length === 0) return;

    const active = document.activeElement as HTMLElement | null;
    const currentIdx = active ? focusable.indexOf(active as HTMLButtonElement) : -1;
    const nextIdx = (() => {
      if (e.shiftKey) {
        if (currentIdx <= 0) return focusable.length - 1;
        return currentIdx - 1;
      }
      if (currentIdx < 0 || currentIdx === focusable.length - 1) return 0;
      return currentIdx + 1;
    })();
    const next = focusable[nextIdx]!;
    e.preventDefault();
    try {
      next.focus({ preventScroll: true });
    } catch {
      next.focus();
    }
  }
  #openNameBoxDropdown(): void {
    if (this.#isNameBoxDropdownOpen) return;
    if (!this.#nameBoxDropdownProvider) {
      // Still focus the address input so keyboard "Go To" still feels natural.
      this.#addressEl.focus();
      return;
    }

    this.#isNameBoxDropdownOpen = true;
    this.#nameBoxDropdownOriginalAddressValue = this.#addressEl.value;

    const rawItems = this.#nameBoxDropdownProvider.getItems();
    const baseItems = Array.isArray(rawItems) ? rawItems.slice() : [];
    const baseByKey = new Map(baseItems.map((item) => [item.key, item]));

    const recentItems: NameBoxDropdownItem[] = [];
    const recentKeySet = new Set<string>();
    for (const key of this.#nameBoxDropdownRecentKeys) {
      const item = baseByKey.get(key);
      if (!item) continue;
      recentKeySet.add(key);
      recentItems.push({ ...item, kind: "recent" });
    }
    const nonRecentItems = baseItems.filter((item) => !recentKeySet.has(item.key));
    this.#nameBoxDropdownAllItems = [...recentItems, ...nonRecentItems];

    // Default sort: keep groups stable, sort labels within each group.
    const kindOrder: Record<NameBoxDropdownItemKind, number> = {
      recent: 0,
      namedRange: 1,
      table: 2,
    };
    const recentRank = new Map(this.#nameBoxDropdownRecentKeys.map((key, index) => [key, index]));
    this.#nameBoxDropdownAllItems.sort((a, b) => {
      const ak = kindOrder[a.kind] ?? 99;
      const bk = kindOrder[b.kind] ?? 99;
      if (ak !== bk) return ak - bk;
      if (a.kind === "recent" && b.kind === "recent") {
        return (recentRank.get(a.key) ?? 99) - (recentRank.get(b.key) ?? 99);
      }
      return a.label.localeCompare(b.label, undefined, { sensitivity: "base" });
    });

    this.#nameBoxDropdownPopupEl.hidden = false;
    this.#addressEl.setAttribute("aria-expanded", "true");
    this.#addressEl.setAttribute("aria-controls", this.#nameBoxDropdownListEl.id);
    this.#nameBoxDropdownEl.setAttribute("aria-expanded", "true");
    this.#nameBoxDropdownEl.setAttribute("aria-controls", this.#nameBoxDropdownListEl.id);

    this.#updateNameBoxDropdownFilter("");
    this.#positionNameBoxDropdown();
    this.#attachNameBoxDropdownGlobalListeners();

    try {
      this.#addressEl.focus({ preventScroll: true });
    } catch {
      this.#addressEl.focus();
    }
    this.#addressEl.select();
  }

  #positionNameBoxDropdown(): void {
    const anchor = this.#nameBoxEl.getBoundingClientRect();
    const margin = 8;
    const gap = 4;

    // Seed position/width so we can measure the popup.
    this.#nameBoxDropdownPopupEl.style.left = `${anchor.left}px`;
    this.#nameBoxDropdownPopupEl.style.top = `${anchor.bottom + gap}px`;
    if (anchor.width > 0) {
      this.#nameBoxDropdownPopupEl.style.minWidth = `${anchor.width}px`;
    }

    const rect = this.#nameBoxDropdownPopupEl.getBoundingClientRect();

    let left = anchor.left;
    let top = anchor.bottom + gap;

    // Prefer opening downward, but flip upward if we would overflow.
    if (top + rect.height + margin > window.innerHeight && anchor.top - rect.height - gap > margin) {
      top = anchor.top - rect.height - gap;
    }

    // Clamp horizontally.
    if (left + rect.width + margin > window.innerWidth) {
      left = window.innerWidth - rect.width - margin;
    }

    left = Math.max(margin, left);
    top = Math.max(margin, top);

    this.#nameBoxDropdownPopupEl.style.left = `${left}px`;
    this.#nameBoxDropdownPopupEl.style.top = `${top}px`;
  }

  #closeNameBoxDropdown(opts: { restoreAddress: boolean; reason: "escape" | "toggle" | "outside" | "resize" | "scroll" | "commit" }): void {
    if (!this.#isNameBoxDropdownOpen) return;
    this.#isNameBoxDropdownOpen = false;

    if (opts.restoreAddress && this.#nameBoxDropdownOriginalAddressValue != null) {
      this.#addressEl.value = this.#nameBoxDropdownOriginalAddressValue;
    }

    this.#nameBoxDropdownOriginalAddressValue = null;
    this.#nameBoxDropdownActiveIndex = -1;
    this.#nameBoxDropdownQuery = "";
    this.#nameBoxDropdownFilteredItems = [];
    this.#nameBoxDropdownOptionEls = [];
    this.#nameBoxDropdownPopupEl.hidden = true;
    this.#nameBoxDropdownListEl.replaceChildren();

    this.#addressEl.setAttribute("aria-expanded", "false");
    this.#addressEl.removeAttribute("aria-controls");
    this.#addressEl.removeAttribute("aria-activedescendant");
    this.#nameBoxDropdownEl.setAttribute("aria-expanded", "false");
    this.#nameBoxDropdownEl.removeAttribute("aria-controls");

    this.#detachNameBoxDropdownGlobalListeners();
  }

  #attachNameBoxDropdownGlobalListeners(): void {
    this.#detachNameBoxDropdownGlobalListeners();

    this.#nameBoxDropdownPointerDownListener = (e: PointerEvent) => {
      if (!this.#isNameBoxDropdownOpen) return;
      const target = e.target as Node | null;
      if (!target) return;
      if (this.#nameBoxDropdownPopupEl.contains(target)) return;
      if (this.#nameBoxEl.contains(target)) return;
      this.#closeNameBoxDropdown({ restoreAddress: true, reason: "outside" });
    };
    window.addEventListener("pointerdown", this.#nameBoxDropdownPointerDownListener, { capture: true, signal: this.#domAbort.signal });

    this.#nameBoxDropdownFocusInListener = (e: FocusEvent) => {
      if (!this.#isNameBoxDropdownOpen) return;
      const target = e.target as Node | null;
      if (!target) return;
      if (this.#nameBoxDropdownPopupEl.contains(target)) return;
      if (this.#nameBoxEl.contains(target)) return;
      this.#closeNameBoxDropdown({ restoreAddress: true, reason: "outside" });
    };
    document.addEventListener("focusin", this.#nameBoxDropdownFocusInListener, { capture: true, signal: this.#domAbort.signal });

    this.#nameBoxDropdownScrollListener = (e: Event) => {
      if (!this.#isNameBoxDropdownOpen) return;
      const target = e.target as Node | null;
      if (target && this.#nameBoxDropdownPopupEl.contains(target)) return;
      this.#closeNameBoxDropdown({ restoreAddress: true, reason: "scroll" });
    };
    window.addEventListener("scroll", this.#nameBoxDropdownScrollListener, { capture: true, signal: this.#domAbort.signal });

    this.#nameBoxDropdownResizeListener = () => {
      if (!this.#isNameBoxDropdownOpen) return;
      this.#closeNameBoxDropdown({ restoreAddress: true, reason: "resize" });
    };
    window.addEventListener("resize", this.#nameBoxDropdownResizeListener, { signal: this.#domAbort.signal });

    this.#nameBoxDropdownBlurListener = () => {
      if (!this.#isNameBoxDropdownOpen) return;
      this.#closeNameBoxDropdown({ restoreAddress: true, reason: "outside" });
    };
    window.addEventListener("blur", this.#nameBoxDropdownBlurListener, { signal: this.#domAbort.signal });
  }

  #detachNameBoxDropdownGlobalListeners(): void {
    if (this.#nameBoxDropdownPointerDownListener) {
      window.removeEventListener("pointerdown", this.#nameBoxDropdownPointerDownListener, true);
      this.#nameBoxDropdownPointerDownListener = null;
    }
    if (this.#nameBoxDropdownFocusInListener) {
      document.removeEventListener("focusin", this.#nameBoxDropdownFocusInListener, true);
      this.#nameBoxDropdownFocusInListener = null;
    }
    if (this.#nameBoxDropdownScrollListener) {
      window.removeEventListener("scroll", this.#nameBoxDropdownScrollListener, true);
      this.#nameBoxDropdownScrollListener = null;
    }
    if (this.#nameBoxDropdownResizeListener) {
      window.removeEventListener("resize", this.#nameBoxDropdownResizeListener);
      this.#nameBoxDropdownResizeListener = null;
    }
    if (this.#nameBoxDropdownBlurListener) {
      window.removeEventListener("blur", this.#nameBoxDropdownBlurListener);
      this.#nameBoxDropdownBlurListener = null;
    }
  }

  #updateNameBoxDropdownFilter(rawQuery: string): void {
    const query = String(rawQuery ?? "").trim().toLowerCase();
    this.#nameBoxDropdownQuery = query;
    const all = this.#nameBoxDropdownAllItems;
    const filtered =
      query === ""
        ? all.slice()
        : all.filter((item) => {
            const label = String(item.label ?? "").toLowerCase();
            const ref = String(item.reference ?? "").toLowerCase();
            return label.startsWith(query) || ref.startsWith(query) || label.includes(query);
          });

    this.#nameBoxDropdownFilteredItems = filtered;
    this.#renderNameBoxDropdownList();
  }

  #renderNameBoxDropdownList(): void {
    const list = this.#nameBoxDropdownListEl;
    list.replaceChildren();
    this.#nameBoxDropdownOptionEls = [];

    const groups = new Map<NameBoxDropdownItemKind, NameBoxDropdownItem[]>();
    for (const item of this.#nameBoxDropdownFilteredItems) {
      const arr = groups.get(item.kind) ?? [];
      arr.push(item);
      groups.set(item.kind, arr);
    }

    const renderGroup = (kind: NameBoxDropdownItemKind, label: string): void => {
      const items = groups.get(kind) ?? [];
      if (items.length === 0) return;

      const group = document.createElement("div");
      group.className = "formula-bar-name-box-group";
      group.setAttribute("role", "group");

      const heading = document.createElement("div");
      heading.className = "formula-bar-name-box-group-label";
      heading.textContent = label;
      heading.id = `${list.id}-group-${kind}`;
      group.setAttribute("aria-labelledby", heading.id);
      group.appendChild(heading);

      for (const item of items) {
        const option = document.createElement("div");
        option.className = "formula-bar-name-box-option";
        option.setAttribute("role", "option");
        option.id = this.#nameBoxDropdownOptionId(item);
        option.dataset.key = item.key;

        const main = document.createElement("div");
        main.className = "formula-bar-name-box-option-main";

        const labelEl = document.createElement("div");
        labelEl.className = "formula-bar-name-box-option-label";
        labelEl.textContent = item.label;
        main.appendChild(labelEl);

        if (item.description) {
          const descEl = document.createElement("div");
          descEl.className = "formula-bar-name-box-option-description";
          descEl.textContent = item.description;
          main.appendChild(descEl);
        }

        option.appendChild(main);

        const index = this.#nameBoxDropdownOptionEls.length;
        this.#nameBoxDropdownOptionEls.push(option);

        option.addEventListener(
          "mousemove",
          () => {
            this.#setNameBoxDropdownActiveIndex(index);
          },
          { signal: this.#domAbort.signal },
        );
        option.addEventListener(
          "mousedown",
          (e) => {
            // Avoid selecting the underlying input text.
            e.preventDefault();
          },
          { signal: this.#domAbort.signal },
        );
        option.addEventListener(
          "click",
          () => {
            const chosen = this.#nameBoxDropdownFilteredItems[index] ?? null;
            if (!chosen) return;
            this.#selectNameBoxDropdownItem(chosen);
          },
          { signal: this.#domAbort.signal },
        );

        group.appendChild(option);
      }

      list.appendChild(group);
    };

    // Match the sort order used for `#nameBoxDropdownAllItems`.
    renderGroup("recent", "Recent");
    renderGroup("namedRange", "Named ranges");
    renderGroup("table", "Tables");

    if (this.#nameBoxDropdownOptionEls.length === 0) {
      const empty = document.createElement("div");
      empty.className = "formula-bar-name-box-empty";
      // Match Excel semantics: an empty dropdown still renders a disabled item rather than
      // being completely blank, so keyboard users get feedback about why nothing is selectable.
      empty.setAttribute("role", "option");
      empty.setAttribute("aria-disabled", "true");
      empty.setAttribute("aria-selected", "false");
      empty.textContent =
        this.#nameBoxDropdownAllItems.length === 0 && this.#nameBoxDropdownQuery === "" ? "No named ranges" : "No matches";
      list.appendChild(empty);
      this.#setNameBoxDropdownActiveIndex(-1);
      return;
    }

    const nextActive = Math.min(Math.max(this.#nameBoxDropdownActiveIndex, 0), this.#nameBoxDropdownOptionEls.length - 1);
    this.#setNameBoxDropdownActiveIndex(nextActive);
  }

  #setNameBoxDropdownActiveIndex(index: number): void {
    this.#nameBoxDropdownActiveIndex = index;
    for (let i = 0; i < this.#nameBoxDropdownOptionEls.length; i += 1) {
      const el = this.#nameBoxDropdownOptionEls[i]!;
      const active = i === index;
      el.setAttribute("aria-selected", active ? "true" : "false");
      el.classList.toggle("formula-bar-name-box-option--active", active);
      if (active) {
        this.#addressEl.setAttribute("aria-activedescendant", el.id);
        try {
          el.scrollIntoView({ block: "nearest" });
        } catch {
          // ignore (older browsers / jsdom)
        }
      }
    }

    if (index < 0) {
      this.#addressEl.removeAttribute("aria-activedescendant");
    }
  }

  #moveNameBoxDropdownSelection(delta: 1 | -1): void {
    const count = this.#nameBoxDropdownOptionEls.length;
    if (count === 0) return;
    const current = this.#nameBoxDropdownActiveIndex;
    const base = current < 0 ? (delta > 0 ? 0 : count - 1) : current;
    const next = (base + delta + count) % count;
    this.#setNameBoxDropdownActiveIndex(next);
  }

  #recordNameBoxDropdownRecent(item: NameBoxDropdownItem): void {
    const key = String(item.key ?? "").trim();
    if (!key) return;

    const deduped = [key, ...this.#nameBoxDropdownRecentKeys.filter((k) => k !== key)];
    this.#nameBoxDropdownRecentKeys = deduped.slice(0, 8);
  }

  #selectNameBoxDropdownItem(item: NameBoxDropdownItem): void {
    // Match Excel UX: selecting an item replaces the name box input text.
    this.#addressEl.value = item.label;
    this.#closeNameBoxDropdown({ restoreAddress: false, reason: "commit" });

    const ref = String(item.reference ?? "").trim();
    if (ref === "") {
      // Some workbook-defined names can refer to formulas/constants rather than a cell/range.
      // In that case, leave the text in the name box for editing instead of attempting navigation.
      this.#clearNameBoxError();
      try {
        this.#addressEl.focus({ preventScroll: true });
      } catch {
        this.#addressEl.focus();
      }
      this.#addressEl.select();
      return;
    }

    const handler = this.#callbacks.onGoTo;
    if (!handler) {
      this.#addressEl.blur();
      return;
    }

    let ok = false;
    try {
      ok = handler(ref) === true;
    } catch {
      ok = false;
    }

    if (!ok) {
      this.#setNameBoxError("Invalid reference");
      try {
        this.#addressEl.focus({ preventScroll: true });
      } catch {
        this.#addressEl.focus();
      }
      this.#addressEl.select();
      return;
    }

    this.#recordNameBoxDropdownRecent(item);
    this.#clearNameBoxError();
    // Blur after navigating so follow-up renders can update the value.
    this.#addressEl.blur();
    // `onGoTo` implementations (e.g. SpreadsheetApp) update the Name Box value via `setActiveCell`
    // synchronously while the input is still focused. Explicitly apply the updated `#nameBoxValue`
    // after blurring so the rendered text reflects the new selection immediately.
    this.#addressEl.value = this.#nameBoxValue;
  }

  #nameBoxDropdownOptionId(item: NameBoxDropdownItem): string {
    const safeKey = String(item.key ?? "")
      .trim()
      .replaceAll(/[^a-zA-Z0-9_-]/g, "_");
    return `${this.#nameBoxDropdownListEl.id}-option-${safeKey}`;
  }
}

const ESCAPE_HTML_RE = /[&<>]/g;
const ESCAPE_HTML_TEST_RE = /[&<>]/;
const ESCAPE_HTML_MAP: Record<string, string> = { "&": "&amp;", "<": "&lt;", ">": "&gt;" };
const ESCAPE_HTML_REPLACER = (ch: string): string => ESCAPE_HTML_MAP[ch] ?? ch;
const EMPTY_COLORED_REFERENCES: ReturnType<FormulaBarModel["coloredReferences"]> = [];
const EMPTY_REFERENCE_ELS: HTMLElement[] = [];

function escapeHtml(text: string): string {
  return text.replace(ESCAPE_HTML_RE, ESCAPE_HTML_REPLACER);
}

function formatPreview(value: unknown): string {
  if (value == null) return "";
  if (typeof value === "string") return ` ${value}`;
  return ` ${String(value)}`;
}

function formatArgumentPreviewValue(value: unknown): string {
  if (value == null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
  return String(value);
}

function isWhitespaceChar(ch: string): boolean {
  return ch === " " || ch === "\t" || ch === "\n" || ch === "\r";
}

/**
 * Collapse indentation/newlines in the argument expression for display in the hint panel.
 *
 * Keep string literals intact (including escaped quotes) so we don't misrepresent
 * user-entered text.
 */
function formatArgumentPreviewExpression(expr: string): string {
  const text = String(expr ?? "");
  let out = "";
  let inString = false;
  let pendingSpace = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i] ?? "";
    if (inString) {
      out += ch;
      if (ch === '"') {
        // Escaped quote inside a string literal: "" -> "
        if (text[i + 1] === '"') {
          out += '"';
          i += 1;
          continue;
        }
        inString = false;
      }
      continue;
    }

    if (ch === '"') {
      if (pendingSpace && out && !out.endsWith(" ")) out += " ";
      pendingSpace = false;
      out += ch;
      inString = true;
      continue;
    }

    if (isWhitespaceChar(ch)) {
      pendingSpace = true;
      continue;
    }

    if (pendingSpace) {
      if (out && !out.endsWith(" ")) out += " ";
      pendingSpace = false;
    }

    out += ch;
  }

  return out.trim();
}

function computeReferenceHighlights(
  text: string,
  opts: ExtractFormulaReferencesOptions | null
): FormulaReferenceHighlight[] {
  if (!isFormulaText(text)) return [];
  const { references } = extractFormulaReferences(text, undefined, undefined, opts ?? undefined);
  if (references.length === 0) return [];
  const { colored } = assignFormulaReferenceColors(references, null);
  return colored.map((ref) => ({
    range: ref.range,
    color: ref.color,
    text: ref.text,
    index: ref.index,
    active: false
  }));
}
