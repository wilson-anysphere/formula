import { searchFunctionResults } from "../command-palette/commandPaletteSearch.js";

import FUNCTION_CATALOG from "../../../../shared/functionCatalog.mjs";
import { getFunctionSignature } from "./highlight/functionSignatures.js";

type FunctionPickerItem = {
  name: string;
  signature?: string;
  summary?: string;
};

type CatalogFunction = { name?: string | null };

type FunctionSignature = NonNullable<ReturnType<typeof getFunctionSignature>>;

const ALL_FUNCTION_NAMES_SORTED: string[] = ((FUNCTION_CATALOG as { functions?: CatalogFunction[] } | null)?.functions ?? [])
  .map((fn) => String(fn?.name ?? "").trim())
  .filter((name) => name.length > 0)
  .sort((a, b) => a.localeCompare(b));

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

export function buildFunctionPickerItems(query: string, limit: number): FunctionPickerItem[] {
  const trimmed = String(query ?? "").trim();
  const cappedLimit = Math.max(0, Math.floor(limit));
  if (cappedLimit === 0) return [];

  if (!trimmed) {
    return DEFAULT_FUNCTION_NAMES.slice(0, cappedLimit).map((name) => functionPickerItemFromName(name));
  }

  return searchFunctionResults(trimmed, { limit: cappedLimit }).map((res) => ({
    name: res.name,
    signature: res.signature,
    summary: res.summary,
  }));
}

export function renderFunctionPickerList(opts: {
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

function functionPickerItemFromName(name: string): FunctionPickerItem {
  const sig = getFunctionSignature(name);
  const signature = sig ? formatSignature(sig) : undefined;
  const summary = sig?.summary?.trim?.() ? sig.summary.trim() : undefined;
  return { name, signature, summary };
}

function formatSignature(sig: FunctionSignature): string {
  const params = sig.params.map((param) => (param.optional ? `[${param.name}]` : param.name)).join(", ");
  return `${sig.name}(${params})`;
}
