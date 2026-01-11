import { excelWildcardToRegExp } from "./wildcards.js";
import { formatCellValue, getValueText } from "./text.js";

function looksLikeNumber(text) {
  const s = String(text).trim();
  if (s === "") return false;
  return /^-?(?:\d+\.?\d*|\d*\.?\d+)(?:[eE]-?\d+)?$/.test(s);
}

function coerceReplacementValue(text) {
  if (looksLikeNumber(text)) return Number(text);
  return String(text);
}

function replaceInString(text, re, replacement) {
  let count = 0;
  const out = String(text).replace(re, () => {
    count++;
    return replacement;
  });
  return { out, count };
}

export function applyReplaceToCell(cell, query, replacement, options = {}, { replaceAll = false } = {}) {
  if (!cell) return { cell: null, replaced: false, replacements: 0 };
  const {
    lookIn = "values",
    valueMode = "display",
    matchCase = false,
    matchEntireCell = false,
    useWildcards = true,
  } = options;

  const re = excelWildcardToRegExp(query, {
    matchCase,
    matchEntireCell,
    useWildcards,
    global: replaceAll,
  });

  // Excel semantics:
  // - Look in: Formulas => replace in the cell "input" (formula string or raw constant)
  // - Look in: Values   => replace in displayed/evaluated value, overwriting formulas with constants
  if (lookIn === "formulas") {
    if (cell.formula != null && cell.formula !== "") {
      const { out, count } = replaceInString(cell.formula, re, replacement);
      if (count === 0) return { cell, replaced: false, replacements: 0 };
      return { cell: { ...cell, formula: out }, replaced: true, replacements: count };
    }

    const original = formatCellValue(cell.value);
    const { out, count } = replaceInString(original, re, replacement);
    if (count === 0) return { cell, replaced: false, replacements: 0 };
    return { cell: { ...cell, value: coerceReplacementValue(out) }, replaced: true, replacements: count };
  }

  // values
  const original = getValueText(cell, valueMode);
  const { out, count } = replaceInString(original, re, replacement);
  if (count === 0) return { cell, replaced: false, replacements: 0 };

  return {
    cell: { ...cell, formula: undefined, value: coerceReplacementValue(out), display: undefined },
    replaced: true,
    replacements: count,
  };
}
