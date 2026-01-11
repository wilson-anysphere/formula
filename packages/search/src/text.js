function isTypedValue(value) {
  return value != null && typeof value === "object" && typeof value.t === "string";
}

function typedValueToString(value) {
  // The formula engine uses a stable JSON encoding (see tools/excel-oracle/value-encoding.md).
  switch (value.t) {
    case "blank":
      return "";
    case "n":
      return value.v == null ? "" : String(value.v);
    case "s":
      return value.v == null ? "" : String(value.v);
    case "b":
      // Excel displays booleans as TRUE/FALSE.
      return value.v ? "TRUE" : "FALSE";
    case "e":
      return value.v == null ? "" : String(value.v);
    case "arr":
      // Dynamic arrays/spills are represented as a matrix. There isn't a single
      // Excel "cell string" to search, but we still want stable behavior.
      // Keep this cheap (avoid deep formatting) while ensuring we don't return
      // "[object Object]" which breaks search UX.
      try {
        return JSON.stringify(value);
      } catch {
        return String(value);
      }
    default:
      return String(value);
  }
}

export function formatCellValue(value) {
  if (value == null) return "";
  if (isTypedValue(value)) return typedValueToString(value);
  return String(value);
}

export function getValueText(cell, valueMode = "display") {
  if (!cell) return "";
  if (valueMode === "raw") return formatCellValue(cell.value);
  if (cell.display != null) return String(cell.display);
  return formatCellValue(cell.value);
}

export function getCellText(cell, { lookIn = "values", valueMode = "display" } = {}) {
  if (!cell) return "";

  if (lookIn === "formulas") {
    if (cell.formula != null && cell.formula !== "") return String(cell.formula);
    // Constants in "Look in: formulas" are treated as raw user input.
    return formatCellValue(cell.value);
  }

  // values
  return getValueText(cell, valueMode);
}

