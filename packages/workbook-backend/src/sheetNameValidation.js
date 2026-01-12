export const EXCEL_MAX_SHEET_NAME_LEN = 31;

export const INVALID_SHEET_NAME_CHARACTERS = [":", "\\", "/", "?", "*", "[", "]"];

const INVALID_SHEET_NAME_CHAR_SET = new Set(INVALID_SHEET_NAME_CHARACTERS);

function normalizeSheetNameForCaseInsensitiveCompare(name) {
  // Excel compares sheet names case-insensitively with Unicode NFKC normalization.
  // The backend (`formula_model::sheet_name_eq_case_insensitive`) approximates Excel via:
  // - NFKC
  // - Unicode uppercasing
  //
  // `String.prototype.normalize("NFKC")` is supported in modern JS engines / Node.
  try {
    return name.normalize("NFKC").toUpperCase();
  } catch {
    return name.toUpperCase();
  }
}

/**
 * Validate a worksheet name using Excel-compatible rules.
 *
 * This mirrors the backend validation (`formula_model::validate_sheet_name`) so the UI can
 * provide immediate feedback and avoid server-side rejection.
 *
 * @param {string} name
 * @param {{ existingNames?: Iterable<string>, ignoreExistingName?: string | null }} [options]
 * @returns {null | { kind: string, message: string, max?: number, character?: string }}
 */
export function getSheetNameValidationError(name, options = {}) {
  // Match the backend behavior: only trim to detect blank/whitespace-only names.
  if (name.trim().length === 0) {
    return { kind: "EmptyName", message: "sheet name cannot be blank" };
  }

  // Excel's 31-character limit is measured in UTF-16 code units.
  // JS `string.length` is the UTF-16 code unit count, so it matches exactly.
  if (name.length > EXCEL_MAX_SHEET_NAME_LEN) {
    return {
      kind: "TooLong",
      message: `sheet name cannot exceed ${EXCEL_MAX_SHEET_NAME_LEN} characters`,
      max: EXCEL_MAX_SHEET_NAME_LEN,
    };
  }

  for (const ch of name) {
    if (INVALID_SHEET_NAME_CHAR_SET.has(ch)) {
      return {
        kind: "InvalidCharacter",
        message: `sheet name contains invalid character \`${ch}\``,
        character: ch,
      };
    }
  }

  if (name.startsWith("'") || name.endsWith("'")) {
    return {
      kind: "LeadingOrTrailingApostrophe",
      message: "sheet name cannot begin or end with an apostrophe",
    };
  }

  const existingNames = options.existingNames;
  if (existingNames) {
    const target = normalizeSheetNameForCaseInsensitiveCompare(name);
    const ignore =
      options.ignoreExistingName == null ? null : normalizeSheetNameForCaseInsensitiveCompare(options.ignoreExistingName);

    for (const existing of existingNames) {
      // Be defensive: callers might have `null` or non-string values in a loose array.
      if (typeof existing !== "string") continue;
      if (ignore && normalizeSheetNameForCaseInsensitiveCompare(existing) === ignore) continue;
      if (normalizeSheetNameForCaseInsensitiveCompare(existing) === target) {
        return { kind: "DuplicateName", message: "sheet name already exists" };
      }
    }
  }

  return null;
}

/**
 * @param {string} name
 * @param {{ existingNames?: Iterable<string>, ignoreExistingName?: string | null }} [options]
 */
export function getSheetNameValidationErrorMessage(name, options = {}) {
  return getSheetNameValidationError(name, options)?.message ?? null;
}
