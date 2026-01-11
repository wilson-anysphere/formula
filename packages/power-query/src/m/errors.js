/**
 * Power Query M language (subset) diagnostics.
 *
 * The goal of these errors is to be actionable in a UI: they include a precise
 * location (line/column), a short list of expected tokens, and a best-effort
 * suggestion.
 */

/**
 * @typedef {{ offset: number; line: number; column: number }} MLocation
 */

/**
 * @typedef {{ start: MLocation; end: MLocation }} MSpan
 */

/**
 * @param {string} source
 * @param {MLocation} loc
 * @returns {{ lineText: string; caretLine: string }}
 */
export function formatSourceLocation(source, loc) {
  const lines = source.split(/\r?\n/);
  const lineIndex = Math.max(0, Math.min(lines.length - 1, loc.line - 1));
  const lineText = lines[lineIndex] ?? "";
  const caretPos = Math.max(0, Math.min(lineText.length, loc.column - 1));
  const caretLine = `${" ".repeat(caretPos)}^`;
  return { lineText, caretLine };
}

/**
 * @param {string[]} expected
 * @param {{ type: string; value?: string } | null} found
 * @returns {string[]}
 */
function suggestionForParse(expected, found) {
  if (!found) return [];
  if (found.type === "eof") return ["The script ended unexpectedly. Did you forget to close a bracket or parenthesis?"];
  if (found.value === ")" && expected.some((e) => e.includes("expression"))) {
    return ["Did you forget an argument before ')' ?"];
  }
  if (found.value === "]" && expected.some((e) => e.includes("field"))) {
    return ["Did you forget a field name inside '[]'?"];
  }
  if (found.value === "in" && expected.includes(",")) {
    return ["Did you forget a comma between let-bindings?"];
  }
  return [];
}

export class MLanguageError extends Error {
  /**
   * @param {string} message
   * @param {{
   *   kind: "parse" | "compile";
   *   location: MLocation;
   *   expected?: string[];
   *   found?: { type: string; value?: string } | null;
   *   source?: string;
   * }} options
   */
  constructor(message, options) {
    const expected = options.expected ?? [];
    const found = options.found ?? null;
    const suggestion = options.kind === "parse" ? suggestionForParse(expected, found) : [];
    const locationText = `line ${options.location.line}, column ${options.location.column}`;
    const expectedText = expected.length ? ` Expected: ${expected.join(", ")}.` : "";
    const foundText = found ? ` Found: ${found.value ?? found.type}.` : "";
    const suggestionText = suggestion.length ? `\nSuggestion: ${suggestion.join(" ")}` : "";

    let contextText = "";
    if (options.source) {
      const { lineText, caretLine } = formatSourceLocation(options.source, options.location);
      contextText = `\n${lineText}\n${caretLine}`;
    }

    super(`${message} (${locationText}).${expectedText}${foundText}${suggestionText}${contextText}`);
    this.name = this.constructor.name;
    this.kind = options.kind;
    this.location = options.location;
    this.expected = expected;
    this.found = found;
  }
}

export class MLanguageSyntaxError extends MLanguageError {
  /**
   * @param {string} message
   * @param {ConstructorParameters<typeof MLanguageError>[1]} options
   */
  constructor(message, options) {
    super(message, { ...options, kind: "parse" });
  }
}

export class MLanguageCompileError extends MLanguageError {
  /**
   * @param {string} message
   * @param {ConstructorParameters<typeof MLanguageError>[1]} options
   */
  constructor(message, options) {
    super(message, { ...options, kind: "compile" });
  }
}

