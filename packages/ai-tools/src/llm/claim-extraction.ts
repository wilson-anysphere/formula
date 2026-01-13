export type SpreadsheetClaimMeasure =
  | "mean"
  | "sum"
  | "count"
  | "median"
  | "mode"
  | "stdev"
  | "variance"
  | "min"
  | "max"
  | "correlation";

export type ExtractedSpreadsheetClaim =
  | {
      kind: "range_stat";
      measure: SpreadsheetClaimMeasure;
      /**
       * A1 reference (cell or range). May omit sheet prefix; callers can rely on
       * the tool executor's default sheet when needed.
       */
      reference?: string;
      expected: number;
      /**
       * Raw matched text snippet from the assistant message.
       */
      source: string;
    }
  | {
      kind: "cell_value";
      reference: string;
      expected: number;
      source: string;
    };

export interface ExtractVerifiableClaimsParams {
  assistantText: string;
  userText?: string;
  attachments?: unknown[] | null;
  toolCalls?: Array<{ name: string; parameters?: unknown; arguments?: unknown }>;
}

const SHEET_NAME_PATTERN = "(?:'(?:[^']|'')+'|[A-Za-z0-9_.-]+)";
const CELL_PATTERN = "\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*";
const A1_REFERENCE_PATTERN = `(?:(?:${SHEET_NAME_PATTERN})!\\s*)?${CELL_PATTERN}(?:\\s*:\\s*${CELL_PATTERN})?`;
const A1_CELL_REFERENCE_PATTERN = `(?:(?:${SHEET_NAME_PATTERN})!\\s*)?${CELL_PATTERN}`;

// Supports:
// - comma-separated thousands ("1,234")
// - decimals ("12.34")
// - leading-decimal floats (".5")
// - scientific notation ("1e-3")
// - optional percent suffix ("10%")
// - optional parentheses for negative formatting ("(1,234)")
const NUMBER_CORE_PATTERN = "[-+]?(?:[€£$])?(?:\\d+(?:,\\d{3})*(?:\\.\\d+)?|\\.\\d+)(?:e[-+]?\\d+)?%?";
const NUMBER_PATTERN = `(?:${NUMBER_CORE_PATTERN}|\\(${NUMBER_CORE_PATTERN}\\))`;

const KEYWORD_TO_MEASURE: Record<string, SpreadsheetClaimMeasure> = {
  average: "mean",
  avg: "mean",
  mean: "mean",
  median: "median",
  mode: "mode",
  sum: "sum",
  total: "sum",
  stdev: "stdev",
  stddev: "stdev",
  "std dev": "stdev",
  "std. dev": "stdev",
  "standard deviation": "stdev",
  var: "variance",
  variance: "variance",
  count: "count",
  min: "min",
  minimum: "min",
  max: "max",
  maximum: "max",
  correlation: "correlation",
  corr: "correlation"
};

function escapeRegexLiteral(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

const KEYWORD_PATTERN = Object.keys(KEYWORD_TO_MEASURE)
  .slice()
  .sort((a, b) => b.length - a.length)
  // Treat keyword spaces as flexible whitespace and escape regex metacharacters so
  // we can safely support shorthands like "std. dev".
  .map((keyword) => escapeRegexLiteral(keyword).replace(/\s+/g, "\\s+"))
  .join("|");

const RANGE_STAT_WITH_REF_RE_1 = new RegExp(
  `\\b(?<keyword>${KEYWORD_PATTERN})\\b(?:\\s+(?:of|for|in|within))?\\s+(?:(?:the\\s+)?range\\s+)?\\(?\\s*(?<ref>${A1_REFERENCE_PATTERN})\\s*\\)?\\s*(?:is|was|=|:|equals)?\\s*(?<num>${NUMBER_PATTERN})`,
  "gi"
);

// Support formula/function-style phrasings like "MEDIAN(A1:A10) = 5".
const RANGE_STAT_WITH_REF_FUNCTION_CALL_RE = new RegExp(
  `\\b(?<keyword>${KEYWORD_PATTERN})\\b\\s*\\(\\s*(?<ref>${A1_REFERENCE_PATTERN})\\s*\\)\\s*(?:is|was|=|:|equals)?\\s*(?<num>${NUMBER_PATTERN})`,
  "gi"
);

const RANGE_STAT_WITH_REF_RE_2 = new RegExp(
  `(?<ref>${A1_REFERENCE_PATTERN})\\s*[,:-]?\\s*(?<keyword>${KEYWORD_PATTERN})\\b\\s*(?:is|was|=|:|equals)?\\s*(?<num>${NUMBER_PATTERN})`,
  "gi"
);

// Support implicit count claims like "There are 99 values in A1:A10" or "A1:A10 has 99 values".
const COUNT_WITH_REF_RE_1 = new RegExp(
  `\\bthere\\s+(?:are|were)\\s+(?<num>${NUMBER_PATTERN})\\s+(?:values|numbers|observations|data\\s+points|datapoints)\\b\\s+(?:in|within|for)\\s+(?:the\\s+)?(?:range\\s+)?\\(?\\s*(?<ref>${A1_REFERENCE_PATTERN})\\s*\\)?`,
  "gi"
);

const COUNT_WITH_REF_RE_2 = new RegExp(
  `\\(?\\s*(?<ref>${A1_REFERENCE_PATTERN})\\s*\\)?\\s+(?:has|contains)\\s+(?<num>${NUMBER_PATTERN})\\s+(?:values|numbers|observations|data\\s+points|datapoints)\\b`,
  "gi"
);

const COUNT_WITH_REF_RE_3 = new RegExp(
  `\\b(?:the\\s+)?(?:number|no\\.)\\s+of\\s+(?:values|numbers|observations|data\\s+points|datapoints)\\b\\s+(?:in|within|for)\\s+(?:the\\s+)?(?:range\\s+)?\\(?\\s*(?<ref>${A1_REFERENCE_PATTERN})\\s*\\)?\\s*(?:is|was|=|:|equals)?\\s*(?<num>${NUMBER_PATTERN})`,
  "gi"
);

const RANGE_STAT_NO_REF_RE = new RegExp(
  `\\b(?<keyword>${KEYWORD_PATTERN})\\b\\s*(?:is|was|=|:|equals)\\s*(?<num>${NUMBER_PATTERN})`,
  "gi"
);

const CELL_VALUE_RE = new RegExp(
  `(?<ref>${A1_CELL_REFERENCE_PATTERN})\\s*(?:is|was|=|:|equals)\\s*(?<num>${NUMBER_PATTERN})`,
  "gi"
);

const A1_REFERENCE_FINDER_RE = new RegExp(A1_REFERENCE_PATTERN, "gi");

interface MatchSpan<T> {
  start: number;
  end: number;
  value: T;
}

/**
 * Deterministically extract spreadsheet-related numeric claims from an assistant
 * response (with lightweight context from user text / attachments / tool calls).
 *
 * This is intentionally heuristic (regex-based). False positives are acceptable
 * as long as the output is deterministic and testable.
 */
export function extractVerifiableClaims(params: ExtractVerifiableClaimsParams): ExtractedSpreadsheetClaim[] {
  const assistantText = String(params.assistantText ?? "");
  const spans: Array<MatchSpan<ExtractedSpreadsheetClaim>> = [];

  spans.push(...collectRangeStatWithRef(RANGE_STAT_WITH_REF_RE_1, assistantText));
  spans.push(...collectRangeStatWithRef(RANGE_STAT_WITH_REF_FUNCTION_CALL_RE, assistantText));
  spans.push(...collectRangeStatWithRef(RANGE_STAT_WITH_REF_RE_2, assistantText));
  spans.push(...collectCountWithRef(COUNT_WITH_REF_RE_1, assistantText));
  spans.push(...collectCountWithRef(COUNT_WITH_REF_RE_2, assistantText));
  spans.push(...collectCountWithRef(COUNT_WITH_REF_RE_3, assistantText));
  spans.push(...collectCellValueClaims(assistantText));

  const occupied = spans.map(({ start, end }) => ({ start, end }));

  const noRef = collectRangeStatNoRef(assistantText).filter((match) => !overlapsAny(match, occupied));
  spans.push(...noRef);

  spans.sort((a, b) => a.start - b.start || a.end - b.end);

  const rawClaims = spans.map((span) => span.value);

  const resolved = resolveMissingReferences(rawClaims, {
    userText: params.userText,
    attachments: params.attachments,
    toolCalls: params.toolCalls
  });

  return dedupeClaims(resolved);
}

function collectCountWithRef(regex: RegExp, text: string): Array<MatchSpan<ExtractedSpreadsheetClaim>> {
  const out: Array<MatchSpan<ExtractedSpreadsheetClaim>> = [];
  for (const match of text.matchAll(regex)) {
    const groups = match.groups ?? {};
    const refRaw = String(groups.ref ?? "").trim();
    const ref = refRaw ? normalizeA1Reference(refRaw) : undefined;
    if (!ref) continue;

    const expected = parseNumberToken(String(groups.num ?? ""));
    if (expected == null) continue;

    out.push({
      start: match.index ?? 0,
      end: (match.index ?? 0) + match[0]!.length,
      value: {
        kind: "range_stat",
        measure: "count",
        reference: ref,
        expected,
        source: match[0]!.trim()
      }
    });
  }
  return out;
}

function collectRangeStatWithRef(
  regex: RegExp,
  text: string
): Array<MatchSpan<ExtractedSpreadsheetClaim>> {
  const out: Array<MatchSpan<ExtractedSpreadsheetClaim>> = [];
  for (const match of text.matchAll(regex)) {
    const groups = match.groups ?? {};
    const keywordRaw = String(groups.keyword ?? "").trim().toLowerCase().replace(/\s+/g, " ");
    const measure = KEYWORD_TO_MEASURE[keywordRaw];
    if (!measure) continue;

    const refRaw = String(groups.ref ?? "").trim();
    const ref = refRaw ? normalizeA1Reference(refRaw) : undefined;

    const expected = parseNumberToken(String(groups.num ?? ""));
    if (expected == null) continue;

    out.push({
      start: match.index ?? 0,
      end: (match.index ?? 0) + match[0]!.length,
      value: {
        kind: "range_stat",
        measure,
        reference: ref,
        expected,
        source: match[0]!.trim()
      }
    });
  }
  return out;
}

function collectRangeStatNoRef(text: string): Array<MatchSpan<ExtractedSpreadsheetClaim>> {
  const out: Array<MatchSpan<ExtractedSpreadsheetClaim>> = [];
  for (const match of text.matchAll(RANGE_STAT_NO_REF_RE)) {
    const groups = match.groups ?? {};
    const keywordRaw = String(groups.keyword ?? "").trim().toLowerCase().replace(/\s+/g, " ");
    const measure = KEYWORD_TO_MEASURE[keywordRaw];
    if (!measure) continue;

    const expected = parseNumberToken(String(groups.num ?? ""));
    if (expected == null) continue;

    out.push({
      start: match.index ?? 0,
      end: (match.index ?? 0) + match[0]!.length,
      value: {
        kind: "range_stat",
        measure,
        reference: undefined,
        expected,
        source: match[0]!.trim()
      }
    });
  }
  return out;
}

function collectCellValueClaims(text: string): Array<MatchSpan<ExtractedSpreadsheetClaim>> {
  const out: Array<MatchSpan<ExtractedSpreadsheetClaim>> = [];
  for (const match of text.matchAll(CELL_VALUE_RE)) {
    const start = match.index ?? 0;
    // Avoid matching the trailing cell in a range reference like "A1:A3 is 2".
    // In that string, "A3 is 2" would otherwise look like a cell-value claim.
    const charBefore = start > 0 ? text[start - 1] : "";
    if (charBefore === ":") continue;

    const groups = match.groups ?? {};
    const refRaw = String(groups.ref ?? "").trim();
    const ref = refRaw ? normalizeA1Reference(refRaw) : "";
    if (!ref) continue;

    const expected = parseNumberToken(String(groups.num ?? ""));
    if (expected == null) continue;

    out.push({
      start,
      end: start + match[0]!.length,
      value: {
        kind: "cell_value",
        reference: ref,
        expected,
        source: match[0]!.trim()
      }
    });
  }
  return out;
}

function overlapsAny(span: { start: number; end: number }, occupied: Array<{ start: number; end: number }>): boolean {
  return occupied.some((other) => overlaps(span, other));
}

function overlaps(a: { start: number; end: number }, b: { start: number; end: number }): boolean {
  return a.start < b.end && b.start < a.end;
}

function resolveMissingReferences(
  claims: ExtractedSpreadsheetClaim[],
  context: { userText?: string; attachments?: unknown[] | null; toolCalls?: Array<{ name: string; parameters?: unknown; arguments?: unknown }> }
): ExtractedSpreadsheetClaim[] {
  const primaryRef = resolvePrimaryReference(context);
  if (!primaryRef) return claims;

  return claims.map((claim) => {
    if (claim.kind !== "range_stat") return claim;
    if (claim.reference) return claim;
    return { ...claim, reference: primaryRef };
  });
}

function resolvePrimaryReference(context: {
  userText?: string;
  attachments?: unknown[] | null;
  toolCalls?: Array<{ name: string; parameters?: unknown; arguments?: unknown }>;
}): string | null {
  const userRefs = uniqueInOrder([
    ...extractA1ReferencesFromText(String(context.userText ?? "")),
    ...extractA1ReferencesFromAttachments(context.attachments)
  ]);
  if (userRefs.length === 1) return userRefs[0]!;

  const toolRefs = uniqueInOrder(extractA1ReferencesFromToolCalls(context.toolCalls));
  if (toolRefs.length === 1) return toolRefs[0]!;

  return null;
}

function extractA1ReferencesFromAttachments(attachments: unknown[] | null | undefined): string[] {
  if (!Array.isArray(attachments)) return [];
  const out: string[] = [];
  for (const item of attachments) {
    if (!item || typeof item !== "object") continue;
    const ref = (item as any).reference;
    if (typeof ref === "string" && ref.trim()) {
      // Attachment references are typically already A1 strings, but run them
      // through the same extractor for safety (e.g. "Sheet1!A1:D10").
      out.push(...extractA1ReferencesFromText(ref));
    }
  }
  return out;
}

function extractA1ReferencesFromToolCalls(
  toolCalls: Array<{ name: string; parameters?: unknown; arguments?: unknown }> | null | undefined
): string[] {
  if (!Array.isArray(toolCalls)) return [];
  const out: string[] = [];
  for (const call of toolCalls) {
    const params = call?.parameters ?? call?.arguments;
    if (!params || typeof params !== "object" || Array.isArray(params)) continue;
    const obj = params as Record<string, unknown>;

    if (typeof obj.range === "string") {
      // read_range / compute_statistics / filter_range / detect_anomalies, etc.
      const refs = extractA1ReferencesFromText(obj.range);
      if (refs.length) {
        out.push(...refs);
        continue;
      }
    }

    if (typeof obj.source_range === "string") {
      // create_pivot_table
      const refs = extractA1ReferencesFromText(obj.source_range);
      if (refs.length) {
        out.push(...refs);
        continue;
      }
    }
    if (typeof obj.sourceRange === "string") {
      // create_pivot_table (camelCase alias; common in some tool-call payloads)
      const refs = extractA1ReferencesFromText(obj.sourceRange);
      if (refs.length) {
        out.push(...refs);
        continue;
      }
    }

    if (typeof obj.data_range === "string") {
      // create_chart
      const refs = extractA1ReferencesFromText(obj.data_range);
      if (refs.length) {
        out.push(...refs);
        continue;
      }
    }
    if (typeof obj.dataRange === "string") {
      // create_chart (camelCase alias; common in some tool-call payloads)
      const refs = extractA1ReferencesFromText(obj.dataRange);
      if (refs.length) {
        out.push(...refs);
        continue;
      }
    }

    if (typeof obj.cell === "string") {
      // write_cell, etc.
      const refs = extractA1ReferencesFromText(obj.cell);
      if (refs.length) {
        out.push(...refs);
        continue;
      }
    }

    if (typeof obj.destination === "string") {
      const refs = extractA1ReferencesFromText(obj.destination);
      if (refs.length) {
        out.push(...refs);
        continue;
      }
    }

    if (typeof obj.position === "string") {
      // create_chart placement anchor (optional)
      out.push(...extractA1ReferencesFromText(obj.position));
    }
  }
  return out;
}

function extractA1ReferencesFromText(text: string): string[] {
  const out: string[] = [];
  if (!text) return out;
  for (const match of text.matchAll(A1_REFERENCE_FINDER_RE)) {
    const raw = match[0];
    if (!raw) continue;
    const ref = normalizeA1Reference(raw);
    if (ref) out.push(ref);
  }
  return uniqueInOrder(out);
}

function normalizeA1Reference(raw: string): string {
  const input = String(raw ?? "").trim();
  if (!input) return "";

  const bangIndex = input.lastIndexOf("!");
  const sheetPrefix = bangIndex === -1 ? "" : input.slice(0, bangIndex + 1).trim();
  let rest = bangIndex === -1 ? input : input.slice(bangIndex + 1);

  rest = rest.replace(/\s+/g, "");
  rest = rest.replace(/[a-z]/g, (char) => char.toUpperCase());

  return sheetPrefix ? `${sheetPrefix}${rest}` : rest;
}

function parseNumberToken(token: string): number | null {
  let raw = String(token ?? "").trim();
  if (!raw) return null;

  const hasParens = raw.startsWith("(") && raw.endsWith(")");
  if (hasParens) {
    raw = raw.slice(1, -1).trim();
  }

  // Common spreadsheet formatting: currency symbols. We intentionally keep this
  // conservative (single leading symbol) to avoid over-matching.
  raw = raw.replace(/^([-+])?[€£$]/, (_match, sign: string | undefined) => sign ?? "");

  const isPercent = raw.endsWith("%");
  const normalized = raw.replace(/,/g, "").replace(/%$/, "");
  const value = Number(normalized);
  if (!Number.isFinite(value)) return null;
  let out = isPercent ? value / 100 : value;
  // Parentheses typically indicate a negative number in spreadsheets.
  // Only apply the negation when the parsed value is positive to avoid
  // double-negating tokens like "(-5)".
  if (hasParens && out > 0) out = -out;
  return out;
}

function dedupeClaims(claims: ExtractedSpreadsheetClaim[]): ExtractedSpreadsheetClaim[] {
  const seen = new Set<string>();
  const out: ExtractedSpreadsheetClaim[] = [];
  for (const claim of claims) {
    const key = stableClaimKey(claim);
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(claim);
  }
  return out;
}

function stableClaimKey(claim: ExtractedSpreadsheetClaim): string {
  if (claim.kind === "cell_value") {
    return `cell_value|${claim.reference}|${claim.expected}`;
  }
  return `range_stat|${claim.measure}|${claim.reference ?? ""}|${claim.expected}`;
}

function uniqueInOrder(values: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const value of values) {
    const normalized = String(value ?? "").trim();
    if (!normalized) continue;
    if (seen.has(normalized)) continue;
    seen.add(normalized);
    out.push(normalized);
  }
  return out;
}
