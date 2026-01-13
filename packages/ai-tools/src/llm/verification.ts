import { extractVerifiableClaims } from "./claim-extraction.ts";
import { parseSpreadsheetNumber } from "../executor/number-parsing.ts";

export interface VerificationResult {
  /**
   * Heuristic signal that the user's query likely needs spreadsheet tools to
   * avoid guessing (e.g. read/compute for questions, or write/apply for actions).
   */
  needs_tools: boolean;
  /**
   * Whether any tool calls were executed at all during the run.
   */
  used_tools: boolean;
  /**
   * Whether the run is considered verified.
   *
   * For tool-usage-only verification (see `verifyToolUsage`), a run is verified
   * when tools were not needed or the required kind of tool call succeeded
   * (by default, any successful tool; for analysis questions this may require a
   * successful read/compute/detect/filter tool).
   *
   * When post-response claim verification runs (see `verifyAssistantClaims`),
   * this reflects whether all extracted numeric claims were validated against
   * spreadsheet computations.
   */
  verified: boolean;
  /**
   * Deterministic confidence score in [0, 1].
   *
   * This is intentionally heuristic and conservative; it exists to surface
   * "compute over hallucination" failures (e.g. answering about data without
   * using tools).
   */
  confidence: number;
  warnings: string[];
  /**
   * Optional claim-level verification details.
   *
   * When present, each entry represents a deterministic check comparing a
   * numeric assistant claim (`expected`) against a spreadsheet computation
   * (`actual`) plus tool evidence.
   */
  claims?: Array<{
    claim: string;
    verified: boolean;
    expected?: number | string | null;
    actual?: number | string | null;
    toolEvidence?: unknown;
  }>;
}

export interface VerifyToolUsageParams {
  needsTools: boolean;
  /**
   * What kind of tool evidence is required to consider the run "verified" when
   * `needsTools` is true.
   *
   * - "any": (default) any successful tool call counts as verification. This is
   *   appropriate for mutation/action requests like "Set A1 to 1".
   * - "verified": require at least one successful read/compute/detect/filter tool
   *   call. This is appropriate for analysis/data questions like "What is the
   *   average of A1:A10?" where a mutation tool (e.g. write_cell) is not evidence
   *   that the answer was computed.
   */
  requiredToolKind?: "any" | "verified";
  /**
   * Tool calls emitted during the run. The verifier only depends on a subset of
   * fields so it can work with both audit logs and UI-level tool call tracking.
   */
  toolCalls: Array<{ name: string; ok?: boolean }>;
}

export interface ClassifyQueryNeedsToolsParams {
  userText: string;
  attachments?: unknown[] | null;
}

export interface ClassifyRequiredToolKindParams {
  userText: string;
  attachments?: unknown[] | null;
}

export interface VerifyAssistantClaimsParams {
  assistantText: string;
  userText?: string;
  attachments?: unknown[] | null;
  toolCalls?: Array<{ name: string; parameters?: unknown; arguments?: unknown }>;
  toolExecutor: { tools?: Array<{ name: string }>; execute: (call: any) => Promise<any> };
  maxClaims?: number;
}

export interface ClaimVerificationSummary {
  claims: NonNullable<VerificationResult["claims"]>;
  verified: boolean;
  confidence: number;
  warnings: string[];
}

export { extractVerifiableClaims, type ExtractedSpreadsheetClaim, type SpreadsheetClaimMeasure } from "./claim-extraction.ts";

// Match simple A1 references like `A1` plus mixed/absolute forms like `A$1`.
// Note: `$A$1` will also be detected via the substring `A$1`.
const A1_REFERENCE_RE = /\b[A-Z]{1,3}\$?\d+\b/i;

// These keywords intentionally skew broad; false positives are acceptable
// because the only consequence is "use tools or mark unverified".
const SPREADSHEET_KEYWORDS = [
  "sheet",
  "range",
  "cell",
  "column",
  "row",
  "sum",
  "average",
  "pivot",
  "filter",
  "sort",
  "table",
  "chart",
  "formula"
];

const SPREADSHEET_KEYWORD_REGEXES = SPREADSHEET_KEYWORDS.map((keyword) => new RegExp(`\\b${keyword}\\b`, "i"));

const VERIFIED_TOOL_PREFIXES = ["read_", "compute_", "detect_", "filter_"];

const VERIFIED_TOOL_NAMES = new Set(["read_range", "compute_statistics", "detect_anomalies", "filter_range"]);

// Analysis-ish keywords used to decide whether a prompt is a "data question" that
// should require read/compute tooling (as opposed to a pure mutation/action).
//
// This list intentionally skews broad. False positives are acceptable because the
// only consequence is marking the run "unverified" unless a read/compute tool is
// used.
const ANALYSIS_KEYWORDS = [
  "average",
  "avg",
  "mean",
  "median",
  "mode",
  "sum",
  "total",
  "count",
  "min",
  "minimum",
  "max",
  "maximum",
  "stdev",
  "stddev",
  "standard deviation",
  "variance",
  "quartile",
  "percentile",
  "percentage",
  "correlation",
  "corr"
];

const ANALYSIS_KEYWORD_REGEXES = ANALYSIS_KEYWORDS.map((keyword) => new RegExp(`\\b${escapeRegex(keyword)}\\b`, "i"));

// Mutation-ish verbs used to identify "action requests". Today this is only used
// as a secondary signal; analysis intent always wins.
const MUTATION_KEYWORDS = ["set", "write", "replace", "update", "fill", "clear", "delete", "insert", "append", "remove"];

const MUTATION_KEYWORD_REGEXES = MUTATION_KEYWORDS.map((keyword) => new RegExp(`\\b${escapeRegex(keyword)}\\b`, "i"));

export function classifyQueryNeedsTools(params: ClassifyQueryNeedsToolsParams): boolean {
  const attachments = params.attachments ?? [];
  if (Array.isArray(attachments) && attachments.length > 0) return true;

  const text = params.userText ?? "";
  if (!text) return false;

  if (A1_REFERENCE_RE.test(text)) return true;

  for (const regex of SPREADSHEET_KEYWORD_REGEXES) {
    if (regex.test(text)) return true;
  }

  return false;
}

/**
 * Infer what kind of tool evidence should be required to mark a run verified.
 *
 * This intentionally distinguishes:
 * - analysis/data questions ("verified"): require at least one successful
 *   read/compute/detect/filter tool call.
 * - mutation/action requests ("any"): any successful tool call counts.
 *
 * If both intents are present, this prefers "verified" (conservative).
 */
export function classifyRequiredToolKind(params: ClassifyRequiredToolKindParams): "any" | "verified" {
  const attachments = params.attachments ?? [];
  if (Array.isArray(attachments) && attachments.length > 0) return "verified";

  const text = params.userText ?? "";
  if (!text) return "any";

  // Strong signal for a "question" intent.
  const looksLikeQuestion =
    text.includes("?") || /^\s*(?:what|which|who|whom|whose|why|how|when|where)\b/i.test(text.trim());

  const hasAnalysisKeyword = ANALYSIS_KEYWORD_REGEXES.some((re) => re.test(text));
  const hasMutationKeyword = MUTATION_KEYWORD_REGEXES.some((re) => re.test(text));

  if (looksLikeQuestion || hasAnalysisKeyword) return "verified";
  if (hasMutationKeyword) return "any";
  return "any";
}

export function verifyToolUsage(params: VerifyToolUsageParams): VerificationResult {
  const needs_tools = params.needsTools;
  const toolCalls = Array.isArray(params.toolCalls) ? params.toolCalls : [];
  const used_tools = toolCalls.length > 0;
  const requiredToolKind = params.requiredToolKind ?? "any";

  const successfulTool = toolCalls.some((call) => call.ok === true);
  const successfulVerifiedTool = toolCalls.some((call) => isVerifiedToolName(call.name) && call.ok === true);

  // If tools aren't required, we treat the response as verified.
  if (!needs_tools) {
    return {
      needs_tools,
      used_tools,
      verified: true,
      confidence: 0.8,
      warnings: []
    };
  }

  const meetsRequirement = requiredToolKind === "verified" ? successfulVerifiedTool : successfulTool;

  if (meetsRequirement) {
    return {
      needs_tools,
      used_tools,
      verified: true,
      confidence: successfulVerifiedTool ? 0.9 : 0.85,
      warnings: []
    };
  }

  const warnings: string[] = [];
  let confidence = 0.4;

  if (!used_tools) {
    warnings.push("No tools were used; answer may be a guess.");
    confidence = 0.2;
  } else if (successfulTool && requiredToolKind === "verified" && !successfulVerifiedTool) {
    warnings.push("No verified data tools succeeded; answer may be unverified.");
    confidence = 0.35;
  } else {
    warnings.push("No tools succeeded; answer may be unverified.");
    confidence = 0.35;
  }

  return {
    needs_tools,
    used_tools,
    verified: false,
    confidence,
    warnings
  };
}

/**
 * Post-response claim verification: extract numeric spreadsheet claims from the
 * assistant response and deterministically verify them via spreadsheet tools.
 *
 * Returns `null` when no verifiable claims are found.
 */
export async function verifyAssistantClaims(params: VerifyAssistantClaimsParams): Promise<ClaimVerificationSummary | null> {
  const extracted = extractVerifiableClaims({
    assistantText: params.assistantText,
    userText: params.userText,
    attachments: params.attachments,
    toolCalls: params.toolCalls
  });

  if (extracted.length === 0) return null;

  const maxClaims = params.maxClaims ?? 8;
  const claimsToVerify = extracted.slice(0, Math.max(0, maxClaims));

  const hasTool = (name: string) => Array.isArray(params.toolExecutor.tools) && params.toolExecutor.tools.some((t) => t?.name === name);

  // Precompute compute_statistics results per range to avoid repeated tool calls
  // when multiple claims reference the same range.
  const computeStatisticsByRange = new Map<
    string,
    { call: any; sanitizedResult?: unknown; rawResult?: any; error?: string }
  >();

  if (hasTool("compute_statistics")) {
    const measuresByRange = new Map<string, Array<import("./claim-extraction.ts").SpreadsheetClaimMeasure>>();
    const rangesInOrder: string[] = [];

    for (const claim of claimsToVerify) {
      if (claim.kind !== "range_stat") continue;
      const reference = claim.reference;
      if (!reference) continue;

      let measures = measuresByRange.get(reference);
      if (!measures) {
        measures = [];
        measuresByRange.set(reference, measures);
        rangesInOrder.push(reference);
      }
      if (!measures.includes(claim.measure)) {
        measures.push(claim.measure);
      }
    }

    for (const range of rangesInOrder) {
      const measures = measuresByRange.get(range) ?? [];
      const toolCall = { name: "compute_statistics", arguments: { range, measures } };
      try {
        const rawResult = await params.toolExecutor.execute(toolCall);
        computeStatisticsByRange.set(range, {
          call: toolCall,
          rawResult,
          sanitizedResult: sanitizeVerificationToolResult(rawResult)
        });
      } catch (error) {
        computeStatisticsByRange.set(range, {
          call: toolCall,
          error: error instanceof Error ? error.message : String(error)
        });
      }
    }
  }

  const verifiedClaims: NonNullable<VerificationResult["claims"]> = [];

  for (const claim of claimsToVerify) {
    if (claim.kind === "range_stat") {
      verifiedClaims.push(
        await verifyRangeStatClaim(claim, { toolExecutor: params.toolExecutor, hasTool, computeStatisticsByRange })
      );
      continue;
    }
    verifiedClaims.push(await verifyCellValueClaim(claim, { toolExecutor: params.toolExecutor, hasTool }));
  }

  const verifiedCount = verifiedClaims.filter((c) => c.verified).length;
  const confidence = verifiedClaims.length ? verifiedCount / verifiedClaims.length : 0;
  const verified = verifiedClaims.length > 0 && verifiedCount === verifiedClaims.length;

  const warnings: string[] = [];
  if (!verified) warnings.push("One or more numeric claims did not match spreadsheet results.");

  return {
    claims: verifiedClaims,
    verified,
    confidence,
    warnings
  };
}

function isVerifiedToolName(name: string): boolean {
  if (!name) return false;
  if (VERIFIED_TOOL_NAMES.has(name)) return true;
  return VERIFIED_TOOL_PREFIXES.some((prefix) => name.startsWith(prefix));
}

function escapeRegex(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

async function verifyRangeStatClaim(
  claim: import("./claim-extraction.ts").ExtractedSpreadsheetClaim,
  ctx: {
    toolExecutor: VerifyAssistantClaimsParams["toolExecutor"];
    hasTool: (name: string) => boolean;
    computeStatisticsByRange: Map<string, { call: any; sanitizedResult?: unknown; rawResult?: any; error?: string }>;
  }
): Promise<NonNullable<VerificationResult["claims"]>[number]> {
  if (claim.kind !== "range_stat") {
    return { claim: claim.source, verified: false, expected: null, actual: null };
  }

  const reference = claim.reference;
  const claimLabel = formatRangeStatClaimLabel(claim.measure, reference, claim.expected);

  if (!reference) {
    return {
      claim: claimLabel,
      verified: false,
      expected: claim.expected,
      actual: null,
      toolEvidence: { error: "missing_reference" }
    };
  }

  if (!ctx.hasTool("compute_statistics")) {
    return {
      claim: claimLabel,
      verified: false,
      expected: claim.expected,
      actual: null,
      toolEvidence: { error: "missing_tool", tool: "compute_statistics" }
    };
  }

  const cached = ctx.computeStatisticsByRange.get(reference);
  if (cached) {
    if (cached.error) {
      return {
        claim: claimLabel,
        verified: false,
        expected: claim.expected,
        actual: null,
        toolEvidence: { call: cached.call, error: cached.error }
      };
    }
    const actual = extractStatisticValue(cached.rawResult, claim.measure);
    const verified = statisticMatchesExpected(claim.measure, actual, claim.expected);
    return {
      claim: claimLabel,
      verified,
      expected: claim.expected,
      actual: actual ?? null,
      toolEvidence: { call: cached.call, result: cached.sanitizedResult }
    };
  }

  // Fallback: should be rare (e.g. missing cache entry). Verify via a single-measure call.
  const toolCall = { name: "compute_statistics", arguments: { range: reference, measures: [claim.measure] } };
  try {
    const result = await ctx.toolExecutor.execute(toolCall);
    const sanitized = sanitizeVerificationToolResult(result);
    const actual = extractStatisticValue(result, claim.measure);
    const verified = statisticMatchesExpected(claim.measure, actual, claim.expected);
    return {
      claim: claimLabel,
      verified,
      expected: claim.expected,
      actual: actual ?? null,
      toolEvidence: { call: toolCall, result: sanitized }
    };
  } catch (error) {
    return {
      claim: claimLabel,
      verified: false,
      expected: claim.expected,
      actual: null,
      toolEvidence: { call: toolCall, error: error instanceof Error ? error.message : String(error) }
    };
  }
}

async function verifyCellValueClaim(
  claim: import("./claim-extraction.ts").ExtractedSpreadsheetClaim,
  ctx: { toolExecutor: VerifyAssistantClaimsParams["toolExecutor"]; hasTool: (name: string) => boolean }
): Promise<NonNullable<VerificationResult["claims"]>[number]> {
  if (claim.kind !== "cell_value") {
    return { claim: claim.source, verified: false, expected: null, actual: null };
  }

  const reference = claim.reference;
  const claimLabel = `value(${reference}) = ${claim.expected}`;

  if (!ctx.hasTool("read_range")) {
    return {
      claim: claimLabel,
      verified: false,
      expected: claim.expected,
      actual: null,
      toolEvidence: { error: "missing_tool", tool: "read_range" }
    };
  }

  const toolCall = {
    name: "read_range",
    arguments: { range: reference, include_formulas: false }
  };

  try {
    const result = await ctx.toolExecutor.execute(toolCall);
    const sanitized = sanitizeVerificationToolResult(result);
    const actual = extractSingleCellValue(result);
    const actualNumber =
      typeof actual === "number" ? actual : typeof actual === "string" ? parseSpreadsheetNumber(actual) : null;
    const verified = numbersApproximatelyEqual(actualNumber, claim.expected);
    return {
      claim: claimLabel,
      verified,
      expected: claim.expected,
      actual: actualNumber ?? null,
      toolEvidence: { call: toolCall, result: sanitized }
    };
  } catch (error) {
    return {
      claim: claimLabel,
      verified: false,
      expected: claim.expected,
      actual: null,
      toolEvidence: { call: toolCall, error: error instanceof Error ? error.message : String(error) }
    };
  }
}

function formatRangeStatClaimLabel(
  measure: import("./claim-extraction.ts").SpreadsheetClaimMeasure,
  reference: string | undefined,
  expected: number
): string {
  const ref = reference ? `(${reference})` : "";
  return `${measure}${ref} = ${expected}`;
}

function sanitizeVerificationToolResult(result: unknown): unknown {
  if (!result || typeof result !== "object" || Array.isArray(result)) return result;
  const { timing, ...rest } = result as any;
  return rest;
}

function extractStatisticValue(result: any, measure: import("./claim-extraction.ts").SpreadsheetClaimMeasure): number | null {
  const stats = result?.data?.statistics;
  if (!stats || typeof stats !== "object") return null;
  const value = (stats as any)[measure];
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function extractSingleCellValue(result: any): unknown {
  const values = result?.data?.values;
  if (!Array.isArray(values) || values.length === 0) return null;
  const firstRow = values[0];
  if (!Array.isArray(firstRow) || firstRow.length === 0) return null;
  return firstRow[0];
}

function numbersApproximatelyEqual(actual: number | null, expected: number): boolean {
  if (actual == null) return false;
  if (!Number.isFinite(actual) || !Number.isFinite(expected)) return false;
  const diff = Math.abs(actual - expected);
  const scale = Math.max(1, Math.abs(actual), Math.abs(expected));
  return diff <= 1e-6 * scale;
}

function statisticMatchesExpected(
  measure: import("./claim-extraction.ts").SpreadsheetClaimMeasure,
  actual: number | null,
  expected: number
): boolean {
  if (actual == null) return false;
  // Count is an integer-valued statistic and should match exactly.
  // Using floating-point tolerance can incorrectly accept off-by-one errors for large ranges.
  if (measure === "count") return actual === expected;
  return numbersApproximatelyEqual(actual, expected);
}
