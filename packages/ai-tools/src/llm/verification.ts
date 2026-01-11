export interface VerificationResult {
  /**
   * Heuristic signal that the user's query likely needs spreadsheet data tools
   * (read/compute/filter) to avoid guessing.
   */
  needs_tools: boolean;
  /**
   * Whether any tool calls were executed at all during the run.
   */
  used_tools: boolean;
  /**
   * Whether the run is considered verified (at least one read-only data tool
   * executed successfully, or tools were not needed).
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
}

export interface VerifyToolUsageParams {
  needsTools: boolean;
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

const A1_REFERENCE_RE = /\b[A-Z]{1,3}\d+\b/i;

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

export function verifyToolUsage(params: VerifyToolUsageParams): VerificationResult {
  const needs_tools = params.needsTools;
  const toolCalls = Array.isArray(params.toolCalls) ? params.toolCalls : [];
  const used_tools = toolCalls.length > 0;

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

  if (successfulVerifiedTool) {
    return {
      needs_tools,
      used_tools,
      verified: true,
      confidence: 0.9,
      warnings: []
    };
  }

  const warnings: string[] = [];
  let confidence = 0.4;

  if (!used_tools) {
    warnings.push("No data tools were used; answer may be a guess.");
    confidence = 0.2;
  } else {
    warnings.push("No read/compute tools succeeded; answer may be unverified.");
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

function isVerifiedToolName(name: string): boolean {
  if (!name) return false;
  if (VERIFIED_TOOL_NAMES.has(name)) return true;
  return VERIFIED_TOOL_PREFIXES.some((prefix) => name.startsWith(prefix));
}
