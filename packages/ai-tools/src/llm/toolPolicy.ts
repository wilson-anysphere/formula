import { TOOL_CAPABILITIES, ToolNameSchema, type ToolName } from "../tool-schema.ts";

export type ToolPolicyMode = "chat" | "agent" | "inline_edit" | "cell_function";

export interface ToolPolicyInput {
  mode: ToolPolicyMode;
  user_text: string;
  has_attachments?: boolean;
  /**
   * Host configuration: only when enabled *and* the user explicitly requests external
   * data will `fetch_external_data` be exposed to the model.
   *
   * Note: `fetch_external_data` also requires an explicit host allowlist
   * (`allowed_external_hosts`). If the allowlist is empty, external fetch is treated
   * as disabled (defense in depth).
   */
  allow_external_data?: boolean;
  /**
   * Explicit host allowlist for `fetch_external_data`.
   *
   * If omitted or empty, external fetch is treated as disabled even when
   * `allow_external_data` is true.
   */
  allowed_external_hosts?: string[];
}

export interface ToolPolicyDecision {
  allowed_tools: ToolName[];
  /**
   * Optional human-readable reasons to aid debugging/audit trails.
   */
  reasons: string[];
}

const ANALYSIS_KEYWORDS_RE =
  /\b(average|avg|mean|median|mode|stdev|std\s*dev|variance|min|max|sum|total|count|correlation|trend|anomal(?:y|ies)|outlier|statistics|stats)\b/i;

const MUTATION_KEYWORDS_RE =
  /\b(set|fill|replace|update|write|put|enter|populate|clear|delete|remove|insert|append|prepend|dedupe|deduplicate|clean)\b/i;

const SORT_KEYWORDS_RE = /\b(sort|re-?order|order\s+by)\b/i;
const FILTER_KEYWORDS_RE = /\bfilter\b/i;
const PIVOT_KEYWORDS_RE = /\bpivot\b/i;
const CHART_KEYWORDS_RE = /\b(chart|plot|graph)\b/i;
const FORMAT_KEYWORDS_RE =
  /\b(format|bold|italic|underline|font|color|colour|highlight|background|number\s*format|currency|percent|percentage|align|alignment)\b/i;

const FORMULA_KEYWORDS_RE = /\b(formula|fill\s+down|autofill)\b/i;

const EXTERNAL_DATA_VERBS_RE = /\b(fetch|import|download|retrieve|pull|get)\b/i;
const EXTERNAL_DATA_CONTEXT_RE = /\b(url|http|https|api|web|internet|external)\b/i;

function toolOrder(): ToolName[] {
  // Preserve a stable ordering for deterministic tests + audits.
  // `ToolNameSchema.options` is insertion-ordered in tool-schema.ts.
  return ToolNameSchema.options as ToolName[];
}

function uniqueInToolOrder(names: Iterable<ToolName>): ToolName[] {
  const set = new Set(names);
  return toolOrder().filter((name) => set.has(name));
}

function toolsByCategory(category: (typeof TOOL_CAPABILITIES)[ToolName]["category"]): ToolName[] {
  return toolOrder().filter((name) => TOOL_CAPABILITIES[name].category === category);
}

const READ_TOOLS = toolsByCategory("read");
const ANALYSIS_TOOLS = toolsByCategory("analysis");
const FORMAT_TOOLS = toolsByCategory("format");
const EXTERNAL_NETWORK_TOOLS = toolsByCategory("external_network");

const INLINE_EDIT_BASE_TOOLS: ToolName[] = ["read_range", "write_cell", "set_range"];

function normalizeText(value: string): string {
  return String(value ?? "").trim();
}

function normalizeAllowedExternalHosts(hosts: ToolPolicyInput["allowed_external_hosts"]): string[] {
  return (hosts ?? [])
    .map((host) => String(host).trim().toLowerCase())
    .filter((host) => host.length > 0);
}

function wantsAnalysis(text: string, hasAttachments: boolean): boolean {
  if (hasAttachments) return true;
  if (!text) return false;
  if (text.includes("?")) return true;
  return ANALYSIS_KEYWORDS_RE.test(text);
}

function wantsMutation(text: string): boolean {
  if (!text) return false;
  if (MUTATION_KEYWORDS_RE.test(text)) return true;
  if (SORT_KEYWORDS_RE.test(text)) return true;
  if (PIVOT_KEYWORDS_RE.test(text)) return true;
  if (CHART_KEYWORDS_RE.test(text)) return true;
  if (FORMAT_KEYWORDS_RE.test(text)) return true;
  // "filter" is ambiguous (could be "filter these rows and show me matching ones"),
  // so we do not treat it as a mutation intent by default.
  return false;
}

function wantsExternalDataFetch(text: string): boolean {
  if (!text) return false;
  const hasUrl = /https?:\/\//i.test(text);
  if (hasUrl && EXTERNAL_DATA_VERBS_RE.test(text)) return true;
  if (EXTERNAL_DATA_VERBS_RE.test(text) && EXTERNAL_DATA_CONTEXT_RE.test(text)) return true;
  if (/\bfetch_external_data\b/i.test(text)) return true;
  return false;
}

function wantsFormatting(text: string): boolean {
  return FORMAT_KEYWORDS_RE.test(text);
}

function wantsChart(text: string): boolean {
  return CHART_KEYWORDS_RE.test(text);
}

function wantsPivot(text: string): boolean {
  return PIVOT_KEYWORDS_RE.test(text);
}

function wantsSort(text: string): boolean {
  return SORT_KEYWORDS_RE.test(text);
}

function wantsFilter(text: string): boolean {
  return FILTER_KEYWORDS_RE.test(text);
}

function wantsFormulaHelpers(text: string): boolean {
  return FORMULA_KEYWORDS_RE.test(text);
}

export function decideAllowedTools(input: ToolPolicyInput): ToolPolicyDecision {
  const mode = input.mode;
  const text = normalizeText(input.user_text);
  const hasAttachments = Boolean(input.has_attachments);
  const allowExternalData = Boolean(input.allow_external_data);
  const allowedExternalHosts = normalizeAllowedExternalHosts(input.allowed_external_hosts);

  const reasons: string[] = [];

  // Side-effect-free modes: never allow mutation/network.
  if (mode === "cell_function") {
    reasons.push("mode=cell_function -> read+compute only");
    return { allowed_tools: uniqueInToolOrder([...READ_TOOLS, ...ANALYSIS_TOOLS]), reasons };
  }

  const externalRequested = wantsExternalDataFetch(text);
  const analysis = wantsAnalysis(text, hasAttachments);
  const mutation = mode === "inline_edit" ? true : wantsMutation(text);

  if (mode === "inline_edit") {
    reasons.push("mode=inline_edit -> no network tools");
    const tools: ToolName[] = [...INLINE_EDIT_BASE_TOOLS];
    if (wantsFormatting(text)) {
      tools.push("apply_formatting");
      reasons.push("formatting intent -> apply_formatting");
    }
    // Inline edit can usually express formula operations via set_range/write_cell, but we
    // optionally expose formula helpers when explicitly requested.
    if (wantsFormulaHelpers(text)) {
      tools.push("apply_formula_column");
      reasons.push("formula intent -> apply_formula_column");
    }
    return { allowed_tools: uniqueInToolOrder(tools), reasons };
  }

  // External data fetch is opt-in and always explicit.
  const externalFetchEnabled = allowExternalData && allowedExternalHosts.length > 0;
  const includeNetwork = externalRequested && externalFetchEnabled;
  if (externalRequested && !allowExternalData) {
    reasons.push("external data requested but allow_external_data=false -> network tools disabled");
  } else if (externalRequested && allowExternalData && allowedExternalHosts.length === 0) {
    reasons.push("external data requested but allowed_external_hosts is empty -> network tools disabled");
  }

  // Chat defaults to analysis unless mutation intent is detected.
  // Agent defaults to allowing mutations unless the prompt is clearly analysis-only.
  const allowMutations = mode === "agent" ? (mutation || !analysis) : mutation;

  const tools: ToolName[] = [];

  // Always allow read tools in non-inline modes; they are necessary for grounding.
  tools.push(...READ_TOOLS);

  if (analysis || mode === "agent") {
    tools.push(...ANALYSIS_TOOLS);
  }

  if (wantsFilter(text)) {
    tools.push("filter_range");
    reasons.push("filter intent -> filter_range");
  }

  if (allowMutations) {
    reasons.push("mutation intent -> mutation tools enabled");
    // Prefer narrower mutation tool surface when possible.
    if (wantsFormatting(text)) {
      tools.push(...FORMAT_TOOLS);
    }

    if (wantsSort(text)) {
      tools.push("sort_range");
    }

    if (wantsPivot(text)) {
      tools.push("create_pivot_table");
    }

    if (wantsChart(text)) {
      tools.push("create_chart");
    }

    const allowWriteEdits = mode === "agent" || MUTATION_KEYWORDS_RE.test(text) || wantsFormulaHelpers(text);
    if (allowWriteEdits) {
      // For direct cell edits, allow writing cells/ranges.
      tools.push("write_cell", "set_range");
      if (wantsFormulaHelpers(text)) {
        tools.push("apply_formula_column");
      }
    }
  } else {
    reasons.push("analysis intent -> read+compute only");
  }

  if (includeNetwork) {
    reasons.push("external data explicitly requested and allowed -> fetch_external_data");
    tools.push(...EXTERNAL_NETWORK_TOOLS);
  }

  return { allowed_tools: uniqueInToolOrder(tools), reasons };
}

export function selectAllowedTools(input: ToolPolicyInput): ToolName[] {
  return decideAllowedTools(input).allowed_tools;
}
