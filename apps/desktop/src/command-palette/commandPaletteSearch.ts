import type { CommandContribution } from "../extensions/commandRegistry.js";
import { getFunctionSignature } from "../formula-bar/highlight/functionSignatures.js";
import { normalizeFormulaLocaleId, normalizeLocaleId } from "../spreadsheet/formulaLocale.js";

import FUNCTION_NAMES from "../../../../shared/functionNames.mjs";

import {
  compileFuzzyQuery,
  fuzzyMatchCommandPrepared,
  fuzzyMatchTokenPrepared,
  type MatchRange,
  type PreparedCommandForFuzzy,
} from "./fuzzy.js";

export type CommandPaletteCommandResult = {
  kind: "command";
  command: PreparedCommandForFuzzy<CommandContribution>;
  score: number;
  titleRanges: MatchRange[];
};

export type CommandPaletteFunctionResult = {
  kind: "function";
  name: string;
  signature?: string;
  summary?: string;
  score: number;
  matchRanges: MatchRange[];
};

export type CommandPaletteResult = CommandPaletteCommandResult | CommandPaletteFunctionResult;

export type CommandPaletteSection = {
  title: "RECENT" | "COMMANDS" | "FUNCTIONS";
  results: CommandPaletteResult[];
};

type FunctionSignature = NonNullable<ReturnType<typeof getFunctionSignature>>;

type PreparedFunctionForSearch = { name: string; nameLower: string; nameLowerNormalized: string };

const FUNCTIONS: PreparedFunctionForSearch[] = (Array.isArray(FUNCTION_NAMES) ? FUNCTION_NAMES : [])
  .map((name) => String(name ?? "").trim())
  .filter((name) => name.length > 0)
  .map((name) => {
    const nameLower = name.toLowerCase();
    // Remove punctuation so dotted function names like `RANK.EQ` are searchable by `rankeq`.
    const nameLowerNormalized = nameLower.replace(/[^a-z0-9_]/g, "");
    return { name, nameLower, nameLowerNormalized };
  });

export function buildCommandPaletteSections(opts: {
  query: string;
  commands: PreparedCommandForFuzzy<CommandContribution>[];
  recentIds?: string[];
  limits: { maxResults: number; maxResultsPerGroup: number };
}): CommandPaletteSection[] {
  const query = normalizeQuery(opts.query);
  const limits = normalizeLimits(opts.limits);
  const recentIds = Array.isArray(opts.recentIds) ? opts.recentIds : [];

  if (!query) {
    return buildEmptyQuerySections(opts.commands, recentIds, limits);
  }

  return buildQuerySections(query, opts.commands, limits);
}

/**
 * Returns palette-ready function matches for a query.
 *
 * This is used by the command palette UI to surface spreadsheet functions alongside commands.
 */
export function searchFunctionResults(query: string, opts: { limit: number }): CommandPaletteFunctionResult[] {
  const normalized = String(query ?? "")
    .trim()
    .replace(/^=+/, "")
    .replace(/\s+/g, "")
    // Users often type formulas like `SUM(` or `=SUM(`; ignore non-word characters.
    .replace(/[^A-Za-z0-9_]/g, "");
  const limit = Math.max(0, Math.floor(opts.limit));
  return scoreFunctionResults(normalized.toLowerCase(), limit);
}

function normalizeQuery(query: string): string {
  // Allow users to type formulas like `=sum` and still find SUM.
  return String(query ?? "")
    .trim()
    .replace(/^=+/, "")
    .replace(/\s+/g, " ");
}

function normalizeLimits(limits: { maxResults: number; maxResultsPerGroup: number }): { maxResults: number; maxResultsPerGroup: number } {
  return {
    maxResults: Math.max(0, Math.floor(limits.maxResults)),
    maxResultsPerGroup: Math.max(1, Math.floor(limits.maxResultsPerGroup)),
  };
}

function buildEmptyQuerySections(
  commands: PreparedCommandForFuzzy<CommandContribution>[],
  recentIds: string[],
  limits: { maxResults: number; maxResultsPerGroup: number },
): CommandPaletteSection[] {
  const byId = new Map(commands.map((cmd) => [cmd.commandId, cmd]));
  const recents: CommandPaletteCommandResult[] = [];
  const recentSet = new Set<string>();

  const recentLimit = Math.min(limits.maxResultsPerGroup, limits.maxResults);
  for (const id of recentIds) {
    if (recents.length >= recentLimit) break;
    const cmd = byId.get(id);
    if (!cmd) continue;
    recentSet.add(id);
    recents.push({ kind: "command", command: cmd, score: 0, titleRanges: [] });
  }

  // Show remaining commands alphabetically. This avoids "random-looking" order when no
  // query is present, and keeps the section stable across refreshes.
  const remaining = commands
    .filter((cmd) => !recentSet.has(cmd.commandId))
    .slice()
    .sort((a, b) => a.title.localeCompare(b.title));

  const remainingSlots = Math.max(0, limits.maxResults - recents.length);
  const commandLimit = Math.min(limits.maxResultsPerGroup, remainingSlots);
  const commandResults: CommandPaletteCommandResult[] = remaining.slice(0, commandLimit).map((cmd) => ({
    kind: "command",
    command: cmd,
    score: 0,
    titleRanges: [],
  }));

  const sections: CommandPaletteSection[] = [];
  if (recents.length) sections.push({ title: "RECENT", results: recents });
  if (commandResults.length) sections.push({ title: "COMMANDS", results: commandResults });
  return sections;
}

function buildQuerySections(
  query: string,
  commands: PreparedCommandForFuzzy<CommandContribution>[],
  limits: { maxResults: number; maxResultsPerGroup: number },
): CommandPaletteSection[] {
  const compiled = compileFuzzyQuery(query);
  const functionQueryLower = compiled.normalizedLower.replace(/\s+/g, "").replace(/[^a-z0-9_]/g, "");

  const commandResults = scoreCommandResults(compiled, commands, limits.maxResults).slice(0, limits.maxResultsPerGroup);
  const functionResults = scoreFunctionResults(functionQueryLower, limits.maxResults).slice(0, limits.maxResultsPerGroup);

  const bestCommand = commandResults[0]?.score ?? Number.NEGATIVE_INFINITY;
  const bestFunction = functionResults[0]?.score ?? Number.NEGATIVE_INFINITY;
  const functionsFirst = bestFunction > bestCommand;

  const ordered: Array<{ title: CommandPaletteSection["title"]; results: CommandPaletteResult[] }> = functionsFirst
    ? [
        { title: "FUNCTIONS", results: functionResults },
        { title: "COMMANDS", results: commandResults },
      ]
    : [
        { title: "COMMANDS", results: commandResults },
        { title: "FUNCTIONS", results: functionResults },
      ];

  const sections: CommandPaletteSection[] = [];
  let remainingSlots = limits.maxResults;
  for (const group of ordered) {
    if (remainingSlots <= 0) break;
    if (group.results.length === 0) continue;
    const slice = group.results.slice(0, remainingSlots);
    if (slice.length === 0) continue;
    sections.push({ title: group.title, results: slice });
    remainingSlots -= slice.length;
  }

  return sections;
}

function scoreCommandResults(
  compiled: ReturnType<typeof compileFuzzyQuery>,
  commands: PreparedCommandForFuzzy<CommandContribution>[],
  limit: number,
): CommandPaletteCommandResult[] {
  // Keep only top N matches to avoid sorting huge arrays.
  const top: CommandPaletteCommandResult[] = [];

  const isBetter = (a: CommandPaletteCommandResult, b: CommandPaletteCommandResult): boolean => {
    if (a.score !== b.score) return a.score > b.score;
    return a.command.title.localeCompare(b.command.title) < 0;
  };

  const worstIndex = (): number => {
    let worst = 0;
    for (let i = 1; i < top.length; i += 1) {
      if (isBetter(top[worst]!, top[i]!)) worst = i;
    }
    return worst;
  };

  for (const cmd of commands) {
    const match = fuzzyMatchCommandPrepared(compiled, cmd);
    if (!match) continue;
    const candidate: CommandPaletteCommandResult = {
      kind: "command",
      command: cmd,
      score: match.score,
      titleRanges: match.titleRanges,
    };

    if (top.length < limit) {
      top.push(candidate);
      continue;
    }

    const idx = worstIndex();
    if (isBetter(candidate, top[idx]!)) {
      top[idx] = candidate;
    }
  }

  top.sort((a, b) => {
    if (a.score !== b.score) return b.score - a.score;
    return a.command.title.localeCompare(b.command.title);
  });

  return top;
}

type FunctionMatch = { name: string; score: number; matchRanges: MatchRange[] };

function scoreFunctionResults(queryLower: string, limit: number): CommandPaletteFunctionResult[] {
  const trimmed = queryLower.trim();
  const cappedLimit = Math.max(0, Math.floor(limit));
  if (!trimmed || cappedLimit === 0) return [];
  const normalizedQuery = trimmed.replace(/[^a-z0-9_]/g, "");
  if (!normalizedQuery) return [];

  const functions = FUNCTIONS;

  // Keep only the top-N matches so we don't allocate/sort huge arrays for large
  // function catalogs.
  const top: FunctionMatch[] = [];

  const isBetter = (a: FunctionMatch, b: FunctionMatch): boolean => {
    if (a.score !== b.score) return a.score > b.score;
    return a.name.localeCompare(b.name) < 0;
  };

  const worstIndex = (): number => {
    let worst = 0;
    for (let i = 1; i < top.length; i += 1) {
      if (isBetter(top[worst]!, top[i]!)) worst = i;
    }
    return worst;
  };

  for (const fn of functions) {
    const match = fuzzyMatchTokenPrepared(normalizedQuery, fn.name, fn.nameLower);
    if (!match) continue;

    let score = match.score;

    // Make exact/prefix matches unambiguous (helps function names beat similarly-named commands like "AutoSum").
    if (fn.nameLowerNormalized === normalizedQuery) score += 10_000;
    else if (fn.nameLowerNormalized.startsWith(normalizedQuery)) score += 2_500;
    else if (fn.nameLowerNormalized.includes(normalizedQuery)) score += 1_000;

    const candidate: FunctionMatch = { name: fn.name, score, matchRanges: match.ranges };

    if (top.length < cappedLimit) {
      top.push(candidate);
      continue;
    }

    const idx = worstIndex();
    if (isBetter(candidate, top[idx]!)) {
      top[idx] = candidate;
    }
  }

  top.sort((a, b) => {
    if (a.score !== b.score) return b.score - a.score;
    return a.name.localeCompare(b.name);
  });

  const localeId = (() => {
    try {
      const raw = typeof document !== "undefined" ? document.documentElement?.lang : "";
      return String(raw ?? "").trim() || "en-US";
    } catch {
      return "en-US";
    }
  })();
  const argSeparator = inferArgSeparator(localeId);

  return top.map((match) => {
    const sig = getFunctionSignature(match.name, { localeId });
    const signature = sig ? formatSignature(sig, argSeparator) : undefined;
    const summary = sig?.summary?.trim() ? sig.summary.trim() : undefined;

    return {
      kind: "function",
      name: match.name,
      signature,
      summary,
      score: match.score,
      matchRanges: match.matchRanges,
    };
  });
}

const ARG_SEPARATOR_CACHE = new Map<string, string>();

function inferArgSeparator(localeId: string): string {
  const locale = normalizeFormulaLocaleId(localeId) ?? normalizeLocaleId(localeId) ?? "en-US";
  const cached = ARG_SEPARATOR_CACHE.get(locale);
  if (cached) return cached;

  try {
    // Excel typically uses `;` as the list/arg separator when the decimal separator is `,`.
    const parts = new Intl.NumberFormat(locale).formatToParts(1.1);
    const decimal = parts.find((p) => p.type === "decimal")?.value ?? ".";
    const sep = decimal === "," ? "; " : ", ";
    ARG_SEPARATOR_CACHE.set(locale, sep);
    return sep;
  } catch {
    return ", ";
  }
}

function formatSignature(sig: FunctionSignature, argSeparator: string): string {
  const params = sig.params.map((param) => (param.optional ? `[${param.name}]` : param.name)).join(argSeparator);
  return `${sig.name}(${params})`;
}
