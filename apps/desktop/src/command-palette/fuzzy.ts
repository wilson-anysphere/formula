export type MatchRange = { start: number; end: number };

export type FuzzyMatch = {
  score: number;
  ranges: MatchRange[];
};

export type CommandLike = {
  commandId: string;
  title: string;
  category: string | null;
  description?: string | null;
  keywords?: string[] | null;
};

export type CommandMatch = {
  score: number;
  titleRanges: MatchRange[];
};

export type CompiledFuzzyQuery = {
  /** Whitespace-normalized query (original case preserved). */
  normalized: string;
  /** Lowercased whitespace-normalized query. */
  normalizedLower: string;
  /**
   * Lowercased query tokens split on spaces.
   *
   * These are precomputed once per query so we don't re-split/re-lowercase
   * inside the per-command matching loop.
   */
  tokens: string[];
};

export function compileFuzzyQuery(query: string): CompiledFuzzyQuery {
  const normalized = normalizeQuery(query);
  const normalizedLower = normalized.toLowerCase();
  const tokens = normalizedLower
    ? normalizedLower
        .split(" ")
        .map((t) => t.trim())
        .filter((t) => t.length > 0)
    : [];
  return { normalized, normalizedLower, tokens };
}

export type PreparedCommandForFuzzy<T extends CommandLike = CommandLike> = T & {
  commandIdLower: string;
  titleLower: string;
  categoryLower: string | null;
  descriptionLower: string | null;
  keywordsJoined: string;
  keywordsJoinedLower: string;
  titleNormalizedLower: string;
};

export function prepareCommandForFuzzy<T extends CommandLike>(command: T): PreparedCommandForFuzzy<T> {
  const title = String(command.title ?? "");
  const category = command.category == null ? null : String(command.category);
  const description = command.description == null ? null : String(command.description);
  const keywordsJoined =
    Array.isArray(command.keywords) && command.keywords.length > 0
      ? command.keywords
          .filter((k) => typeof k === "string" && k.trim() !== "")
          .map((k) => k.trim())
          .join(" ")
      : "";

  return {
    ...command,
    commandIdLower: String(command.commandId ?? "").toLowerCase(),
    titleLower: title.toLowerCase(),
    categoryLower: category ? category.toLowerCase() : null,
    descriptionLower: description ? description.toLowerCase() : null,
    keywordsJoined,
    keywordsJoinedLower: keywordsJoined.toLowerCase(),
    titleNormalizedLower: normalizeQuery(title).toLowerCase(),
  };
}

function isWordBoundary(text: string, index: number): boolean {
  if (index <= 0) return true;
  const prev = text[index - 1] ?? "";
  const cur = text[index] ?? "";

  // Whitespace / separators.
  if (/[\s\-_\/:.]/.test(prev)) return true;

  // camelCase boundary: `fooBar` → boundary at `B`.
  if (prev.toLowerCase() === prev && cur.toLowerCase() !== cur) return true;

  return false;
}

function indicesToRanges(indices: number[]): MatchRange[] {
  if (indices.length === 0) return [];
  const out: MatchRange[] = [];
  let start = indices[0]!;
  let end = start + 1;
  for (let i = 1; i < indices.length; i += 1) {
    const idx = indices[i]!;
    if (idx === end) {
      end += 1;
      continue;
    }
    out.push({ start, end });
    start = idx;
    end = idx + 1;
  }
  out.push({ start, end });
  return out;
}

export function mergeRanges(ranges: MatchRange[]): MatchRange[] {
  if (ranges.length <= 1) return ranges;
  const sorted = [...ranges].sort((a, b) => (a.start !== b.start ? a.start - b.start : a.end - b.end));
  const merged: MatchRange[] = [];

  let current = { ...sorted[0]! };
  for (const next of sorted.slice(1)) {
    if (next.start <= current.end) {
      current.end = Math.max(current.end, next.end);
      continue;
    }
    merged.push(current);
    current = { ...next };
  }
  merged.push(current);
  return merged;
}

/**
 * A small, dependency-free fuzzy matcher.
 *
 * - Case-insensitive subsequence matching
 * - Rewards contiguous runs and word-boundary alignment
 * - Penalizes gaps and longer candidate strings
 */
export function fuzzyMatchToken(query: string, candidate: string): FuzzyMatch | null {
  const qLower = String(query ?? "").trim().toLowerCase();
  if (!qLower) return { score: 0, ranges: [] };

  const text = String(candidate ?? "");
  if (!text) return null;

  return fuzzyMatchTokenPrepared(qLower, text, text.toLowerCase());
}

/**
 * A slightly lower-level variant of {@link fuzzyMatchToken} that accepts a
 * pre-lowercased query and candidate.
 *
 * This is useful for large catalogs (commands/functions) where we want to avoid
 * allocating/lowercasing the same strings on every keystroke.
 */
export function fuzzyMatchTokenPrepared(qLower: string, text: string, lower: string): FuzzyMatch | null {
  if (!qLower) return { score: 0, ranges: [] };
  let score = 0;
  const indices: number[] = [];

  let lastIndex = -1;
  let firstIndex = -1;

  for (let qi = 0; qi < qLower.length; qi += 1) {
    const ch = qLower[qi]!;
    const idx = lower.indexOf(ch, lastIndex + 1);
    if (idx === -1) return null;

    if (firstIndex === -1) firstIndex = idx;
    indices.push(idx);

    const gap = idx - lastIndex - 1;
    const consecutive = idx === lastIndex + 1;
    const boundary = isWordBoundary(text, idx);

    // Base reward for matching a character.
    score += 10;

    // Reward contiguous runs; they usually mean "typed a substring".
    if (consecutive) score += 18;

    // Reward word boundaries (start of word, after separators, camelCase boundaries).
    if (boundary) score += 14;

    // Penalize skipping characters.
    if (gap > 0) score -= gap * 2;

    lastIndex = idx;
  }

  const isSubstring =
    indices.length > 0 && indices.every((idx, i) => idx === indices[0]! + i) && lower.includes(qLower);
  if (isSubstring) score += 30;

  if (firstIndex === 0) score += 20;
  else if (firstIndex > 0 && isWordBoundary(text, firstIndex)) score += 8;

  // Prefer shorter candidates for the same match quality.
  score -= Math.min(40, text.length);

  return { score, ranges: indicesToRanges(indices) };
}

function normalizeQuery(query: string): string {
  return String(query ?? "")
    .trim()
    .replace(/\s+/g, " ");
}

function splitQueryTokens(query: string): string[] {
  return compileFuzzyQuery(query).tokens;
}

type TokenMatch = { score: number; field: "title" | "category" | "id" | "description" | "keywords"; ranges: MatchRange[] };

function bestTokenMatch(tokenLower: string, command: PreparedCommandForFuzzy): TokenMatch | null {
  const titleMatch = fuzzyMatchTokenPrepared(tokenLower, command.title, command.titleLower);
  const categoryMatch =
    command.category && command.categoryLower ? fuzzyMatchTokenPrepared(tokenLower, command.category, command.categoryLower) : null;
  const idMatch = fuzzyMatchTokenPrepared(tokenLower, command.commandId, command.commandIdLower);
  const descriptionMatch =
    command.description && command.descriptionLower ? fuzzyMatchTokenPrepared(tokenLower, command.description, command.descriptionLower) : null;
  const keywordsMatch =
    command.keywordsJoined && command.keywordsJoinedLower
      ? fuzzyMatchTokenPrepared(tokenLower, command.keywordsJoined, command.keywordsJoinedLower)
      : null;

  const candidates: Array<TokenMatch | null> = [
    titleMatch ? { score: titleMatch.score * 3, field: "title", ranges: titleMatch.ranges } : null,
    categoryMatch ? { score: categoryMatch.score * 2, field: "category", ranges: categoryMatch.ranges } : null,
    idMatch ? { score: idMatch.score, field: "id", ranges: idMatch.ranges } : null,
    // Allow searching by description and keywords (lower weight than title/category).
    descriptionMatch ? { score: descriptionMatch.score, field: "description", ranges: descriptionMatch.ranges } : null,
    keywordsMatch ? { score: keywordsMatch.score, field: "keywords", ranges: keywordsMatch.ranges } : null,
  ];

  let best: TokenMatch | null = null;
  for (const cand of candidates) {
    if (!cand) continue;
    if (!best || cand.score > best.score) best = cand;
  }
  return best;
}

export function fuzzyMatchCommandPrepared(query: CompiledFuzzyQuery, command: PreparedCommandForFuzzy): CommandMatch | null {
  const tokens = query.tokens;

  // Empty query: everything matches with neutral score.
  if (tokens.length === 0) return { score: 0, titleRanges: [] };

  let score = 0;
  const titleRanges: MatchRange[] = [];

  for (const token of tokens) {
    const match = bestTokenMatch(token, command);
    if (!match) return null;
    score += match.score;
    if (match.field === "title") titleRanges.push(...match.ranges);
  }

  // Make exact title matches unambiguous (e.g. "Freeze Panes" > "Unfreeze Panes").
  const normQueryLower = query.normalizedLower;
  const normTitleLower = command.titleNormalizedLower;
  if (normQueryLower === normTitleLower) score += 10_000;
  else if (normQueryLower && normTitleLower.startsWith(normQueryLower)) score += 2_500;
  else if (normQueryLower && normTitleLower.includes(normQueryLower)) score += 1_000;

  // Prefer commands with categories when searching by category, but don't dominate title matches.
  if (command.category) score += 5;

  return { score, titleRanges: mergeRanges(titleRanges) };
}

export function fuzzyMatchCommand(query: string, command: CommandLike): CommandMatch | null {
  return fuzzyMatchCommandPrepared(compileFuzzyQuery(query), prepareCommandForFuzzy(command));
}
