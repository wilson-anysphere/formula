const GLOBAL_KEY = "__formula_semver_range__";

const SEMVER_RE =
  /^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?(?:\+([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?$/;

function parseSemver(version) {
  if (typeof version !== "string") return null;
  const raw = version.trim();
  const match = SEMVER_RE.exec(raw);
  if (!match) return null;
  const major = match[1];
  const minor = match[2];
  const patch = match[3];
  const prerelease = match[4];
  const build = match[5];
  return {
    major: Number(major),
    minor: Number(minor),
    patch: Number(patch),
    prerelease: prerelease ? prerelease.split(".") : null,
    build: build ? build.split(".") : null,
    raw,
  };
}

function isValidSemver(version) {
  return Boolean(parseSemver(version));
}

function compareIdentifiers(a, b) {
  const aNum = /^[0-9]+$/.test(a);
  const bNum = /^[0-9]+$/.test(b);
  if (aNum && bNum) {
    const aVal = Number(a);
    const bVal = Number(b);
    if (aVal !== bVal) return aVal < bVal ? -1 : 1;
    return 0;
  }

  if (aNum !== bNum) {
    // Numeric identifiers have lower precedence than non-numeric.
    return aNum ? -1 : 1;
  }

  if (a === b) return 0;
  return a < b ? -1 : 1;
}

/**
 * @param {{ major: number, minor: number, patch: number, prerelease: string[] | null }} a
 * @param {{ major: number, minor: number, patch: number, prerelease: string[] | null }} b
 */
function compareSemverParsed(a, b) {
  if (a.major !== b.major) return a.major < b.major ? -1 : 1;
  if (a.minor !== b.minor) return a.minor < b.minor ? -1 : 1;
  if (a.patch !== b.patch) return a.patch < b.patch ? -1 : 1;

  const aPre = a.prerelease;
  const bPre = b.prerelease;

  // A version without prerelease has higher precedence than one with prerelease.
  if (!aPre && !bPre) return 0;
  if (!aPre) return 1;
  if (!bPre) return -1;

  const max = Math.max(aPre.length, bPre.length);
  for (let i = 0; i < max; i++) {
    const aId = aPre[i];
    const bId = bPre[i];
    if (aId === undefined) return -1;
    if (bId === undefined) return 1;
    const cmp = compareIdentifiers(aId, bId);
    if (cmp !== 0) return cmp;
  }

  return 0;
}

function compareSemver(a, b) {
  const aParsed = parseSemver(a);
  const bParsed = parseSemver(b);
  if (!aParsed || !bParsed) {
    throw new Error(`Invalid semver compare: "${a}" vs "${b}"`);
  }
  return compareSemverParsed(aParsed, bParsed);
}

function maxSemver(versions) {
  if (!Array.isArray(versions)) return null;
  const parsed = [];
  for (const v of versions) {
    const p = parseSemver(v);
    if (p) parsed.push(p);
  }
  if (parsed.length === 0) return null;

  let best = parsed[0];
  for (const p of parsed.slice(1)) {
    if (compareSemverParsed(p, best) > 0) best = p;
  }
  return best.raw;
}

function makeSemver(major, minor, patch) {
  return { major, minor, patch, prerelease: null, build: null, raw: `${major}.${minor}.${patch}` };
}

function caretUpperBound(base) {
  if (base.major > 0) return makeSemver(base.major + 1, 0, 0);
  if (base.minor > 0) return makeSemver(0, base.minor + 1, 0);
  return makeSemver(0, 0, base.patch + 1);
}

function tildeUpperBound(base) {
  return makeSemver(base.major, base.minor + 1, 0);
}

function parseComparatorToken(token) {
  if (token === "*") return () => true;

  if (token.startsWith("^")) {
    const base = parseSemver(token.slice(1));
    if (!base) return null;
    const upper = caretUpperBound(base);
    return (v) => compareSemverParsed(v, base) >= 0 && compareSemverParsed(v, upper) < 0;
  }

  if (token.startsWith("~")) {
    const base = parseSemver(token.slice(1));
    if (!base) return null;
    const upper = tildeUpperBound(base);
    return (v) => compareSemverParsed(v, base) >= 0 && compareSemverParsed(v, upper) < 0;
  }

  const match = /^(>=|<=|>|<|=)(.+)$/.exec(token);
  if (match) {
    const op = match[1];
    const base = parseSemver(match[2]);
    if (!base) return null;
    if (op === ">=") return (v) => compareSemverParsed(v, base) >= 0;
    if (op === ">") return (v) => compareSemverParsed(v, base) > 0;
    if (op === "<=") return (v) => compareSemverParsed(v, base) <= 0;
    if (op === "<") return (v) => compareSemverParsed(v, base) < 0;
    if (op === "=") return (v) => compareSemverParsed(v, base) === 0;
    return null;
  }

  const exact = parseSemver(token);
  if (exact) return (v) => compareSemverParsed(v, exact) === 0;
  return null;
}

function satisfies(version, range) {
  const v = parseSemver(version);
  if (!v) return false;

  if (typeof range !== "string") return false;
  const rawRange = range.trim();
  if (!rawRange) return false;
  if (rawRange === "*") return true;

  const groups = rawRange.split(/\s*\|\|\s*/);
  for (const group of groups) {
    const rawGroup = group.trim();
    if (!rawGroup) continue;
    if (rawGroup === "*") return true;

    const tokens = rawGroup.split(/\s+/).map((t) => t.trim()).filter(Boolean);
    if (tokens.length === 0) continue;

    let ok = true;
    for (const token of tokens) {
      const predicate = parseComparatorToken(token);
      if (!predicate || !predicate(v)) {
        ok = false;
        break;
      }
    }
    if (ok) return true;
  }

  return false;
}

const exportsObj = {
  compareSemver,
  isValidSemver,
  maxSemver,
  parseSemver,
  satisfies,
};

// ESM-friendly export: register on globalThis so browser runtimes can import the `.mjs` wrapper and
// access named exports without requiring CommonJS interop.
try {
  if (typeof globalThis !== "undefined") {
    globalThis[GLOBAL_KEY] = exportsObj;
  }
} catch {
  // ignore
}

// CommonJS export (Node).
try {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = exportsObj;
  }
} catch {
  // ignore
}
