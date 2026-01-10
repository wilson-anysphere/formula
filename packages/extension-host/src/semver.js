function parseSemver(version) {
  if (typeof version !== "string") return null;
  const match =
    /^(\d+)\.(\d+)\.(\d+)(?:-([0-9A-Za-z-.]+))?(?:\+([0-9A-Za-z-.]+))?$/.exec(version.trim());
  if (!match) return null;
  return {
    major: Number(match[1]),
    minor: Number(match[2]),
    patch: Number(match[3]),
    prerelease: match[4] ?? null,
    build: match[5] ?? null,
    raw: version.trim()
  };
}

function compareSemver(a, b) {
  if (a.major !== b.major) return a.major - b.major;
  if (a.minor !== b.minor) return a.minor - b.minor;
  if (a.patch !== b.patch) return a.patch - b.patch;

  // Ignore prerelease/build ordering for our needs; treat them as equal.
  return 0;
}

function isValidSemver(version) {
  return parseSemver(version) !== null;
}

function satisfies(version, range) {
  const v = parseSemver(version);
  if (!v) return false;

  if (typeof range !== "string" || range.trim().length === 0) return false;
  const r = range.trim();

  if (r === "*") return true;

  if (r.startsWith("^")) {
    const base = parseSemver(r.slice(1));
    if (!base) return false;
    if (v.major !== base.major) return false;
    return compareSemver(v, base) >= 0;
  }

  if (r.startsWith("~")) {
    const base = parseSemver(r.slice(1));
    if (!base) return false;
    if (v.major !== base.major) return false;
    if (v.minor !== base.minor) return false;
    return compareSemver(v, base) >= 0;
  }

  if (r.startsWith(">=")) {
    const base = parseSemver(r.slice(2));
    if (!base) return false;
    return compareSemver(v, base) >= 0;
  }

  if (r.startsWith("<=")) {
    const base = parseSemver(r.slice(2));
    if (!base) return false;
    return compareSemver(v, base) <= 0;
  }

  const exact = parseSemver(r);
  if (exact) return compareSemver(v, exact) === 0;

  return false;
}

module.exports = {
  parseSemver,
  compareSemver,
  isValidSemver,
  satisfies
};

