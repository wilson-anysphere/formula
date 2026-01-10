const SEMVER_RE =
  /^(?<major>0|[1-9]\d*)\.(?<minor>0|[1-9]\d*)\.(?<patch>0|[1-9]\d*)(?:-(?<prerelease>[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?(?:\+(?<build>[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?$/;

function parseSemver(version) {
  const match = SEMVER_RE.exec(version);
  if (!match) return null;
  const { major, minor, patch, prerelease } = match.groups;
  return {
    major: Number(major),
    minor: Number(minor),
    patch: Number(patch),
    prerelease: prerelease ? prerelease.split(".") : null,
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

function compareSemver(a, b) {
  const aParsed = parseSemver(a);
  const bParsed = parseSemver(b);
  if (!aParsed || !bParsed) {
    throw new Error(`Invalid semver compare: "${a}" vs "${b}"`);
  }

  if (aParsed.major !== bParsed.major) return aParsed.major < bParsed.major ? -1 : 1;
  if (aParsed.minor !== bParsed.minor) return aParsed.minor < bParsed.minor ? -1 : 1;
  if (aParsed.patch !== bParsed.patch) return aParsed.patch < bParsed.patch ? -1 : 1;

  const aPre = aParsed.prerelease;
  const bPre = bParsed.prerelease;

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

function maxSemver(versions) {
  const list = versions.filter(isValidSemver);
  if (list.length === 0) return null;
  return list.reduce((best, v) => (compareSemver(v, best) > 0 ? v : best), list[0]);
}

module.exports = {
  compareSemver,
  isValidSemver,
  maxSemver,
  parseSemver,
};
