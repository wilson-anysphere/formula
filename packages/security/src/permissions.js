import path from "node:path";

export function createLockedDownGrant() {
  return {
    filesystem: { read: [], readwrite: [] },
    network: { mode: "none", allowlist: [] },
    clipboard: false,
    notifications: false,
    automation: false
  };
}

/**
 * @param {unknown} input
 * @returns {input is {type: string, id: string}}
 */
export function isPrincipal(input) {
  return (
    typeof input === "object" &&
    input !== null &&
    typeof input.type === "string" &&
    typeof input.id === "string"
  );
}

export function principalKey(principal) {
  if (!isPrincipal(principal)) throw new TypeError("Invalid principal");
  return `${principal.type}:${principal.id}`;
}

export function normalizeScopePath(scopePath) {
  return path.resolve(scopePath);
}

export function isPathWithinScope(targetPath, scopePath) {
  const absTarget = path.resolve(targetPath);
  const absScope = path.resolve(scopePath);
  const relative = path.relative(absScope, absTarget);
  return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
}

function normalizeNetworkAllowlistEntry(entry) {
  return String(entry).trim();
}

export function isUrlAllowedByAllowlist(urlString, allowlist) {
  const url = new URL(urlString);
  const origin = url.origin;
  const host = url.hostname;

  for (const rawEntry of allowlist) {
    const entry = normalizeNetworkAllowlistEntry(rawEntry);
    if (!entry) continue;

    // Origin match, e.g. "https://api.example.com"
    if (entry.includes("://")) {
      if (origin === entry) return true;
      continue;
    }

    // Host wildcard match, e.g. "*.example.com"
    if (entry.startsWith("*.")) {
      const suffix = entry.slice(2);
      if (host === suffix) return true;
      if (host.endsWith(`.${suffix}`)) return true;
      continue;
    }

    // Host exact match, e.g. "example.com"
    if (host === entry) return true;
  }

  return false;
}
