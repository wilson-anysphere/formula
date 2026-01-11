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

/**
 * Evaluates a permission grant snapshot (or locked-down default grant) against a request.
 * Intended for deterministic checks inside sandboxes where we cannot share mutable state.
 *
 * @param {any} grant
 * @param {any} request
 * @returns {{allowed: true} | {allowed: false, reason: string}}
 */
export function checkPermissionGrant(grant, request) {
  const effective = grant ?? createLockedDownGrant();

  switch (request?.kind) {
    case "filesystem": {
      const readScopes = new Set([
        ...(effective.filesystem?.read ?? []).map(normalizeScopePath),
        ...(effective.filesystem?.readwrite ?? []).map(normalizeScopePath)
      ]);
      const writeScopes = new Set((effective.filesystem?.readwrite ?? []).map(normalizeScopePath));

      const absPath = normalizeScopePath(request.path);
      const access = request.access === "readwrite" ? "readwrite" : "read";

      if (access === "readwrite") {
        for (const scope of writeScopes) {
          if (isPathWithinScope(absPath, scope)) return { allowed: true };
        }
        return { allowed: false, reason: `Filesystem write access denied for ${absPath}` };
      }

      for (const scope of readScopes) {
        if (isPathWithinScope(absPath, scope)) return { allowed: true };
      }
      return { allowed: false, reason: `Filesystem read access denied for ${absPath}` };
    }
    case "network": {
      const mode = effective.network?.mode ?? "none";
      if (mode === "full") return { allowed: true };
      if (mode === "none") {
        return { allowed: false, reason: `Network access denied for ${request.url}` };
      }

      const allowlist = effective.network?.allowlist ?? [];
      if (isUrlAllowedByAllowlist(request.url, allowlist)) return { allowed: true };
      return { allowed: false, reason: `Network access denied for ${request.url}` };
    }
    case "clipboard":
      return effective.clipboard ? { allowed: true } : { allowed: false, reason: "Clipboard permission denied" };
    case "notifications":
      return effective.notifications
        ? { allowed: true }
        : { allowed: false, reason: "Notifications permission denied" };
    case "automation":
      return effective.automation
        ? { allowed: true }
        : { allowed: false, reason: "Automation permission denied" };
    default:
      return {
        allowed: false,
        reason: `Unknown permission kind: ${String(request?.kind)}`
      };
  }
}
