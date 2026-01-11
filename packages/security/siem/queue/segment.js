const SEGMENT_PREFIX = "segment";

export const SEGMENT_STATES = {
  PENDING: "pending",
  OPEN: "open",
  INFLIGHT: "inflight",
  ACKED: "acked",
};

function randomToken() {
  if (globalThis.crypto?.getRandomValues) {
    const bytes = new Uint8Array(8);
    globalThis.crypto.getRandomValues(bytes);
    return Array.from(bytes, (b) => b.toString(16).padStart(2, "0")).join("");
  }

  return Math.random().toString(16).slice(2) + Math.random().toString(16).slice(2);
}

export function createSegmentBaseName({ now = Date.now() } = {}) {
  return `${SEGMENT_PREFIX}-${now}-${randomToken()}`;
}

export function segmentFileName(baseName, state) {
  if (!baseName) throw new Error("segmentFileName requires baseName");
  if (!state || state === SEGMENT_STATES.PENDING) return `${baseName}.jsonl`;
  return `${baseName}.${state}.jsonl`;
}

export function cursorFileName(baseName) {
  if (!baseName) throw new Error("cursorFileName requires baseName");
  return `${baseName}.cursor.json`;
}

export function lockFileName(baseName) {
  if (!baseName) throw new Error("lockFileName requires baseName");
  return `${baseName}.open.lock`;
}

export function parseSegmentFileName(fileName) {
  const match = String(fileName).match(
    /^segment-(?<createdAtMs>\d+)-(?<token>[A-Za-z0-9]+)(?:\.(?<state>open|inflight|acked))?\.jsonl$/
  );
  if (!match?.groups) return null;

  const createdAtMs = Number(match.groups.createdAtMs);
  if (!Number.isFinite(createdAtMs)) return null;

  const baseName = `segment-${match.groups.createdAtMs}-${match.groups.token}`;
  const state = match.groups.state ?? SEGMENT_STATES.PENDING;

  return { baseName, createdAtMs, state };
}
