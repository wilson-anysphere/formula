import type { Logger } from "pino";
import WebSocket from "ws";

import type { AuthContext } from "./auth.js";

type MessageListener = (data: WebSocket.RawData, isBinary: boolean) => void;

type MessageGuardResult =
  | { data: WebSocket.RawData; isBinary: boolean }
  | { drop: true };

type MessageGuard = (
  data: WebSocket.RawData,
  isBinary: boolean
) => MessageGuardResult;

const textDecoder = new TextDecoder();
const textEncoder = new TextEncoder();

// Tracks which websocket "owns" an awareness clientID for a given doc.
// Used to reject attempts to send awareness updates for another live connection.
const awarenessClientIdOwnersByDoc = new Map<string, Map<number, WebSocket>>();

function getAwarenessOwnerMap(docName: string): Map<number, WebSocket> {
  let map = awarenessClientIdOwnersByDoc.get(docName);
  if (!map) {
    map = new Map();
    awarenessClientIdOwnersByDoc.set(docName, map);
  }
  return map;
}

function maybeReleaseDocMap(docName: string): void {
  const map = awarenessClientIdOwnersByDoc.get(docName);
  if (map && map.size === 0) awarenessClientIdOwnersByDoc.delete(docName);
}

function toUint8Array(data: WebSocket.RawData): Uint8Array | null {
  if (typeof data === "string") return null;
  if (Array.isArray(data)) return Buffer.concat(data);
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  // Buffer is a Uint8Array
  if (data instanceof Uint8Array) return data;
  return null;
}

function readVarUint(
  buf: Uint8Array,
  offset: number
): { value: number; offset: number } {
  let value = 0;
  let multiplier = 1;
  while (true) {
    if (offset >= buf.length) {
      throw new Error("Unexpected end of buffer while reading varUint");
    }
    const byte = buf[offset++];
    value += (byte & 0x7f) * multiplier;
    if (byte < 0x80) break;
    multiplier *= 0x80;
    if (!Number.isSafeInteger(value)) {
      throw new Error("varUint exceeds safe integer range");
    }
  }
  return { value, offset };
}

function encodeVarUint(value: number): Uint8Array {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new Error("Invalid varUint value");
  }
  const bytes: number[] = [];
  let v = value;
  while (v > 0x7f) {
    bytes.push(0x80 | (v % 0x80));
    v = Math.floor(v / 0x80);
  }
  bytes.push(v);
  return Uint8Array.from(bytes);
}

function readVarString(
  buf: Uint8Array,
  offset: number
): { value: string; offset: number } {
  const lenRes = readVarUint(buf, offset);
  const length = lenRes.value;
  const start = lenRes.offset;
  const end = start + length;
  if (end > buf.length) {
    throw new Error("Unexpected end of buffer while reading varString");
  }
  const value = textDecoder.decode(buf.subarray(start, end));
  return { value, offset: end };
}

function encodeVarString(value: string): Uint8Array {
  const bytes = textEncoder.encode(value);
  return concatUint8Arrays([encodeVarUint(bytes.length), bytes]);
}

function concatUint8Arrays(arrays: Uint8Array[]): Uint8Array {
  const total = arrays.reduce((sum, arr) => sum + arr.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const arr of arrays) {
    out.set(arr, offset);
    offset += arr.length;
  }
  return out;
}

type AwarenessEntry = { clientID: number; clock: number; stateJSON: string };

function decodeAwarenessUpdate(update: Uint8Array): AwarenessEntry[] {
  let offset = 0;
  const { value: count, offset: afterCount } = readVarUint(update, offset);
  offset = afterCount;

  const entries: AwarenessEntry[] = [];
  for (let i = 0; i < count; i += 1) {
    const clientRes = readVarUint(update, offset);
    const clientID = clientRes.value;
    offset = clientRes.offset;

    const clockRes = readVarUint(update, offset);
    const clock = clockRes.value;
    offset = clockRes.offset;

    const stateRes = readVarString(update, offset);
    const stateJSON = stateRes.value;
    offset = stateRes.offset;

    entries.push({ clientID, clock, stateJSON });
  }
  return entries;
}

function encodeAwarenessUpdate(entries: AwarenessEntry[]): Uint8Array {
  const chunks: Uint8Array[] = [encodeVarUint(entries.length)];
  for (const entry of entries) {
    chunks.push(encodeVarUint(entry.clientID));
    chunks.push(encodeVarUint(entry.clock));
    chunks.push(encodeVarString(entry.stateJSON));
  }
  return concatUint8Arrays(chunks);
}

function sanitizeAwarenessStateJson(stateJSON: string, userId: string): string | null {
  let parsed: unknown;
  try {
    parsed = JSON.parse(stateJSON);
  } catch {
    return null;
  }

  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    return stateJSON;
  }

  const obj = parsed as Record<string, unknown>;

  const presence = obj.presence;
  const hasPresenceObject =
    presence !== null &&
    typeof presence === "object" &&
    !Array.isArray(presence);
  if (hasPresenceObject) {
    (presence as Record<string, unknown>).id = userId;
  }

  const hasUserIdField = Object.prototype.hasOwnProperty.call(obj, "userId");
  if (hasUserIdField) {
    obj.userId = userId;
  }

  const user = obj.user;
  const hasUserObject = user !== null && typeof user === "object" && !Array.isArray(user);
  if (hasUserObject) {
    const userObj = user as Record<string, unknown>;
    if (Object.prototype.hasOwnProperty.call(userObj, "id")) {
      userObj.id = userId;
    }
  }

  // Only rewrite top-level `id` when it looks like a user identity field to avoid
  // clobbering unrelated application identifiers.
  if (
    Object.prototype.hasOwnProperty.call(obj, "id") &&
    (hasPresenceObject || hasUserObject || hasUserIdField)
  ) {
    obj.id = userId;
  }

  return JSON.stringify(obj);
}

function guardSyncMessage(data: Uint8Array, readOnly: boolean): boolean {
  if (!readOnly) return true;

  let offset = 0;
  const outer = readVarUint(data, offset);
  offset = outer.offset;
  if (outer.value !== 0) return true;

  const inner = readVarUint(data, offset);
  // Allow SyncStep1 (0), drop SyncStep2 (1) and Update (2).
  return inner.value === 0;
}

function patchWebSocketMessageHandlers(ws: WebSocket, guard: MessageGuard): void {
  const wrappedListeners = new WeakMap<MessageListener, MessageListener>();

  const originalOn = ws.on.bind(ws);
  const originalAddListener = ws.addListener.bind(ws);
  const originalOnce = ws.once.bind(ws);
  const originalOff = ws.off ? ws.off.bind(ws) : ws.removeListener.bind(ws);
  const originalRemoveListener = ws.removeListener.bind(ws);

  const wrap = (listener: MessageListener): MessageListener => {
    const existing = wrappedListeners.get(listener);
    if (existing) return existing;

    const wrapped: MessageListener = (data, isBinary) => {
      const result = guard(data, isBinary);
      if ("drop" in result) return;
      listener(result.data, result.isBinary);
    };
    wrappedListeners.set(listener, wrapped);
    return wrapped;
  };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ws.on = ((event: any, listener: any) => {
    if (event === "message" && typeof listener === "function") {
      return originalOn(event, wrap(listener as MessageListener));
    }
    return originalOn(event, listener);
  }) as typeof ws.on;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ws.addListener = ((event: any, listener: any) => {
    if (event === "message" && typeof listener === "function") {
      return originalAddListener(event, wrap(listener as MessageListener));
    }
    return originalAddListener(event, listener);
  }) as typeof ws.addListener;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ws.once = ((event: any, listener: any) => {
    if (event === "message" && typeof listener === "function") {
      return originalOnce(event, wrap(listener as MessageListener));
    }
    return originalOnce(event, listener);
  }) as typeof ws.once;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ws.off = ((event: any, listener: any) => {
    if (event === "message" && typeof listener === "function") {
      const wrapped = wrappedListeners.get(listener as MessageListener) ?? listener;
      return originalOff(event, wrapped);
    }
    return originalOff(event, listener);
  }) as typeof ws.off;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ws.removeListener = ((event: any, listener: any) => {
    if (event === "message" && typeof listener === "function") {
      const wrapped = wrappedListeners.get(listener as MessageListener) ?? listener;
      return originalRemoveListener(event, wrapped);
    }
    return originalRemoveListener(event, listener);
  }) as typeof ws.removeListener;
}

export function installYwsSecurity(
  ws: WebSocket,
  params: { docName: string; auth: AuthContext | undefined; logger: Logger }
): void {
  const { docName, auth, logger } = params;
  const role = auth?.role ?? "viewer";
  const userId = auth?.userId ?? "unknown";
  const readOnly = role === "viewer" || role === "commenter";

  let allowedAwarenessClientId: number | null = null;
  let loggedAwarenessSpoofAttempt = false;

  const releaseAwarenessClientId = () => {
    if (allowedAwarenessClientId === null) return;
    const ownerMap = awarenessClientIdOwnersByDoc.get(docName);
    if (!ownerMap) return;

    const currentOwner = ownerMap.get(allowedAwarenessClientId);
    if (currentOwner === ws) {
      ownerMap.delete(allowedAwarenessClientId);
      maybeReleaseDocMap(docName);
    }
    allowedAwarenessClientId = null;
  };

  ws.on("close", releaseAwarenessClientId);

  const claimAwarenessClientId = (clientId: number): boolean => {
    const ownerMap = getAwarenessOwnerMap(docName);
    const existingOwner = ownerMap.get(clientId);
    if (existingOwner && existingOwner !== ws) {
      // Best-effort cleanup for stale connections.
      const isStale =
        existingOwner.readyState === WebSocket.CLOSED ||
        existingOwner.readyState === WebSocket.CLOSING;
      if (!isStale) {
        logger.warn(
          { docName, clientId, userId, role },
          "awareness_client_id_collision"
        );
        ws.close(1008, "awareness clientID collision");
        return false;
      }
      ownerMap.delete(clientId);
    }

    ownerMap.set(clientId, ws);
    allowedAwarenessClientId = clientId;
    return true;
  };

  const guard: MessageGuard = (raw, isBinary) => {
    const message = toUint8Array(raw);
    if (!message) return { drop: true };

    let outerType: number;
    let offset: number;
    try {
      const outer = readVarUint(message, 0);
      outerType = outer.value;
      offset = outer.offset;
    } catch {
      // Malformed message.
      ws.close(1003, "malformed message");
      return { drop: true };
    }

    // 0 = sync, 1 = awareness (y-websocket).
    if (outerType === 0) {
      try {
        if (!guardSyncMessage(message, readOnly)) return { drop: true };
      } catch {
        ws.close(1003, "malformed sync message");
        return { drop: true };
      }
      return { data: raw, isBinary };
    }

    if (outerType !== 1) return { data: raw, isBinary };

    // Awareness anti-spoofing: enforce one clientID per connection and sanitize
    // identity fields to match the authenticated user.
    // y-websocket encodes awareness updates as a length-prefixed Uint8Array:
    // writeVarUint8Array(encoder, awarenessUpdate).
    let payloadLength: number;
    let payloadOffset: number;
    try {
      const lenRes = readVarUint(message, offset);
      payloadLength = lenRes.value;
      payloadOffset = lenRes.offset;
    } catch {
      ws.close(1003, "malformed awareness update");
      return { drop: true };
    }

    const payloadEnd = payloadOffset + payloadLength;
    if (payloadEnd > message.length) {
      ws.close(1003, "malformed awareness update");
      return { drop: true };
    }

    const awarenessUpdate = message.subarray(payloadOffset, payloadEnd);
    let entries: AwarenessEntry[];
    try {
      entries = decodeAwarenessUpdate(awarenessUpdate);
    } catch {
      ws.close(1003, "malformed awareness update");
      return { drop: true };
    }

    const firstClientId = entries[0]?.clientID;
    if (allowedAwarenessClientId === null) {
      if (firstClientId === undefined) return { drop: true };
      if (!claimAwarenessClientId(firstClientId)) return { drop: true };
    }

    const allowedId = allowedAwarenessClientId;
    if (allowedId === null) return { drop: true };

    let sawOtherClientIds = false;
    const filtered: AwarenessEntry[] = [];
    for (const entry of entries) {
      if (entry.clientID !== allowedId) {
        sawOtherClientIds = true;
        continue;
      }

      const sanitizedJson = sanitizeAwarenessStateJson(entry.stateJSON, userId);
      if (sanitizedJson === null) continue;

      filtered.push({ ...entry, stateJSON: sanitizedJson });
    }

    if (sawOtherClientIds && !loggedAwarenessSpoofAttempt) {
      loggedAwarenessSpoofAttempt = true;
      logger.warn(
        { docName, userId, role, allowedId },
        "awareness_spoof_attempt_filtered"
      );
    }

    if (filtered.length === 0) return { drop: true };

    const sanitizedUpdate = encodeAwarenessUpdate(filtered);
    const sanitizedMessage = concatUint8Arrays([
      encodeVarUint(1),
      encodeVarUint(sanitizedUpdate.length),
      sanitizedUpdate,
    ]);
    return { data: Buffer.from(sanitizedMessage), isBinary: true };
  };

  patchWebSocketMessageHandlers(ws, guard);
}
