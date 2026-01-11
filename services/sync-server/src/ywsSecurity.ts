import type { Logger } from "pino";
import WebSocket from "ws";

import { getCellPermissions } from "../../../packages/collab/permissions/index.js";

import type { AuthContext } from "./auth.js";
import { Y } from "./yjs.js";

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

function rawDataByteLength(raw: WebSocket.RawData): number {
  if (typeof raw === "string") return Buffer.byteLength(raw);
  if (Array.isArray(raw)) {
    return raw.reduce((sum, chunk) => sum + chunk.byteLength, 0);
  }
  if (raw instanceof ArrayBuffer) return raw.byteLength;
  // Buffer is a Uint8Array
  if (raw instanceof Uint8Array) return raw.byteLength;
  return 0;
}

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

function readVarUint8Array(
  buf: Uint8Array,
  offset: number
): { value: Uint8Array; offset: number } {
  const lenRes = readVarUint(buf, offset);
  const length = lenRes.value;
  const start = lenRes.offset;
  const end = start + length;
  if (end > buf.length) {
    throw new Error("Unexpected end of buffer while reading varUint8Array");
  }
  return { value: buf.subarray(start, end), offset: end };
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

type CellAddress = { sheetId: string; row: number; col: number };

function parseCellKey(key: string, defaultSheetId: string = "Sheet1"): CellAddress | null {
  if (typeof key !== "string" || key.length === 0) return null;

  const isValidIndex = (value: number): boolean =>
    Number.isSafeInteger(value) && value >= 0;

  const parts = key.split(":");
  if (parts.length === 3) {
    const sheetId = parts[0] || defaultSheetId;
    const row = Number(parts[1]);
    const col = Number(parts[2]);
    if (!isValidIndex(row) || !isValidIndex(col)) return null;
    return { sheetId, row, col };
  }

  // Some internal modules use `${sheetId}:${row},${col}`.
  if (parts.length === 2) {
    const sheetId = parts[0] || defaultSheetId;
    const m = parts[1]?.match(/^(\d+),(\d+)$/);
    if (m) {
      const row = Number(m[1]);
      const col = Number(m[2]);
      if (!isValidIndex(row) || !isValidIndex(col)) return null;
      return { sheetId, row, col };
    }
  }

  const m = key.match(/^r(\d+)c(\d+)$/);
  if (m) {
    const row = Number(m[1]);
    const col = Number(m[2]);
    if (!isValidIndex(row) || !isValidIndex(col)) return null;
    return { sheetId: defaultSheetId, row, col };
  }

  return null;
}

function decodeAwarenessUpdate(
  update: Uint8Array,
  limits: { maxEntries: number; maxStateBytes: number }
): AwarenessEntry[] {
  let offset = 0;
  const { value: count, offset: afterCount } = readVarUint(update, offset);
  offset = afterCount;

  const entries: AwarenessEntry[] = [];
  const maxEntries = Math.max(0, limits.maxEntries);
  const limitCount = Math.min(count, maxEntries);
  for (let i = 0; i < limitCount; i += 1) {
    const clientRes = readVarUint(update, offset);
    const clientID = clientRes.value;
    offset = clientRes.offset;

    const clockRes = readVarUint(update, offset);
    const clock = clockRes.value;
    offset = clockRes.offset;

    const lenRes = readVarUint(update, offset);
    const stateLength = lenRes.value;
    const start = lenRes.offset;
    const end = start + stateLength;
    if (end > update.length) {
      throw new Error("Unexpected end of buffer while reading awareness stateJSON");
    }
    offset = end;

    if (stateLength > limits.maxStateBytes) continue;

    const stateJSON = textDecoder.decode(update.subarray(start, end));

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

function patchWebSocketMessageHandlers(ws: WebSocket, guard: MessageGuard): void {
  const wrappedListeners = new WeakMap<MessageListener, MessageListener>();

  const originalOn = ws.on.bind(ws);
  const originalAddListener = ws.addListener.bind(ws);
  const originalOnce = ws.once.bind(ws);
  const originalPrependListener = ws.prependListener.bind(ws);
  const originalPrependOnceListener = ws.prependOnceListener.bind(ws);
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
  ws.prependListener = ((event: any, listener: any) => {
    if (event === "message" && typeof listener === "function") {
      return originalPrependListener(event, wrap(listener as MessageListener));
    }
    return originalPrependListener(event, listener);
  }) as typeof ws.prependListener;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ws.prependOnceListener = ((event: any, listener: any) => {
    if (event === "message" && typeof listener === "function") {
      return originalPrependOnceListener(event, wrap(listener as MessageListener));
    }
    return originalPrependOnceListener(event, listener);
  }) as typeof ws.prependOnceListener;

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
  params: {
    docName: string;
    auth: AuthContext | undefined;
    logger: Logger;
    ydoc: any;
    limits: {
      maxMessageBytes: number;
      maxAwarenessStateBytes: number;
      maxAwarenessEntries: number;
    };
    enforceRangeRestrictions?: boolean;
  }
): void {
  const {
    docName,
    auth,
    logger,
    ydoc,
    limits,
    enforceRangeRestrictions: enforceRangeRestrictionsConfig,
  } = params;
  const role = auth?.role ?? "viewer";
  const userId = auth?.userId ?? "unknown";
  const readOnly = role === "viewer" || role === "commenter";

  const rangeRestrictions =
    auth?.tokenType === "jwt" && Array.isArray(auth.rangeRestrictions)
      ? auth.rangeRestrictions
      : null;

  // Approach B (Shadow-doc apply + diff):
  // Maintain a per-connection shadow Y.Doc seeded from the current server doc state,
  // and keep it updated via server doc update events. Incoming updates are applied
  // to the shadow doc first to observe which cell keys are touched.
  //
  // Pros: avoids Yjs internal decoding; works for incremental updates against existing docs.
  // Cons: extra CPU/memory for the shadow doc on restricted connections.
  const enforceRangeRestrictions =
    Boolean(enforceRangeRestrictionsConfig) &&
    rangeRestrictions !== null &&
    rangeRestrictions.length > 0;
  const shadowDoc = enforceRangeRestrictions ? new Y.Doc() : null;
  const shadowCells = shadowDoc ? shadowDoc.getMap("cells") : null;

  if (shadowDoc && shadowCells) {
    try {
      Y.applyUpdate(shadowDoc, Y.encodeStateAsUpdate(ydoc));
    } catch {
      // Best-effort: if we can't seed the shadow state, enforcement will fail closed
      // once we start validating incoming updates.
    }

    const applyServerUpdate = (update: Uint8Array) => {
      try {
        Y.applyUpdate(shadowDoc, update);
      } catch {
        // ignore
      }
    };

    if (ydoc && typeof ydoc.on === "function") {
      ydoc.on("update", applyServerUpdate);
    }

    ws.on("close", () => {
      if (ydoc && typeof ydoc.off === "function") {
        try {
          ydoc.off("update", applyServerUpdate);
        } catch {
          // ignore
        }
      }
      try {
        shadowDoc.destroy();
      } catch {
        // ignore
      }
    });
  }

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
    const messageBytes = rawDataByteLength(raw);
    if (limits.maxMessageBytes > 0 && messageBytes > limits.maxMessageBytes) {
      ws.close(1009, "Message too big");
      return { drop: true };
    }

    if (typeof raw === "string") {
      // y-websocket is a binary protocol; reject string messages early.
      ws.close(1003, "binary messages only");
      return { drop: true };
    }

    // ws can deliver text frames as a Buffer with `isBinary=false`. Treat the
    // bytes equivalently regardless of the `isBinary` flag.
    const normalizedRaw: WebSocket.RawData = Array.isArray(raw)
      ? Buffer.concat(raw, messageBytes)
      : raw;

    const message = toUint8Array(normalizedRaw);
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
      let innerType: number;
      try {
        const inner = readVarUint(message, offset);
        innerType = inner.value;
        offset = inner.offset;
      } catch {
        ws.close(1003, "malformed sync message");
        return { drop: true };
      }

      if (readOnly && innerType !== 0) {
        // Allow SyncStep1 (0), drop SyncStep2 (1) and Update (2).
        return { drop: true };
      }

      if (
        enforceRangeRestrictions &&
        (innerType === 1 || innerType === 2) &&
        shadowDoc &&
        shadowCells
      ) {
        let updateBytes: Uint8Array;
        try {
          const updateRes = readVarUint8Array(message, offset);
          updateBytes = updateRes.value;
        } catch {
          ws.close(1003, "malformed sync update");
          return { drop: true };
        }

        const preStateVector = Y.encodeStateVector(shadowDoc);
        const touchedCellKeys = new Set<string>();

        const store = (shadowDoc as any).store as {
          pendingStructs: unknown;
          pendingDs: unknown;
        };

        if (store.pendingStructs || store.pendingDs) {
          logger.warn({ docName, userId, role }, "range_restriction_shadow_pending");
          ws.close(1008, "range restrictions validation failed");
          return { drop: true };
        }

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const observer = (events: any[]) => {
          for (const event of events) {
            const path = event?.path;
            const topCellKey =
              Array.isArray(path) && typeof path[0] === "string" ? path[0] : null;
            if (topCellKey) touchedCellKeys.add(topCellKey);

            const keys = event?.changes?.keys;
            if (!keys) continue;

            // When `observeDeep` is attached to the `cells` map, `event.path` is:
            // - `[]` for changes on the map itself (keys are cell keys)
            // - `[cellKey]` for changes inside a cell's Y.Map (keys are cell fields)
            if (!topCellKey) {
              if (typeof keys.entries === "function") {
                for (const [key] of keys.entries()) {
                  if (typeof key !== "string") continue;
                  touchedCellKeys.add(key);
                }
              } else if (typeof keys.keys === "function") {
                for (const key of keys.keys()) {
                  if (typeof key === "string") {
                    touchedCellKeys.add(key);
                  }
                }
              }
            }
          }
        };

        shadowCells.observeDeep(observer);
        try {
          Y.applyUpdate(shadowDoc, updateBytes);
        } catch (err) {
          const message = err instanceof Error ? err.message : String(err);
          logger.warn({ docName, userId, role, err: message }, "range_restriction_apply_failed");
          ws.close(1008, "range restrictions validation failed");
          return { drop: true };
        } finally {
          shadowCells.unobserveDeep(observer);
        }

        if (store.pendingStructs || store.pendingDs) {
          // The update could not be applied cleanly against our shadow state, so
          // we cannot confidently determine which cells were affected. Fail closed.
          logger.warn({ docName, userId, role }, "range_restriction_update_pending");
          ws.close(1008, "range restrictions validation failed");
          return { drop: true };
        }

        for (const cellKey of touchedCellKeys) {
          const parsed = parseCellKey(cellKey);
          if (!parsed) {
            logger.warn({ docName, userId, role, cellKey }, "range_restriction_unparseable_cell");
            ws.close(1008, "unparseable cell key");
            return { drop: true };
          }

          let canEdit: boolean;
          try {
            ({ canEdit } = getCellPermissions({
              role,
              restrictions: rangeRestrictions,
              userId,
              cell: {
                sheetId: parsed.sheetId,
                row: parsed.row,
                col: parsed.col,
              },
            }));
          } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            logger.warn(
              { docName, userId, role, cellKey, err: message },
              "range_restriction_permission_check_failed"
            );
            ws.close(1008, "range restrictions validation failed");
            return { drop: true };
          }

          if (!canEdit) {
            logger.warn({ docName, userId, role, cellKey }, "permission_violation");
            ws.close(1008, "permission violation");
            return { drop: true };
          }
        }

        // Audit sanitization: ensure `modifiedBy` matches the authenticated user
        // for any touched cell, so clients can't spoof identity by writing a
        // different userId or by leaving a prior user's value intact.
        let needsModifiedByRewrite = false;
        for (const cellKey of touchedCellKeys) {
          const cell = shadowCells.get(cellKey);
          if (!(cell instanceof Y.Map)) continue;
          if (cell.get("modifiedBy") !== userId) {
            needsModifiedByRewrite = true;
            break;
          }
        }

        if (needsModifiedByRewrite) {
          shadowDoc.transact(() => {
            for (const cellKey of touchedCellKeys) {
              const cell = shadowCells.get(cellKey);
              if (!(cell instanceof Y.Map)) continue;
              if (cell.get("modifiedBy") !== userId) {
                (cell as any).set("modifiedBy", userId);
              }
            }
          });

          const sanitizedUpdate = Y.encodeStateAsUpdate(shadowDoc, preStateVector);
          const sanitizedMessage = concatUint8Arrays([
            encodeVarUint(0),
            encodeVarUint(innerType),
            encodeVarUint(sanitizedUpdate.length),
            sanitizedUpdate,
          ]);
          return { data: Buffer.from(sanitizedMessage), isBinary };
        }
      }

      return { data: normalizedRaw, isBinary };
    }

    if (outerType !== 1) return { data: normalizedRaw, isBinary };

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
      entries = decodeAwarenessUpdate(awarenessUpdate, {
        maxEntries: limits.maxAwarenessEntries,
        maxStateBytes: limits.maxAwarenessStateBytes,
      });
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
      if (filtered.length >= limits.maxAwarenessEntries) break;
      if (entry.clientID !== allowedId) {
        sawOtherClientIds = true;
        continue;
      }

      const sanitizedJson = sanitizeAwarenessStateJson(entry.stateJSON, userId);
      if (sanitizedJson === null) continue;
      if (
        limits.maxAwarenessStateBytes > 0 &&
        Buffer.byteLength(sanitizedJson, "utf8") > limits.maxAwarenessStateBytes
      ) {
        continue;
      }

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
    return { data: Buffer.from(sanitizedMessage), isBinary };
  };

  patchWebSocketMessageHandlers(ws, guard);
}
