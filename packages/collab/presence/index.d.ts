export const PRESENCE_VERSION: number;

export interface PresenceCursor {
  row: number;
  col: number;
}

/** Inclusive row/col endpoints. */
export interface PresenceRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export interface PresenceState {
  id: string;
  name: string;
  color: string;
  activeSheet: string;
  cursor: PresenceCursor | null;
  selections: PresenceRange[];
  lastActive: number;
}

export interface SerializedPresenceState {
  v: number;
  id: string;
  name: string;
  color: string;
  sheet: string;
  cursor: PresenceCursor | null;
  selections: PresenceRange[];
  lastActive: number;
}

export function serializePresenceState(state: PresenceState): SerializedPresenceState;
export function deserializePresenceState(payload: unknown): PresenceState | null;

export interface ThrottleOptions {
  now?: () => number;
  setTimeout?: typeof globalThis.setTimeout;
  clearTimeout?: typeof globalThis.clearTimeout;
}

export interface ThrottledFunction {
  (): void;
  cancel(): void;
}

export function throttle(fn: () => void, waitMs: number, options?: ThrottleOptions): ThrottledFunction;

export interface AwarenessChange {
  added: number[];
  updated: number[];
  removed: number[];
}

export interface AwarenessLike {
  clientID: number;
  getStates(): Map<number, unknown>;
  getLocalState?(): unknown;
  setLocalState?(state: unknown | null): void;
  setLocalStateField(field: string, value: unknown): void;
  on?(eventName: "change", handler: (change: AwarenessChange, origin: unknown) => void): void;
  off?(eventName: "change", handler: (change: AwarenessChange, origin: unknown) => void): void;
}

export interface PresenceManagerOptions {
  user: { id: string; name: string; color: string };
  activeSheet: string;
  throttleMs?: number;
  /** Filter out remote presences whose `lastActive` is older than `now() - staleAfterMs`. */
  staleAfterMs?: number;
  now?: () => number;
  setTimeout?: typeof globalThis.setTimeout;
  clearTimeout?: typeof globalThis.clearTimeout;
}

export interface RemotePresenceState extends PresenceState {
  clientId: number;
}

export class PresenceManager {
  constructor(awareness: AwarenessLike, options: PresenceManagerOptions);

  awareness: AwarenessLike;
  now: () => number;
  localPresence: PresenceState;

  destroy(): void;
  setActiveSheet(activeSheet: string): void;
  setUser(user: { id: string; name: string; color: string }): void;
  setCursor(cursor: PresenceCursor | null): void;
  setSelections(selections: Array<PresenceRange | { start: PresenceCursor; end: PresenceCursor }>): void;

  getRemotePresences(options?: { activeSheet?: string; includeOtherSheets?: boolean; staleAfterMs?: number }): RemotePresenceState[];

  subscribe(
    listener: (presences: RemotePresenceState[]) => void,
    opts?: { includeOtherSheets?: boolean },
  ): () => void;
}

export class InMemoryAwarenessHub {
  constructor();

  createAwareness(clientID?: number): AwarenessLike & { destroy(): void };
}
