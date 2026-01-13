import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryAwarenessHub, PresenceManager, serializePresenceState } from "../index.js";

class FakeScheduler {
  constructor() {
    this.nowMs = 0;
    this._tasks = [];
    this._nextId = 1;
  }

  now() {
    return this.nowMs;
  }

  setTimeout(cb, delayMs) {
    const id = this._nextId++;
    this._tasks.push({ id, runAt: this.nowMs + delayMs, cb });
    this._tasks.sort((a, b) => a.runAt - b.runAt);
    return id;
  }

  clearTimeout(id) {
    this._tasks = this._tasks.filter((task) => task.id !== id);
  }

  advance(ms) {
    this.nowMs += ms;
    this._runDue();
  }

  _runDue() {
    while (true) {
      const task = this._tasks[0];
      if (!task || task.runAt > this.nowMs) return;
      this._tasks.shift();
      task.cb();
    }
  }
}

test("PresenceManager evicts stale remote presences via staleAfterMs", () => {
  const hub = new InMemoryAwarenessHub();
  const awarenessA = hub.createAwareness(1);
  const awarenessB = hub.createAwareness(2);

  let nowMs = 0;
  const now = () => nowMs;

  const presenceA = new PresenceManager(awarenessA, {
    user: { id: "u1", name: "Ada", color: "#ff2d55" },
    activeSheet: "Sheet1",
    throttleMs: 0,
    staleAfterMs: 500,
    now,
  });

  const presenceB = new PresenceManager(awarenessB, {
    user: { id: "u2", name: "Grace", color: "#4c8bf5" },
    activeSheet: "Sheet1",
    throttleMs: 0,
    now,
  });

  assert.equal(presenceA.getRemotePresences().length, 1);

  nowMs = 1_000;
  assert.deepEqual(presenceA.getRemotePresences(), []);

  presenceB.setCursor({ row: 1, col: 1 });
  assert.equal(presenceA.getRemotePresences().length, 1);
});

test("PresenceManager.subscribe evicts stale remote presences without remote awareness updates", () => {
  const scheduler = new FakeScheduler();
  const hub = new InMemoryAwarenessHub();
  const awarenessA = hub.createAwareness(1);
  const awarenessB = hub.createAwareness(2);

  const presenceA = new PresenceManager(awarenessA, {
    user: { id: "u1", name: "Ada", color: "#ff2d55" },
    activeSheet: "Sheet1",
    throttleMs: 0,
    staleAfterMs: 100,
    now: () => scheduler.now(),
    setTimeout: scheduler.setTimeout.bind(scheduler),
    clearTimeout: scheduler.clearTimeout.bind(scheduler),
  });

  new PresenceManager(awarenessB, {
    user: { id: "u2", name: "Grace", color: "#4c8bf5" },
    activeSheet: "Sheet1",
    throttleMs: 0,
    now: () => scheduler.now(),
    setTimeout: scheduler.setTimeout.bind(scheduler),
    clearTimeout: scheduler.clearTimeout.bind(scheduler),
  });

  /** @type {string[][]} */
  const updates = [];
  const unsubscribe = presenceA.subscribe((presences) => {
    updates.push(presences.map((presence) => presence.id));
  });

  assert.deepEqual(updates, [["u2"]]);

  scheduler.advance(101);
  assert.deepEqual(updates, [["u2"], []]);

  unsubscribe();
});

test("PresenceManager.subscribe eviction timer accounts for remote users on other sheets", () => {
  const scheduler = new FakeScheduler();
  const hub = new InMemoryAwarenessHub();
  const awarenessA = hub.createAwareness(1);
  const awarenessB = hub.createAwareness(2);

  const presenceA = new PresenceManager(awarenessA, {
    user: { id: "u1", name: "Ada", color: "#ff2d55" },
    activeSheet: "Sheet1",
    throttleMs: 0,
    staleAfterMs: 100,
    now: () => scheduler.now(),
    setTimeout: scheduler.setTimeout.bind(scheduler),
    clearTimeout: scheduler.clearTimeout.bind(scheduler),
  });

  new PresenceManager(awarenessB, {
    user: { id: "u2", name: "Grace", color: "#4c8bf5" },
    activeSheet: "Sheet2",
    throttleMs: 0,
    now: () => scheduler.now(),
    setTimeout: scheduler.setTimeout.bind(scheduler),
    clearTimeout: scheduler.clearTimeout.bind(scheduler),
  });

  /** @type {string[][]} */
  const updates = [];
  const unsubscribe = presenceA.subscribe(
    (presences) => {
      updates.push(presences.map((presence) => presence.id));
    },
    { includeOtherSheets: true },
  );

  assert.deepEqual(updates, [["u2"]]);

  scheduler.advance(101);
  assert.deepEqual(updates, [["u2"], []]);

  unsubscribe();
});

test("PresenceManager.subscribe eviction timer accounts for remote users on other sheets (legacy getRemotePresences usage)", () => {
  const scheduler = new FakeScheduler();
  const hub = new InMemoryAwarenessHub();
  const awarenessA = hub.createAwareness(1);
  const awarenessB = hub.createAwareness(2);

  const presenceA = new PresenceManager(awarenessA, {
    user: { id: "u1", name: "Ada", color: "#ff2d55" },
    activeSheet: "Sheet1",
    throttleMs: 0,
    staleAfterMs: 100,
    now: () => scheduler.now(),
    setTimeout: scheduler.setTimeout.bind(scheduler),
    clearTimeout: scheduler.clearTimeout.bind(scheduler),
  });

  new PresenceManager(awarenessB, {
    user: { id: "u2", name: "Grace", color: "#4c8bf5" },
    activeSheet: "Sheet2",
    throttleMs: 0,
    now: () => scheduler.now(),
    setTimeout: scheduler.setTimeout.bind(scheduler),
    clearTimeout: scheduler.clearTimeout.bind(scheduler),
  });

  /** @type {string[][]} */
  const updates = [];
  const unsubscribe = presenceA.subscribe(() => {
    updates.push(presenceA.getRemotePresences({ includeOtherSheets: true }).map((presence) => presence.id));
  });

  assert.deepEqual(updates, [["u2"]]);

  scheduler.advance(101);
  assert.deepEqual(updates, [["u2"], []]);

  unsubscribe();
});

test("PresenceManager.subscribe schedules eviction even when now changes between snapshots", () => {
  const scheduler = new FakeScheduler();
  const hub = new InMemoryAwarenessHub();
  const awarenessA = hub.createAwareness(1);
  const awarenessB = hub.createAwareness(2);

  const now = () => scheduler.nowMs++;

  const presenceA = new PresenceManager(awarenessA, {
    user: { id: "u1", name: "Ada", color: "#ff2d55" },
    activeSheet: "Sheet1",
    throttleMs: 0,
    staleAfterMs: 0,
    now,
    setTimeout: scheduler.setTimeout.bind(scheduler),
    clearTimeout: scheduler.clearTimeout.bind(scheduler),
  });

  awarenessB.setLocalStateField(
    "presence",
    serializePresenceState({
      id: "u2",
      name: "Grace",
      color: "#4c8bf5",
      activeSheet: "Sheet1",
      cursor: null,
      selections: [],
      lastActive: 1,
    }),
  );

  /** @type {string[][]} */
  const updates = [];
  const unsubscribe = presenceA.subscribe((presences) => {
    updates.push(presences.map((presence) => presence.id));
  });

  assert.deepEqual(updates, [["u2"]]);

  // `now` advances on each call; without using a single snapshot for both the
  // subscription callback and eviction scheduling, the eviction timer might not
  // be scheduled at all.
  scheduler.advance(0);
  assert.deepEqual(updates, [["u2"], []]);

  unsubscribe();
});

test("PresenceManager.getRemotePresences can include users on other sheets", () => {
  const hub = new InMemoryAwarenessHub();
  const awarenessA = hub.createAwareness(1);
  const awarenessB = hub.createAwareness(2);
  const awarenessC = hub.createAwareness(3);

  const presenceA = new PresenceManager(awarenessA, {
    user: { id: "u1", name: "Ada", color: "#ff2d55" },
    activeSheet: "Sheet1",
  });

  new PresenceManager(awarenessB, {
    user: { id: "u2", name: "Grace", color: "#4c8bf5" },
    activeSheet: "Sheet2",
  });

  new PresenceManager(awarenessC, {
    user: { id: "u3", name: "Linus", color: "#34c759" },
    activeSheet: "Sheet1",
  });

  const activeSheetPresences = presenceA.getRemotePresences();
  assert.equal(activeSheetPresences.length, 1);
  assert.equal(activeSheetPresences[0].id, "u3");

  const allPresences = presenceA.getRemotePresences({ includeOtherSheets: true });
  assert.equal(allPresences.length, 2);

  const byId = new Map(allPresences.map((presence) => [presence.id, presence]));
  assert.equal(byId.get("u2")?.activeSheet, "Sheet2");
  assert.equal(byId.get("u3")?.activeSheet, "Sheet1");
});

test("PresenceManager.getRemotePresences returns results in a stable order", () => {
  const now = () => 0;

  const makeState = (id) => ({
    id,
    name: id,
    color: "#000000",
    activeSheet: "Sheet1",
    cursor: null,
    selections: [],
    lastActive: now(),
  });

  const statesA = new Map([
    [1, { presence: serializePresenceState(makeState("local")) }],
    [2, { presence: serializePresenceState(makeState("u2")) }],
    [4, { presence: serializePresenceState(makeState("u1")) }],
    [3, { presence: serializePresenceState(makeState("u1")) }],
  ]);

  const statesB = new Map([
    [1, { presence: serializePresenceState(makeState("local")) }],
    [3, { presence: serializePresenceState(makeState("u1")) }],
    [2, { presence: serializePresenceState(makeState("u2")) }],
    [4, { presence: serializePresenceState(makeState("u1")) }],
  ]);

  let flip = false;
  const awareness = {
    clientID: 1,
    setLocalStateField() {},
    getStates() {
      flip = !flip;
      return flip ? statesA : statesB;
    },
  };

  const presence = new PresenceManager(awareness, {
    user: { id: "local", name: "Local", color: "#123456" },
    activeSheet: "Sheet1",
    now,
  });

  const first = presence.getRemotePresences();
  const second = presence.getRemotePresences();

  const toKey = (list) => list.map(({ id, clientId }) => ({ id, clientId }));

  assert.deepEqual(toKey(first), [
    { id: "u1", clientId: 3 },
    { id: "u1", clientId: 4 },
    { id: "u2", clientId: 2 },
  ]);
  assert.deepEqual(toKey(second), toKey(first));
});
