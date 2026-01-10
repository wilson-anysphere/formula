import assert from "node:assert/strict";
import test from "node:test";

import { PresenceManager } from "../index.js";

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

test("PresenceManager throttles cursor updates (~100ms)", () => {
  const scheduler = new FakeScheduler();
  const calls = [];

  const awareness = {
    clientID: 1,
    setLocalStateField(field, value) {
      calls.push({ field, value, at: scheduler.now() });
    },
    getStates() {
      return new Map();
    },
  };

  const presence = new PresenceManager(awareness, {
    user: { id: "u1", name: "Ada", color: "#123456" },
    activeSheet: "Sheet1",
    throttleMs: 100,
    now: () => scheduler.now(),
    setTimeout: scheduler.setTimeout.bind(scheduler),
    clearTimeout: scheduler.clearTimeout.bind(scheduler),
  });

  calls.length = 0;

  presence.setCursor({ row: 1, col: 1 });
  presence.setCursor({ row: 1, col: 2 });
  presence.setCursor({ row: 1, col: 3 });

  assert.equal(calls.length, 1);
  assert.equal(calls[0].value.cursor.col, 1);

  scheduler.advance(50);
  presence.setCursor({ row: 1, col: 4 });
  assert.equal(calls.length, 1);

  scheduler.advance(50);
  assert.equal(calls.length, 2);
  assert.equal(calls[1].value.cursor.col, 4);
});

