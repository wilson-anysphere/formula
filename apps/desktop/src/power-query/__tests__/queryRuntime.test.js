import assert from "node:assert/strict";
import test from "node:test";

import { deriveQueryListRows, reduceQueryRuntimeState } from "../queryRuntime.ts";

test("reduceQueryRuntimeState tracks refresh/apply lifecycle", () => {
  let state = {};

  state = reduceQueryRuntimeState(state, { type: "queued", job: { id: "refresh_1", queryId: "q1" } });
  assert.equal(state.q1.status, "queued");
  assert.equal(state.q1.jobId, "refresh_1");

  state = reduceQueryRuntimeState(state, { type: "started", job: { id: "refresh_1", queryId: "q1" } });
  assert.equal(state.q1.status, "refreshing");

  state = reduceQueryRuntimeState(state, {
    type: "completed",
    job: { id: "refresh_1", queryId: "q1", completedAt: new Date(1234) },
    result: { meta: { refreshedAt: new Date(1000) } },
  });
  assert.equal(state.q1.status, "success");
  assert.equal(state.q1.lastRefreshAtMs, 1000);
  assert.equal(state.q1.lastError, null);

  state = reduceQueryRuntimeState(state, { type: "apply:started", jobId: "refresh_1", queryId: "q1" });
  assert.equal(state.q1.status, "applying");
  assert.equal(state.q1.rowsWritten, 0);

  state = reduceQueryRuntimeState(state, { type: "apply:progress", jobId: "refresh_1", queryId: "q1", rowsWritten: 42 });
  assert.equal(state.q1.status, "applying");
  assert.equal(state.q1.rowsWritten, 42);

  state = reduceQueryRuntimeState(state, { type: "apply:completed", jobId: "refresh_1", queryId: "q1", result: { rows: 42, cols: 2 } });
  assert.equal(state.q1.status, "success");
  assert.equal(state.q1.lastError, null);
  assert.equal(state.q1.rowsWritten, undefined);
});

test("deriveQueryListRows uses runtime + persisted metadata", () => {
  const queries = [
    {
      id: "q1",
      name: "Query 1",
      source: { type: "range", range: { values: [["A"], [1]], hasHeaders: true } },
      steps: [],
      destination: { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true, lastOutputSize: { rows: 2, cols: 1 } },
      refreshPolicy: { type: "manual" },
    },
    {
      id: "q2",
      name: "Query 2",
      source: { type: "api", url: "https://example.com/data", method: "GET", auth: { type: "oauth2", providerId: "example" } },
      steps: [],
      refreshPolicy: { type: "manual" },
    },
  ];

  const runtime = {
    q1: { status: "idle" },
    q2: { status: "error", lastError: "Permission denied" },
  };

  const rows = deriveQueryListRows(queries, runtime, { q1: 50, q2: 60 });
  assert.equal(rows.length, 2);

  const row1 = rows.find((r) => r.id === "q1");
  assert.equal(row1.destination, "Sheet1!A1:A2");
  assert.equal(row1.lastRefreshAtMs, 50);
  assert.equal(row1.errorSummary, null);

  const row2 = rows.find((r) => r.id === "q2");
  assert.equal(row2.authRequired, true);
  assert.equal(row2.authLabel, "OAuth2: example");
  assert.equal(row2.lastRefreshAtMs, 60);
  assert.equal(row2.errorSummary, "Permission denied");
});
