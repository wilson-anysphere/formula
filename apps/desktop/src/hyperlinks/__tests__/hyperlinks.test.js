import test from "node:test";
import assert from "node:assert/strict";

import { handleHyperlinkClick } from "../activate.js";
import { isHyperlinkActivation } from "../activation.js";

test("isHyperlinkActivation matches Ctrl/Cmd+click", () => {
  assert.equal(isHyperlinkActivation({ button: 0, ctrlKey: true, metaKey: false }), true);
  assert.equal(isHyperlinkActivation({ button: 0, ctrlKey: false, metaKey: true }), true);
  assert.equal(isHyperlinkActivation({ button: 0, ctrlKey: false, metaKey: false }), false);
  assert.equal(isHyperlinkActivation({ button: 2, ctrlKey: true, metaKey: false }), false);
});

test("internal hyperlink navigates to sheet + cell on Ctrl/Cmd+click", async () => {
  /** @type {string[]} */
  const calls = [];
  const navigator = {
    async activateSheet(name) {
      calls.push(`sheet:${name}`);
    },
    async selectCell(a1) {
      calls.push(`cell:${a1}`);
    },
  };

  const hyperlink = {
    range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
    target: { type: "internal", sheet: "Sheet2", cell: { row: 1, col: 1 } }, // B2
  };

  const activated = await handleHyperlinkClick(hyperlink, { button: 0, ctrlKey: true }, { navigator });
  assert.equal(activated, true);
  assert.deepEqual(calls, ["sheet:Sheet2", "cell:B2"]);
});

test("external hyperlink opens via shell.open with allowlisted protocol", async () => {
  /** @type {string[]} */
  const calls = [];
  const deps = {
    async shellOpen(uri) {
      calls.push(`open:${uri}`);
    },
    permissions: {
      async request(permission, context) {
        calls.push(`perm:${permission}:${context.protocol}`);
        return true;
      },
    },
    async confirmUntrustedProtocol() {
      calls.push("confirm");
      return true;
    },
  };

  const hyperlink = {
    range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
    target: { type: "external_url", uri: "https://example.com" },
  };

  const activated = await handleHyperlinkClick(hyperlink, { button: 0, metaKey: true }, deps);
  assert.equal(activated, true);
  assert.deepEqual(calls, ["perm:external_navigation:https", "open:https://example.com"]);
});

test("untrusted protocol prompts + requests untrusted permission", async () => {
  /** @type {string[]} */
  const calls = [];
  const deps = {
    async shellOpen(uri) {
      calls.push(`open:${uri}`);
    },
    permissions: {
      async request(permission, context) {
        calls.push(`perm:${permission}:${context.protocol}`);
        return true;
      },
    },
    async confirmUntrustedProtocol() {
      calls.push("confirm");
      return true;
    },
  };

  const hyperlink = {
    range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
    target: { type: "external_url", uri: "ftp://example.com" },
  };

  const activated = await handleHyperlinkClick(hyperlink, { button: 0, ctrlKey: true }, deps);
  assert.equal(activated, true);
  assert.deepEqual(calls, [
    "confirm",
    "perm:external_navigation:ftp",
    "perm:external_navigation_untrusted_protocol:ftp",
    "open:ftp://example.com",
  ]);
});

test("blocked protocols never open", async () => {
  /** @type {string[]} */
  const calls = [];
  const deps = {
    async shellOpen(uri) {
      calls.push(`open:${uri}`);
    },
    async confirmUntrustedProtocol() {
      calls.push("confirm");
      return true;
    },
  };

  const hyperlink = {
    range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
    target: { type: "external_url", uri: "javascript:alert(1)" },
  };

  const activated = await handleHyperlinkClick(hyperlink, { button: 0, ctrlKey: true }, deps);
  assert.equal(activated, false);
  assert.deepEqual(calls, []);
});

