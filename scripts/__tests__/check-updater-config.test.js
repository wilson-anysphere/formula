import assert from "node:assert/strict";
import test from "node:test";

import { checkUpdaterConfig } from "../check-updater-config.mjs";

// Minimal value that satisfies `looksLikeMinisignPublicKey()` in `check-updater-config.mjs`.
// This is base64 for 42 bytes: "Ed" + 40 null bytes.
const VALID_MINISIGN_PUBKEY = "RWQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";

function formatBlocks(blocks) {
  return blocks.map((b) => `${b.heading}\n${b.details.join("\n")}`).join("\n\n");
}

test("check-updater-config: https endpoints are accepted", () => {
  const config = {
    plugins: {
      updater: {
        active: true,
        pubkey: VALID_MINISIGN_PUBKEY,
        endpoints: ["https://github.com/wilson-anysphere/formula/releases/latest/download/latest.json"],
      },
    },
  };

  const result = checkUpdaterConfig(config);
  assert.equal(result.skipped, false);
  assert.equal(result.ok, true, formatBlocks(result.errorBlocks));
});

test("check-updater-config: http endpoints are rejected", () => {
  const config = {
    plugins: {
      updater: {
        active: true,
        pubkey: VALID_MINISIGN_PUBKEY,
        endpoints: ["http://updates.example.org/latest.json"],
      },
    },
  };

  const result = checkUpdaterConfig(config);
  assert.equal(result.ok, false);
  const text = formatBlocks(result.errorBlocks);
  assert.match(text, /endpoints\[0\]/);
  assert.match(text, /http:\/\//);
  assert.match(text, /https:\/\//);
  assert.match(text, /apps\/desktop\/src-tauri\/tauri\.conf\.json/);
  assert.match(text, /plugins\.updater\.endpoints/);
});

test("check-updater-config: relative endpoints are rejected", () => {
  const config = {
    plugins: {
      updater: {
        active: true,
        pubkey: VALID_MINISIGN_PUBKEY,
        endpoints: ["latest.json"],
      },
    },
  };

  const result = checkUpdaterConfig(config);
  assert.equal(result.ok, false);
  const text = formatBlocks(result.errorBlocks);
  assert.match(text, /endpoints\[0\]/);
  assert.match(text, /latest\.json/);
  assert.match(text, /absolute URL/i);
  assert.match(text, /https:\/\//);
});

test("check-updater-config: inactive updater skips endpoint validation", () => {
  const config = {
    plugins: {
      updater: {
        active: false,
        pubkey: "",
        endpoints: ["http://updates.example.org/latest.json"],
      },
    },
  };

  const result = checkUpdaterConfig(config);
  assert.equal(result.skipped, true);
  assert.equal(result.ok, true);
  assert.equal(result.errorBlocks.length, 0);
});
