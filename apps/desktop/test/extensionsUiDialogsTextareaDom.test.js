import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

let showInputBox = null;
try {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const mainSource = stripComments(fs.readFileSync(mainPath, "utf8"));
  const uiImportMatch = mainSource.match(
    /import\s+\{[^}]*\bshowInputBox\b[^}]*\}\s+from\s+["'](\.\/extensions\/ui[^"']*)["'];/,
  );

  if (uiImportMatch?.[1]) {
    const candidatePath = path.join(__dirname, "..", "src", uiImportMatch[1]);
    if (fs.existsSync(candidatePath)) {
      const mod = await import(pathToFileURL(candidatePath).href);
      if (typeof mod.showInputBox === "function") {
        showInputBox = mod.showInputBox;
      }
    }
  }
} catch {
  // ignore; fall back to probing known filenames below.
}

if (!showInputBox) {
  try {
    const extensionsDir = path.join(__dirname, "..", "src", "extensions");
    const candidates = ["ui.js", "ui.ts"];
    for (const candidate of candidates) {
      const candidatePath = path.join(extensionsDir, candidate);
      if (!fs.existsSync(candidatePath)) continue;
      // eslint-disable-next-line no-await-in-loop
      const mod = await import(new URL(`../src/extensions/${candidate}`, import.meta.url).href);
      if (typeof mod.showInputBox === "function") {
        showInputBox = mod.showInputBox;
        break;
      }
    }
  } catch {
    // ignore; test will be skipped below if showInputBox is unavailable.
  }
}

let JSDOM = null;
try {
  // `jsdom` is optional for lightweight node:test runs (some agent environments do not
  // install workspace dev dependencies). Skip DOM-specific tests when unavailable.
  // eslint-disable-next-line node/no-unsupported-features/es-syntax
  ({ JSDOM } = await import("jsdom"));
} catch {
  // ignore
}

const hasDom = Boolean(JSDOM);

async function withDom(fn) {
  const dom = new JSDOM("<!doctype html><html><body></body></html>", { url: "http://localhost" });

  /** @type {Record<string, any>} */
  const prev = {
    window: globalThis.window,
    document: globalThis.document,
    Event: globalThis.Event,
    KeyboardEvent: globalThis.KeyboardEvent,
    HTMLElement: globalThis.HTMLElement,
    HTMLDialogElement: globalThis.HTMLDialogElement,
    HTMLInputElement: globalThis.HTMLInputElement,
    HTMLTextAreaElement: globalThis.HTMLTextAreaElement,
    HTMLButtonElement: globalThis.HTMLButtonElement,
  };

  globalThis.window = dom.window;
  globalThis.document = dom.window.document;
  globalThis.Event = dom.window.Event;
  globalThis.KeyboardEvent = dom.window.KeyboardEvent;
  globalThis.HTMLElement = dom.window.HTMLElement;
  globalThis.HTMLDialogElement = dom.window.HTMLDialogElement;
  globalThis.HTMLInputElement = dom.window.HTMLInputElement;
  globalThis.HTMLTextAreaElement = dom.window.HTMLTextAreaElement;
  globalThis.HTMLButtonElement = dom.window.HTMLButtonElement;

  try {
    await fn(dom);
  } finally {
    for (const [key, value] of Object.entries(prev)) {
      if (value === undefined) {
        delete globalThis[key];
      } else {
        globalThis[key] = value;
      }
    }
  }
}

test("showInputBox textarea mode renders without inline styles", { skip: !hasDom || !showInputBox }, async () => {
  await withDom(async () => {
    const promise = showInputBox({ prompt: "JSON", value: "{}", type: "textarea" });
    const dialog = document.querySelector('dialog[data-testid="input-box"]');
    assert.ok(dialog, "Expected showInputBox to insert a dialog into the DOM");

    const field = dialog.querySelector('[data-testid="input-box-field"]');
    assert.ok(field, "Expected showInputBox to insert an input field");
    assert.equal(field.tagName, "TEXTAREA");
    assert.ok(field.classList.contains("dialog__field"));

    // Guard against accidental return of inline styling (regression in textarea mode).
    assert.equal(field.getAttribute("style") ?? "", "");
    // @ts-expect-error - field is an Element in JS tests.
    assert.equal(field.style.cssText, "");

    const cancel = dialog.querySelector('[data-testid="input-box-cancel"]');
    assert.ok(cancel);
    cancel.click();

    const result = await promise;
    assert.equal(result, null);
    assert.equal(document.querySelector('dialog[data-testid="input-box"]'), null);
  });
});
