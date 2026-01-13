import test from "node:test";
import assert from "node:assert/strict";

import { readRibbonSchemaSource } from "./ribbonSchemaSource.js";

function collectSchemaButtonsWithSize(source, sizeValue) {
  const buttons = new Map();
  const re = new RegExp(`\\bsize\\s*:\\s*["']${sizeValue}["']`, "g");
  let match;
  while ((match = re.exec(source))) {
    const objectStart = findEnclosingObjectStart(source, match.index);
    const objectEnd = findEnclosingObjectEnd(source, objectStart);
    const objectSource = source.slice(objectStart, objectEnd + 1);
    const id = readTopLevelStringProp(objectSource, "id");
    const size = readTopLevelStringProp(objectSource, "size");
    assert.equal(size, sizeValue, `Expected size '${sizeValue}' for button ${id}`);
    const iconId = readOptionalTopLevelStringProp(objectSource, "iconId");
    buttons.set(id, { id, iconId });
  }
  return [...buttons.values()].sort((a, b) => a.id.localeCompare(b.id));
}

function findEnclosingObjectStart(source, fromIndex) {
  let depth = 0;
  let inString = false;
  let quote = "";

  for (let i = fromIndex; i >= 0; i--) {
    const ch = source[i];
    if (inString) {
      if (ch === quote && source[i - 1] !== "\\") {
        inString = false;
        quote = "";
      }
      continue;
    }

    if (ch === '"' || ch === "'") {
      inString = true;
      quote = ch;
      continue;
    }

    if (ch === "}") {
      depth++;
      continue;
    }

    if (ch === "{") {
      if (depth === 0) return i;
      depth--;
    }
  }

  throw new Error(`Could not find enclosing object start for index ${fromIndex}`);
}

function findEnclosingObjectEnd(source, objectStartIndex) {
  let depth = 0;
  let inString = false;
  let quote = "";

  for (let i = objectStartIndex; i < source.length; i++) {
    const ch = source[i];
    if (inString) {
      if (ch === quote && source[i - 1] !== "\\") {
        inString = false;
        quote = "";
      }
      continue;
    }

    if (ch === '"' || ch === "'") {
      inString = true;
      quote = ch;
      continue;
    }

    if (ch === "{") {
      depth++;
      continue;
    }

    if (ch === "}") {
      depth--;
      if (depth === 0) return i;
    }
  }

  throw new Error(`Could not find enclosing object end for start index ${objectStartIndex}`);
}

function readTopLevelStringProp(objectSource, propName) {
  let depth = 0;
  let inString = false;
  let quote = "";

  for (let i = 0; i < objectSource.length; i++) {
    const ch = objectSource[i];
    if (inString) {
      if (ch === quote && objectSource[i - 1] !== "\\") {
        inString = false;
        quote = "";
      }
      continue;
    }

    if (ch === '"' || ch === "'") {
      inString = true;
      quote = ch;
      continue;
    }

    if (ch === "{") {
      depth++;
      continue;
    }

    if (ch === "}") {
      depth--;
      continue;
    }

    if (depth !== 1) continue;
    if (!isIdentifierAt(objectSource, i, propName)) continue;

    let cursor = i + propName.length;
    cursor = skipWhitespace(objectSource, cursor);
    if (objectSource[cursor] !== ":") continue;
    cursor++;
    cursor = skipWhitespace(objectSource, cursor);

    const valueQuote = objectSource[cursor];
    if (valueQuote !== '"' && valueQuote !== "'") continue;
    cursor++;

    let value = "";
    for (; cursor < objectSource.length; cursor++) {
      const c = objectSource[cursor];
      if (c === valueQuote && objectSource[cursor - 1] !== "\\") break;
      value += c;
    }
    return value;
  }

  throw new Error(`Missing top-level '${propName}' string property in object:\n${objectSource.slice(0, 200)}…`);
}

function readOptionalTopLevelStringProp(objectSource, propName) {
  try {
    return readTopLevelStringProp(objectSource, propName);
  } catch {
    return undefined;
  }
}

function isIdentifierAt(source, index, ident) {
  if (!source.startsWith(ident, index)) return false;
  const before = source[index - 1];
  const after = source[index + ident.length];
  if (before && /[a-zA-Z0-9_$]/.test(before)) return false;
  if (after && /[a-zA-Z0-9_$]/.test(after)) return false;
  return true;
}

function skipWhitespace(source, index) {
  let i = index;
  while (i < source.length && /\s/.test(source[i])) i++;
  return i;
}

test('ribbon schema assigns an iconId for every button with size: "icon"', () => {
  const schemaSource = readRibbonSchemaSource();

  const iconButtons = collectSchemaButtonsWithSize(schemaSource, "icon");
  assert.ok(
    iconButtons.some((button) => button.id === "format.toggleBold"),
    "Sanity-check: expected an icon-sized button id",
  );

  const missing = iconButtons.filter((button) => !button.iconId).map((button) => button.id);

  assert.deepEqual(
    missing,
    [],
    `Missing iconId entries for schema size:\"icon\" buttons:\n${missing.map((id) => `- ${id}`).join("\n")}`,
  );
});

test('ribbon schema assigns an iconId for every button with size: "large"', () => {
  const schemaSource = readRibbonSchemaSource();

  const largeButtons = collectSchemaButtonsWithSize(schemaSource, "large");
  assert.ok(
    largeButtons.some((button) => button.id === "file.save.save"),
    "Sanity-check: expected a large-sized button id",
  );

  const missing = largeButtons.filter((button) => !button.iconId).map((button) => button.id);

  assert.deepEqual(
    missing,
    [],
    `Missing iconId entries for schema size:\"large\" buttons:\n${missing.map((id) => `- ${id}`).join("\n")}`,
  );
});
