import FUNCTION_CATALOG from "../../../../../shared/functionCatalog.mjs";
import { getFunctionSignature, signatureParts } from "../highlight/functionSignatures.js";

import type { FormulaBarView } from "../FormulaBarView.js";

type CatalogFunction = { name?: string };

type CompletionContext = {
  /**
   * Inclusive start index of the identifier-like token in the input.
   */
  replaceStart: number;
  /**
   * Exclusive end index of the identifier-like token in the input.
   */
  replaceEnd: number;
  /**
   * Token prefix typed before the caret (used for filtering + casing preservation).
   */
  typedPrefix: string;
  /**
   * Optional function qualifier prefix (e.g. `_xlfn.`) that is preserved during insertion.
   */
  qualifier: string;
  /**
   * Uppercased prefix used for matching against the catalog.
   */
  matchPrefixUpper: string;
};

type FunctionSuggestion = { name: string; signature: string };

let AUTOCOMPLETE_INSTANCE_ID = 0;

const FUNCTION_NAMES: string[] = (() => {
  const names = new Set<string>();
  const items = (FUNCTION_CATALOG as { functions?: CatalogFunction[] } | null)?.functions ?? [];
  for (const fn of items) {
    const name = typeof fn?.name === "string" ? fn.name.trim() : "";
    if (!name) continue;
    names.add(name);
  }
  return Array.from(names).sort((a, b) => a.localeCompare(b));
})();

const FUNCTION_NAMES_UPPER = new Set(FUNCTION_NAMES.map((name) => name.toUpperCase()));

const DEFAULT_ARG_SEPARATOR = (() => {
  const locale = (() => {
    try {
      const nav = (globalThis as any).navigator;
      const lang = typeof nav?.language === "string" ? nav.language : "";
      return lang || "en-US";
    } catch {
      return "en-US";
    }
  })();

  try {
    const parts = new Intl.NumberFormat(locale).formatToParts(1.1);
    const decimal = parts.find((p) => p.type === "decimal")?.value ?? ".";
    return decimal === "," ? "; " : ", ";
  } catch {
    return ", ";
  }
})();

function isWhitespace(ch: string): boolean {
  return ch === " " || ch === "\t" || ch === "\n" || ch === "\r";
}

function isIdentifierChar(ch: string): boolean {
  // Match the formula tokenizer's identifier rules closely enough for completion.
  // Excel function names allow dots (e.g. `COVARIANCE.P`) and digits (e.g. `LOG10`).
  return (
    ch === "_" ||
    ch === "." ||
    (ch >= "0" && ch <= "9") ||
    (ch >= "A" && ch <= "Z") ||
    (ch >= "a" && ch <= "z")
  );
}

function clampCursor(input: string, cursorPosition: number): number {
  if (!Number.isInteger(cursorPosition)) return input.length;
  if (cursorPosition < 0) return 0;
  if (cursorPosition > input.length) return input.length;
  return cursorPosition;
}

function firstNonWhitespaceIndex(text: string): number {
  for (let i = 0; i < text.length; i += 1) {
    if (!isWhitespace(text[i]!)) return i;
  }
  return -1;
}

function findCompletionContext(input: string, cursorPosition: number): CompletionContext | null {
  const cursor = clampCursor(input, cursorPosition);

  const firstNonWhitespace = firstNonWhitespaceIndex(input);
  if (firstNonWhitespace < 0) return null;
  if (input[firstNonWhitespace] !== "=") return null;

  // Require a collapsed selection (caller ensures selectionStart === selectionEnd).
  let replaceStart = cursor;
  while (replaceStart > 0 && isIdentifierChar(input[replaceStart - 1]!)) replaceStart -= 1;

  let replaceEnd = cursor;
  while (replaceEnd < input.length && isIdentifierChar(input[replaceEnd]!)) replaceEnd += 1;

  const typedPrefix = input.slice(replaceStart, cursor);
  if (typedPrefix.length < 1) return null;
  // Only trigger on identifier-looking starts.
  // (We handle `_xlfn.` separately below.)
  if (!/^[_A-Za-z]/.test(typedPrefix)) return null;

  // Ensure we're at the start of an expression-like position:
  // `=VLO`, `=1+VLO`, `=SUM(VLO`, `=SUM(A, VLO)`
  let prev = replaceStart - 1;
  while (prev >= 0 && isWhitespace(input[prev]!)) prev -= 1;
  if (prev < 0) return null;

  const prevChar = input[prev]!;
  const startsExpression =
    prevChar === "=" || prevChar === "(" || prevChar === "," || prevChar === ";" || "+-*/^&=><%@".includes(prevChar);
  if (!startsExpression) return null;

  // In argument positions (after `(` or `,`), very short alphabetic identifiers are
  // much more likely to be column/range refs (e.g. `SUM(A` / `SUM(AB`) than function
  // names. Be conservative here so we don't steal Tab from range completion.
  if (prevChar === "(" || prevChar === "," || prevChar === ";") {
    if (/^[A-Za-z]+$/.test(typedPrefix)) {
      if (typedPrefix.length === 1) return null;
      if (typedPrefix.length === 2 && !FUNCTION_NAMES_UPPER.has(typedPrefix.toUpperCase())) return null;
    }
  }

  // Support Excel `_xlfn.` function prefix in pasted formulas.
  const upper = typedPrefix.toUpperCase();
  const qualifierUpper = "_XLFN.";
  if (upper.startsWith(qualifierUpper)) {
    const qualifier = typedPrefix.slice(0, qualifierUpper.length);
    const rest = typedPrefix.slice(qualifierUpper.length);
    if (rest.length < 1) return null;
    return {
      replaceStart,
      replaceEnd,
      typedPrefix,
      qualifier,
      matchPrefixUpper: rest.toUpperCase(),
    };
  }

  return {
    replaceStart,
    replaceEnd,
    typedPrefix,
    qualifier: "",
    matchPrefixUpper: upper,
  };
}

function signaturePreview(name: string): string {
  const sig = getFunctionSignature(name);
  if (!sig) return `${name}(…)`;
  return signatureParts(sig, null, { argSeparator: DEFAULT_ARG_SEPARATOR })
    .map((p) => p.text)
    .join("");
}

function preserveTypedCasing(typedPrefix: string, canonical: string): string {
  if (!typedPrefix) return canonical;
  if (typedPrefix.length >= canonical.length) return typedPrefix;

  // Infer case preference from the *letters* the user typed (ignore digits, dots, underscores).
  // This yields nicer results for common patterns like:
  //   "=vlo"  -> "=vlookup("
  //   "=VLO"  -> "=VLOOKUP("
  //   "=Vlo"  -> "=Vlookup("
  const letters = typedPrefix.replaceAll(/[^A-Za-z]/g, "");
  if (!letters) return typedPrefix + canonical.slice(typedPrefix.length);

  const lower = letters.toLowerCase();
  const upper = letters.toUpperCase();
  if (letters === lower) return canonical.toLowerCase();
  if (letters === upper) return canonical.toUpperCase();

  // Title-ish casing: first letter uppercase, remainder lowercase.
  if (letters[0] === upper[0] && letters.slice(1) === lower.slice(1)) {
    const lowered = canonical.toLowerCase();
    const firstLetterIdx = lowered.search(/[a-z]/);
    if (firstLetterIdx >= 0) {
      return lowered.slice(0, firstLetterIdx) + lowered[firstLetterIdx]!.toUpperCase() + lowered.slice(firstLetterIdx + 1);
    }
    return lowered;
  }

  // Fallback: preserve the exact prefix the user typed and append the canonical tail.
  return typedPrefix + canonical.slice(typedPrefix.length);
}

function buildSuggestions(prefixUpper: string, limit: number): FunctionSuggestion[] {
  const out: FunctionSuggestion[] = [];
  if (!prefixUpper) return out;

  for (const name of FUNCTION_NAMES) {
    if (!name.toUpperCase().startsWith(prefixUpper)) continue;
    out.push({ name, signature: signaturePreview(name) });
    if (out.length >= limit) break;
  }
  return out;
}

interface FormulaBarFunctionAutocompleteControllerOptions {
  formulaBar: FormulaBarView;
  /**
   * Element used as the positioning context (should be `position: relative`).
   */
  anchor: HTMLElement;
  maxItems?: number;
}

export class FormulaBarFunctionAutocompleteController {
  readonly #formulaBar: FormulaBarView;
  readonly #textarea: HTMLTextAreaElement;
  readonly #maxItems: number;

  readonly #dropdownEl: HTMLDivElement;
  readonly #listboxId: string;
  readonly #optionIdPrefix: string;
  #itemEls: HTMLButtonElement[] = [];

  #context: CompletionContext | null = null;
  #suggestions: FunctionSuggestion[] = [];
  #selectedIndex = 0;
  #activeDescendantId: string | null = null;

  readonly #unsubscribe: Array<() => void> = [];

  constructor(opts: FormulaBarFunctionAutocompleteControllerOptions) {
    this.#formulaBar = opts.formulaBar;
    this.#textarea = opts.formulaBar.textarea;
    this.#maxItems = Math.max(1, Math.min(50, opts.maxItems ?? 12));

    AUTOCOMPLETE_INSTANCE_ID += 1;
    this.#listboxId = `formula-function-autocomplete-${AUTOCOMPLETE_INSTANCE_ID}`;
    this.#optionIdPrefix = `${this.#listboxId}-option`;

    const dropdown = document.createElement("div");
    dropdown.className = "formula-bar-function-autocomplete";
    dropdown.dataset.testid = "formula-function-autocomplete";
    dropdown.setAttribute("role", "listbox");
    dropdown.id = this.#listboxId;
    dropdown.hidden = true;
    opts.anchor.appendChild(dropdown);
    this.#dropdownEl = dropdown;

    // Keep the textarea focused while navigating the listbox, using the
    // active-descendant pattern for screen readers.
    this.#textarea.setAttribute("aria-haspopup", "listbox");
    this.#textarea.setAttribute("aria-controls", this.#listboxId);
    this.#textarea.setAttribute("aria-expanded", "false");

    const updateNow = () => this.update();
    const onBlur = () => this.close();
    this.#textarea.addEventListener("input", updateNow);
    this.#textarea.addEventListener("click", updateNow);
    this.#textarea.addEventListener("keyup", updateNow);
    this.#textarea.addEventListener("focus", updateNow);
    this.#textarea.addEventListener("select", updateNow);
    this.#textarea.addEventListener("blur", onBlur);

    this.#unsubscribe.push(() => this.#textarea.removeEventListener("input", updateNow));
    this.#unsubscribe.push(() => this.#textarea.removeEventListener("click", updateNow));
    this.#unsubscribe.push(() => this.#textarea.removeEventListener("keyup", updateNow));
    this.#unsubscribe.push(() => this.#textarea.removeEventListener("focus", updateNow));
    this.#unsubscribe.push(() => this.#textarea.removeEventListener("select", updateNow));
    this.#unsubscribe.push(() => this.#textarea.removeEventListener("blur", onBlur));
  }

  destroy(): void {
    for (const stop of this.#unsubscribe.splice(0)) stop();
    this.#dropdownEl.remove();
  }

  isOpen(): boolean {
    return !this.#dropdownEl.hidden;
  }

  close(): void {
    if (this.#dropdownEl.hidden) return;
    this.#dropdownEl.hidden = true;
    this.#dropdownEl.textContent = "";
    this.#itemEls = [];
    this.#context = null;
    this.#suggestions = [];
    this.#selectedIndex = 0;
    this.#activeDescendantId = null;
    this.#textarea.removeAttribute("aria-activedescendant");
    this.#textarea.setAttribute("aria-expanded", "false");
  }

  update(): void {
    if (!this.#formulaBar.model.isEditing) {
      this.close();
      return;
    }

    const start = this.#textarea.selectionStart ?? this.#textarea.value.length;
    const end = this.#textarea.selectionEnd ?? this.#textarea.value.length;
    if (start !== end) {
      this.close();
      return;
    }

    const input = this.#textarea.value;
    const ctx = findCompletionContext(input, start);
    if (!ctx) {
      this.close();
      return;
    }

    const suggestions = buildSuggestions(ctx.matchPrefixUpper, this.#maxItems);
    if (suggestions.length === 0) {
      this.close();
      return;
    }

    // Preserve selection when possible.
    const prevSelectedName = this.#suggestions[this.#selectedIndex]?.name ?? null;
    let nextIndex = 0;
    if (prevSelectedName) {
      const found = suggestions.findIndex((s) => s.name === prevSelectedName);
      if (found >= 0) nextIndex = found;
    }

    this.#context = ctx;
    this.#suggestions = suggestions;
    this.#selectedIndex = Math.min(nextIndex, suggestions.length - 1);
    this.#render();
  }

  handleKeyDown(e: KeyboardEvent): boolean {
    if (!this.isOpen()) return false;

    if (e.key === "ArrowDown") {
      e.preventDefault();
      this.#selectedIndex = Math.min(this.#selectedIndex + 1, this.#suggestions.length - 1);
      this.#syncSelection();
      return true;
    }

    if (e.key === "ArrowUp") {
      e.preventDefault();
      this.#selectedIndex = Math.max(this.#selectedIndex - 1, 0);
      this.#syncSelection();
      return true;
    }

    if (e.key === "Escape") {
      e.preventDefault();
      this.close();
      return true;
    }

    // Match typical editor UX:
    // - Tab accepts the selected item (Shift+Tab should keep its usual meaning in the formula bar)
    // - Enter accepts (Shift+Enter remains available for formula-bar commit/navigation semantics)
    if ((e.key === "Tab" && !e.shiftKey) || (e.key === "Enter" && !e.altKey && !e.shiftKey)) {
      e.preventDefault();
      this.acceptSelected();
      return true;
    }

    return false;
  }

  acceptSelected(): void {
    const ctx = this.#context;
    const selected = this.#suggestions[this.#selectedIndex] ?? null;
    if (!ctx || !selected) {
      this.close();
      return;
    }

    const input = this.#textarea.value;
    // Preserve the user-typed casing for the function name portion while keeping
    // any `_xlfn.` qualifier prefix intact (Excel compatibility).
    const typedNamePrefix = ctx.typedPrefix.slice(ctx.qualifier.length);
    const casedName = preserveTypedCasing(typedNamePrefix, selected.name);
    const insertedName = `${ctx.qualifier}${casedName}`;
    // Avoid duplicating the opening paren if the user already has one in the text (e.g. editing `=VLO()`).
    const hasParen = input[ctx.replaceEnd] === "(";
    const inserted = hasParen ? insertedName : `${insertedName}(`;

    const nextText = input.slice(0, ctx.replaceStart) + inserted + input.slice(ctx.replaceEnd);
    // Always place the caret inside the parens (after `(`). When `(` already exists,
    // step over it after inserting the function name.
    const nextCursor = ctx.replaceStart + inserted.length + (hasParen ? 1 : 0);

    this.#textarea.value = nextText;
    this.#textarea.setSelectionRange(nextCursor, nextCursor);
    // Keep editing inside the formula bar.
    try {
      this.#textarea.focus({ preventScroll: true });
    } catch {
      this.#textarea.focus();
    }

    // Notify FormulaBarView listeners (model sync + highlight render).
    this.#textarea.dispatchEvent(new Event("input"));
    this.close();
  }

  #render(): void {
    this.#dropdownEl.hidden = false;
    this.#dropdownEl.textContent = "";
    this.#itemEls = [];
    this.#activeDescendantId = null;
    this.#textarea.setAttribute("aria-expanded", "true");

    for (let i = 0; i < this.#suggestions.length; i += 1) {
      const item = this.#suggestions[i]!;
      const button = document.createElement("button");
      button.type = "button";
      button.className = "formula-bar-function-autocomplete-item";
      button.setAttribute("role", "option");
      button.dataset.testid = "formula-function-autocomplete-item";
      button.dataset.name = item.name;
      button.setAttribute("aria-selected", i === this.#selectedIndex ? "true" : "false");
      const id = `${this.#optionIdPrefix}-${i}`;
      button.id = id;
      if (i === this.#selectedIndex) {
        this.#activeDescendantId = id;
      }

      // Prevent the mousedown from stealing focus from the textarea.
      button.addEventListener("mousedown", (e) => e.preventDefault());
      button.addEventListener("mouseenter", () => {
        this.#selectedIndex = i;
        this.#syncSelection();
      });
      button.addEventListener("click", () => this.acceptSelected());

      const name = document.createElement("div");
      name.className = "formula-bar-function-autocomplete-name";
      name.textContent = item.name;

      const sig = document.createElement("div");
      sig.className = "formula-bar-function-autocomplete-signature";
      sig.textContent = item.signature;

      button.appendChild(name);
      button.appendChild(sig);

      this.#dropdownEl.appendChild(button);
      this.#itemEls.push(button);
    }

    this.#syncSelection();
  }

  #syncSelection(): void {
    this.#activeDescendantId = null;
    for (let i = 0; i < this.#itemEls.length; i += 1) {
      const el = this.#itemEls[i]!;
      const selected = i === this.#selectedIndex;
      el.setAttribute("aria-selected", selected ? "true" : "false");
      if (selected) {
        this.#activeDescendantId = el.id || null;
        try {
          el.scrollIntoView({ block: "nearest" });
        } catch {
          // jsdom doesn't implement layout/scroll; ignore.
        }
      }
    }

    if (this.#activeDescendantId) {
      this.#textarea.setAttribute("aria-activedescendant", this.#activeDescendantId);
    } else {
      this.#textarea.removeAttribute("aria-activedescendant");
    }
  }
}
