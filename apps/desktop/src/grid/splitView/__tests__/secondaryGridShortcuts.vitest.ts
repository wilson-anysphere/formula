/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { handleSecondaryGridKeyDown } from "../secondaryGridShortcuts";

function createAppSpies() {
  return {
    copy: vi.fn(),
    cut: vi.fn(),
    paste: vi.fn(),
    clearSelection: vi.fn(),
    openCommentsPanel: vi.fn(),
    fillDown: vi.fn(),
    fillRight: vi.fn(),
    insertDate: vi.fn(),
    insertTime: vi.fn(),
    autoSum: vi.fn(),
  };
}

describe("handleSecondaryGridKeyDown (split-view secondary pane)", () => {
  it("maps Ctrl+D to Fill Down", () => {
    const app = createAppSpies();
    const event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: "d", ctrlKey: true });

    const handled = handleSecondaryGridKeyDown(event, {
      app,
      isSpreadsheetEditing: () => false,
      isTextInputTarget: () => false,
    });

    expect(handled).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    expect(app.fillDown).toHaveBeenCalledTimes(1);
  });

  it("maps Ctrl+R to Fill Right", () => {
    const app = createAppSpies();
    const event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: "r", ctrlKey: true });

    const handled = handleSecondaryGridKeyDown(event, {
      app,
      isSpreadsheetEditing: () => false,
      isTextInputTarget: () => false,
    });

    expect(handled).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    expect(app.fillRight).toHaveBeenCalledTimes(1);
  });

  it("maps Ctrl+; to Insert Date and Ctrl+Shift+; to Insert Time", () => {
    const app = createAppSpies();

    const insertDateEvent = new KeyboardEvent("keydown", {
      bubbles: true,
      cancelable: true,
      key: ";",
      code: "Semicolon",
      ctrlKey: true,
    });
    const handledDate = handleSecondaryGridKeyDown(insertDateEvent, {
      app,
      isSpreadsheetEditing: () => false,
      isTextInputTarget: () => false,
    });

    expect(handledDate).toBe(true);
    expect(insertDateEvent.defaultPrevented).toBe(true);
    expect(app.insertDate).toHaveBeenCalledTimes(1);
    expect(app.insertTime).toHaveBeenCalledTimes(0);

    const insertTimeEvent = new KeyboardEvent("keydown", {
      bubbles: true,
      cancelable: true,
      key: ":",
      code: "Semicolon",
      ctrlKey: true,
      shiftKey: true,
    });
    const handledTime = handleSecondaryGridKeyDown(insertTimeEvent, {
      app,
      isSpreadsheetEditing: () => false,
      isTextInputTarget: () => false,
    });

    expect(handledTime).toBe(true);
    expect(insertTimeEvent.defaultPrevented).toBe(true);
    expect(app.insertTime).toHaveBeenCalledTimes(1);
  });

  it("maps Alt+= to AutoSum", () => {
    const app = createAppSpies();
    const event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: "=", code: "Equal", altKey: true });

    const handled = handleSecondaryGridKeyDown(event, {
      app,
      isSpreadsheetEditing: () => false,
      isTextInputTarget: () => false,
    });

    expect(handled).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    expect(app.autoSum).toHaveBeenCalledTimes(1);
  });

  it("does not steal shortcuts while editing", () => {
    const app = createAppSpies();
    const event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: "d", ctrlKey: true });

    const handled = handleSecondaryGridKeyDown(event, {
      app,
      isSpreadsheetEditing: () => true,
      isTextInputTarget: () => false,
    });

    expect(handled).toBe(false);
    expect(event.defaultPrevented).toBe(false);
    expect(app.fillDown).toHaveBeenCalledTimes(0);
  });
});

