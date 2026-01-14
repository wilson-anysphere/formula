/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { tagDrawingContextClickPointerDown } from "../tagDrawingContextClick";

describe("tagDrawingContextClickPointerDown", () => {
  it("tags mouse right-click pointerdowns that hit a drawing", () => {
    const canvas = document.createElement("canvas");
    const app = { hitTestDrawingAtClientPoint: vi.fn(() => ({ id: 1 })) };
    const event = {
      pointerType: "mouse",
      button: 2,
      clientX: 10,
      clientY: 20,
      ctrlKey: false,
      metaKey: false,
      target: canvas,
    } as any as PointerEvent;

    expect(tagDrawingContextClickPointerDown(event, app, { isMacPlatform: false, requireCanvasTarget: true })).toBe(true);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((event as any).__formulaDrawingContextClick).toBe(true);
    expect(app.hitTestDrawingAtClientPoint).toHaveBeenCalledWith(10, 20);
  });

  it("does not tag when the pointerdown does not hit a drawing", () => {
    const canvas = document.createElement("canvas");
    const app = { hitTestDrawingAtClientPoint: vi.fn(() => null) };
    const event = {
      pointerType: "mouse",
      button: 2,
      clientX: 10,
      clientY: 20,
      ctrlKey: false,
      metaKey: false,
      target: canvas,
    } as any as PointerEvent;

    expect(tagDrawingContextClickPointerDown(event, app, { isMacPlatform: false, requireCanvasTarget: true })).toBe(false);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((event as any).__formulaDrawingContextClick).toBeUndefined();
  });

  it("does not tag non-context-click pointerdowns", () => {
    const canvas = document.createElement("canvas");
    const app = { hitTestDrawingAtClientPoint: vi.fn(() => ({ id: 1 })) };
    const event = {
      pointerType: "mouse",
      button: 0,
      clientX: 10,
      clientY: 20,
      ctrlKey: false,
      metaKey: false,
      target: canvas,
    } as any as PointerEvent;

    expect(tagDrawingContextClickPointerDown(event, app, { isMacPlatform: false, requireCanvasTarget: true })).toBe(false);
    expect(app.hitTestDrawingAtClientPoint).not.toHaveBeenCalled();
  });

  it("treats macOS Ctrl+click as a context click", () => {
    const canvas = document.createElement("canvas");
    const app = { hitTestDrawingAtClientPoint: vi.fn(() => ({ id: 1 })) };
    const event = {
      pointerType: "mouse",
      button: 0,
      clientX: 10,
      clientY: 20,
      ctrlKey: true,
      metaKey: false,
      target: canvas,
    } as any as PointerEvent;

    expect(tagDrawingContextClickPointerDown(event, app, { isMacPlatform: true, requireCanvasTarget: true })).toBe(true);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((event as any).__formulaDrawingContextClick).toBe(true);
  });
});

