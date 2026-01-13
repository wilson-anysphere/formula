// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CanvasGrid } from "../CanvasGrid";
import type { CellProvider } from "../../model/CellProvider";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

function createPointerEvent(type: string, options: { clientX: number; clientY: number; pointerId: number }): Event {
  const PointerEventCtor = (window as unknown as { PointerEvent?: typeof PointerEvent }).PointerEvent;
  if (PointerEventCtor) {
    return new PointerEventCtor(type, {
      bubbles: true,
      cancelable: true,
      clientX: options.clientX,
      clientY: options.clientY,
      pointerId: options.pointerId,
      pointerType: "mouse"
    } as PointerEventInit);
  }

  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    clientX: options.clientX,
    clientY: options.clientY
  });
  Object.defineProperty(event, "pointerId", { value: options.pointerId });
  Object.defineProperty(event, "pointerType", { value: "mouse" });
  return event;
}

describe("CanvasGrid resize hit testing", () => {
  beforeEach(() => {
    vi.stubGlobal(
      "ResizeObserver",
      class ResizeObserver {
        observe(): void {}
        unobserve(): void {}
        disconnect(): void {}
      }
    );

    // Avoid running full render frames; these tests only validate interaction wiring.
    vi.stubGlobal("requestAnimationFrame", vi.fn(() => 0));

    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue({
      left: 0,
      top: 0,
      right: 300,
      bottom: 200,
      width: 300,
      height: 200,
      x: 0,
      y: 0,
      toJSON: () => ({})
    } as unknown as DOMRect);

    const ctxStub: Partial<CanvasRenderingContext2D> = {
      setTransform: vi.fn(),
      measureText: (text: string) =>
        ({
          width: text.length * 6,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2
        }) as TextMetrics
    };

    vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockImplementation(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => ctxStub as any
    );
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("only enables resize cursors within headerRows/headerCols (not the full frozen region)", async () => {
    const provider: CellProvider = { getCell: () => null };

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={provider}
          rowCount={100}
          colCount={100}
          headerRows={1}
          headerCols={1}
          frozenRows={3}
          frozenCols={3}
          defaultRowHeight={20}
          defaultColWidth={50}
          enableResize
        />
      );
    });

    const selectionCanvas = host.querySelector('[data-testid="canvas-grid-selection"]') as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    // Control: inside the true column header row (y < headerHeight) near a column boundary.
    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { clientX: 100, clientY: 10, pointerId: 1 }));
    });
    expect(selectionCanvas.style.cursor).toBe("col-resize");

    // Frozen but below the header row should not show a column resize cursor.
    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { clientX: 100, clientY: 30, pointerId: 1 }));
    });
    expect(selectionCanvas.style.cursor).toBe("default");

    // Control: inside the true row header column (x < headerWidth) near a row boundary.
    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { clientX: 10, clientY: 40, pointerId: 1 }));
    });
    expect(selectionCanvas.style.cursor).toBe("row-resize");

    // Frozen but to the right of the row header column should not show a row resize cursor.
    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { clientX: 60, clientY: 40, pointerId: 1 }));
    });
    expect(selectionCanvas.style.cursor).toBe("default");

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("updates resize hit testing when headers become controlled", async () => {
    const provider: CellProvider = { getCell: () => null };

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={provider}
          rowCount={100}
          colCount={100}
          frozenRows={3}
          frozenCols={3}
          defaultRowHeight={20}
          defaultColWidth={50}
          enableResize
        />
      );
    });

    const selectionCanvas = host.querySelector('[data-testid="canvas-grid-selection"]') as HTMLCanvasElement;
    expect(selectionCanvas).toBeTruthy();

    // Uncontrolled headers use legacy behavior: treat the first frozen row/col as headers.
    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { clientX: 100, clientY: 30, pointerId: 1 }));
    });
    expect(selectionCanvas.style.cursor).toBe("default");

    // Switch to controlled headers; now the second header row should also be resize-active.
    await act(async () => {
      root.render(
        <CanvasGrid
          provider={provider}
          rowCount={100}
          colCount={100}
          headerRows={2}
          headerCols={1}
          frozenRows={3}
          frozenCols={3}
          defaultRowHeight={20}
          defaultColWidth={50}
          enableResize
        />
      );
    });

    await act(async () => {
      selectionCanvas.dispatchEvent(createPointerEvent("pointermove", { clientX: 100, clientY: 30, pointerId: 1 }));
    });
    expect(selectionCanvas.style.cursor).toBe("col-resize");

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});
