// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => {
  return {
    pickLocalImageFiles: vi.fn<[], Promise<File[]>>(),
  };
});

vi.mock("./drawings/pickLocalImageFiles.js", () => ({
  pickLocalImageFiles: mocks.pickLocalImageFiles,
}));

describe("Ribbon Insert → Pictures commands", () => {
  afterEach(() => {
    mocks.pickLocalImageFiles.mockReset();
    vi.restoreAllMocks();
    document.body.innerHTML = "";
    delete (globalThis as any).__formulaSpreadsheetIsEditing;
  });

  it("inserts pictures from This Device via SpreadsheetApp.insertPicturesFromFiles", async () => {
    document.body.innerHTML = `<div id="toast-root"></div>`;

    const file = new File(["fake"], "cat.png", { type: "image/png" });
    mocks.pickLocalImageFiles.mockResolvedValue([file]);

    const insertPicturesFromFiles = vi.fn<Parameters<any>, Promise<void>>().mockResolvedValue(undefined);
    const app = {
      insertPicturesFromFiles,
      focus: vi.fn(),
    };

    const { handleInsertPicturesRibbonCommand } = await import("./main.insertPicturesRibbonCommand");
    await handleInsertPicturesRibbonCommand("insert.illustrations.pictures.thisDevice", app as any);

    expect(insertPicturesFromFiles).toHaveBeenCalledTimes(1);
    expect(insertPicturesFromFiles).toHaveBeenCalledWith([file]);
  });

  it("blocks inserting pictures while the spreadsheet is editing (split-view secondary editor via global flag)", async () => {
    document.body.innerHTML = `<div id="toast-root"></div>`;

    const file = new File(["fake"], "cat.png", { type: "image/png" });
    mocks.pickLocalImageFiles.mockResolvedValue([file]);

    const insertPicturesFromFiles = vi.fn<Parameters<any>, Promise<void>>().mockResolvedValue(undefined);
    const app = {
      isEditing: () => false,
      insertPicturesFromFiles,
      focus: vi.fn(),
    };

    (globalThis as any).__formulaSpreadsheetIsEditing = true;
    const { handleInsertPicturesRibbonCommand } = await import("./main.insertPicturesRibbonCommand");
    const handled = await handleInsertPicturesRibbonCommand("insert.illustrations.pictures.thisDevice", app as any);

    expect(handled).toBe(true);
    expect(mocks.pickLocalImageFiles).not.toHaveBeenCalled();
    expect(insertPicturesFromFiles).not.toHaveBeenCalled();
  });

  it("blocks inserting pictures in read-only mode", async () => {
    document.body.innerHTML = `<div id="toast-root"></div>`;

    const file = new File(["fake"], "cat.png", { type: "image/png" });
    mocks.pickLocalImageFiles.mockResolvedValue([file]);

    const insertPicturesFromFiles = vi.fn<Parameters<any>, Promise<void>>().mockResolvedValue(undefined);
    const focus = vi.fn();
    const app = {
      isReadOnly: () => true,
      insertPicturesFromFiles,
      focus,
    };

    const { handleInsertPicturesRibbonCommand } = await import("./main.insertPicturesRibbonCommand");
    await handleInsertPicturesRibbonCommand("insert.illustrations.pictures.thisDevice", app as any);

    expect(mocks.pickLocalImageFiles).not.toHaveBeenCalled();
    expect(insertPicturesFromFiles).not.toHaveBeenCalled();
    expect(focus).toHaveBeenCalledTimes(1);

    const toast = document.querySelector<HTMLElement>('[data-testid="toast"]');
    expect(toast?.textContent).toBe("Read-only: you don't have permission to insert pictures.");
  });
});
