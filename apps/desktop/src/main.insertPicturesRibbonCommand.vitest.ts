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
});

