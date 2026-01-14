import { describe, expect, it } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import { WorkbookSheetStore } from "../../../sheets/workbookSheetStore.js";

import { buildImportedDrawingLayerSnapshotAdditions } from "../hydrateImportedDrawings.js";

describe("buildImportedDrawingLayerSnapshotAdditions", () => {
  it("hydrates imported drawing objects + images into the DocumentController snapshot schema", () => {
    const workbookSheetStore = new WorkbookSheetStore([{ id: "sheet-1", name: "Sheet1", visibility: "visible" }]);

    const imported = {
      drawings: [
        {
          sheet_name: "Sheet1",
          sheet_part: "xl/worksheets/sheet1.xml",
          drawing_part: "xl/drawings/drawing1.xml",
          objects: [
            {
              id: 1,
              kind: { Image: { image_id: "image1.png" } },
              anchor: {
                TwoCell: {
                  from: { cell: { row: 0, col: 0 }, offset: { x_emu: 0, y_emu: 0 } },
                  to: { cell: { row: 10, col: 5 }, offset: { x_emu: 0, y_emu: 0 } },
                },
              },
              z_order: 0,
            },
          ],
        },
      ],
      images: [{ id: "image1.png", bytesBase64: "AQID", mimeType: "image/png" }],
    };

    const additions = buildImportedDrawingLayerSnapshotAdditions(imported, workbookSheetStore);
    expect(additions).not.toBeNull();
    expect(additions?.images).toEqual([{ id: "image1.png", bytesBase64: "AQID", mimeType: "image/png" }]);
    expect(Object.keys(additions?.drawingsBySheet ?? {})).toEqual(["sheet-1"]);
    expect(additions?.drawingsBySheet["sheet-1"]?.length).toBe(1);
    expect(additions?.drawingsBySheet["sheet-1"]?.[0]?.id).toBe("1");

    const snapshotState: any = {
      schemaVersion: 1,
      sheets: [{ id: "sheet-1", name: "Sheet1", visibility: "visible", cells: [] }],
      images: additions?.images,
      drawingsBySheet: additions?.drawingsBySheet,
    };

    const encoded = new TextEncoder().encode(JSON.stringify(snapshotState));
    const doc = new DocumentController();
    doc.applyState(encoded);

    expect(doc.getImage("image1.png")).not.toBeNull();
    expect(doc.getSheetDrawings("sheet-1").length).toBeGreaterThan(0);
  });
});

