import { describe, expect, it } from "vitest";

import { DocumentController } from "../document/documentController.js";
import {
  rewriteDeletedSheetReferencesInFormula,
  rewriteDocumentFormulasForSheetDelete,
  rewriteDocumentFormulasForSheetRename,
} from "./sheetFormulaRewrite";

describe("sheetFormulaRewrite", () => {
  describe("rewriteDeletedSheetReferencesInFormula", () => {
    it("rewrites quoted sheet-qualified refs to #REF!", () => {
      expect(
        rewriteDeletedSheetReferencesInFormula("=SUM('My Sheet'!A1,2)", "My Sheet", ["Sheet1", "My Sheet"]),
      ).toBe("=SUM(#REF!,2)");
    });

    it("handles escaped apostrophes in quoted sheet names", () => {
      expect(rewriteDeletedSheetReferencesInFormula("='O''Brien'!A1+1", "O'Brien", ["O'Brien", "Sheet1"])).toBe("=#REF!+1");
    });

    it("rewrites unquoted sheet-qualified refs case-insensitively", () => {
      expect(rewriteDeletedSheetReferencesInFormula("=sheet1!A1+1", "Sheet1", ["Sheet1", "Sheet2"])).toBe("=#REF!+1");
    });

    it("rewrites 3D refs when deleted sheet is start/end", () => {
      const order = ["Sheet1", "Sheet2", "Sheet3"];
      expect(rewriteDeletedSheetReferencesInFormula("=SUM(Sheet1:Sheet3!A1)", "Sheet1", order)).toBe("=SUM(Sheet2:Sheet3!A1)");
      expect(rewriteDeletedSheetReferencesInFormula("=SUM(Sheet1:Sheet3!A1)", "Sheet3", order)).toBe("=SUM(Sheet1:Sheet2!A1)");
    });

    it("does not rewrite inside string literals", () => {
      expect(rewriteDeletedSheetReferencesInFormula('="Sheet1!A1"&Sheet1!A1', "Sheet1", ["Sheet1"])).toBe(
        '="Sheet1!A1"&#REF!',
      );
    });

    it("rewrites sheet-qualified refs using Unicode NFKC matching (e.g. Å == Å)", () => {
      expect(rewriteDeletedSheetReferencesInFormula("='Å'!A1+1", "Å", ["Å"])).toBe("=#REF!+1");
    });
  });

  describe("rewriteDocumentFormulasForSheetRename", () => {
    it("rewrites formulas across all sheets and batches updates", () => {
      const doc = new DocumentController();
      doc.setCellFormula("S1", { row: 0, col: 0 }, "='My Sheet'!A1");
      doc.setCellFormula("S2", { row: 0, col: 0 }, "=SUM(Sheet1:Sheet3!A1)");
      doc.markSaved();

      rewriteDocumentFormulasForSheetRename(doc, "My Sheet", "New Sheet");

      expect(doc.getCell("S1", { row: 0, col: 0 }).formula).toBe("='New Sheet'!A1");
      // Unrelated formula stays intact.
      expect(doc.getCell("S2", { row: 0, col: 0 }).formula).toBe("=SUM(Sheet1:Sheet3!A1)");
      expect(doc.isDirty).toBe(true);
    });

    it("handles quoted sheet names with escaped apostrophes", () => {
      const doc = new DocumentController();
      doc.setCellFormula("S1", { row: 0, col: 0 }, "='O''Brien'!A1");

      rewriteDocumentFormulasForSheetRename(doc, "O'Brien", "O'Brien2");

      expect(doc.getCell("S1", { row: 0, col: 0 }).formula).toBe("='O''Brien2'!A1");
    });

    it("rewrites sheet names in 3D refs", () => {
      const doc = new DocumentController();
      doc.setCellFormula("S1", { row: 0, col: 0 }, "=SUM(Sheet1:Sheet3!A1)");

      rewriteDocumentFormulasForSheetRename(doc, "Sheet3", "End");

      expect(doc.getCell("S1", { row: 0, col: 0 }).formula).toBe("=SUM(Sheet1:End!A1)");
    });

    it("matches old sheet names using Unicode NFKC semantics when rewriting", () => {
      const doc = new DocumentController();
      doc.setCellFormula("S1", { row: 0, col: 0 }, "='Å'!A1");

      // Angstrom sign (U+212B) NFKC-normalizes to Å (U+00C5).
      rewriteDocumentFormulasForSheetRename(doc, "Å", "Budget");

      expect(doc.getCell("S1", { row: 0, col: 0 }).formula).toBe("=Budget!A1");
    });

    it("quotes reserved / ambiguous sheet names in output (e.g. TRUE, A1, R1C1)", () => {
      {
        const doc = new DocumentController();
        doc.setCellFormula("S1", { row: 0, col: 0 }, "=Sheet1!A1");
        rewriteDocumentFormulasForSheetRename(doc, "Sheet1", "TRUE");
        expect(doc.getCell("S1", { row: 0, col: 0 }).formula).toBe("='TRUE'!A1");
      }

      {
        const doc = new DocumentController();
        doc.setCellFormula("S1", { row: 0, col: 0 }, "=Sheet1!A1");
        rewriteDocumentFormulasForSheetRename(doc, "Sheet1", "A1");
        expect(doc.getCell("S1", { row: 0, col: 0 }).formula).toBe("='A1'!A1");
      }

      {
        const doc = new DocumentController();
        doc.setCellFormula("S1", { row: 0, col: 0 }, "=Sheet1!A1");
        rewriteDocumentFormulasForSheetRename(doc, "Sheet1", "R1C1");
        expect(doc.getCell("S1", { row: 0, col: 0 }).formula).toBe("='R1C1'!A1");
      }
    });

    it("rewrites unquoted Unicode sheet refs (e.g. résumé)", () => {
      const doc = new DocumentController();
      doc.setCellFormula("S1", { row: 0, col: 0 }, "=résumé!A1+1");

      rewriteDocumentFormulasForSheetRename(doc, "Résumé", "Data");

      expect(doc.getCell("S1", { row: 0, col: 0 }).formula).toBe("=Data!A1+1");
    });
  });

  describe("rewriteDocumentFormulasForSheetDelete", () => {
    it("rewrites formulas referencing the deleted sheet to #REF!", () => {
      const doc = new DocumentController();
      doc.setCellFormula("S1", { row: 0, col: 0 }, "='My Sheet'!A1+1");
      doc.setCellFormula("S1", { row: 0, col: 1 }, "=SUM(Sheet1:Sheet3!A1)");

      rewriteDocumentFormulasForSheetDelete(doc, "My Sheet", ["Sheet1", "Sheet2", "Sheet3", "My Sheet"]);

      expect(doc.getCell("S1", { row: 0, col: 0 }).formula).toBe("=#REF!+1");
      // Unrelated formula stays intact.
      expect(doc.getCell("S1", { row: 0, col: 1 }).formula).toBe("=SUM(Sheet1:Sheet3!A1)");
    });

    it("rewrites deleted sheet endpoints in 3D refs", () => {
      const doc = new DocumentController();
      doc.setCellFormula("S1", { row: 0, col: 0 }, "=SUM(Sheet1:Sheet3!A1)");

      rewriteDocumentFormulasForSheetDelete(doc, "Sheet1", ["Sheet1", "Sheet2", "Sheet3"]);

      expect(doc.getCell("S1", { row: 0, col: 0 }).formula).toBe("=SUM(Sheet2:Sheet3!A1)");
    });
  });
});
