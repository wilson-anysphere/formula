use std::path::Path;

use formula_model::{Cell, CellRef, Font, Style};

#[test]
fn xlsx_document_style_roundtrip_is_stable() -> Result<(), Box<dyn std::error::Error>> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/styles/styles.xlsx");

    let mut doc = formula_xlsx::load_from_path(&fixture)?;

    // Create a new style derived from default with italic applied.
    let new_style_id = doc.workbook.intern_style(Style {
        font: Some(Font {
            italic: true,
            ..Default::default()
        }),
        ..Default::default()
    });

    // Apply the new style to A1.
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet_mut(sheet_id).unwrap();
    let cell_ref = CellRef::from_a1("A1").unwrap();
    let mut cell = sheet.cell(cell_ref).cloned().unwrap_or_else(Cell::default);
    cell.style_id = new_style_id;
    sheet.set_cell(cell_ref, cell);

    // Save to disk so we can reload + diff parts.
    let tmpdir = tempfile::tempdir()?;
    let out1 = tmpdir.path().join("styled.xlsx");
    std::fs::write(&out1, doc.save_to_vec()?)?;

    // Reload and assert the style is still present and applied.
    let doc2 = formula_xlsx::load_from_path(&out1)?;
    let sheet2_id = doc2.workbook.sheets[0].id;
    let sheet2 = doc2.workbook.sheet(sheet2_id).unwrap();
    let cell2 = sheet2.cell(cell_ref).unwrap();
    let style2 = doc2.workbook.styles.get(cell2.style_id).unwrap();
    assert!(style2.font.as_ref().is_some_and(|font| font.italic));

    // Save again to validate that the style mapping doesn't churn (xf indices remain stable).
    let out2 = tmpdir.path().join("styled_roundtrip.xlsx");
    std::fs::write(&out2, doc2.save_to_vec()?)?;

    let report = xlsx_diff::diff_workbooks(&out1, &out2)?;
    if report.has_at_least(xlsx_diff::Severity::Critical) {
        eprintln!("Critical diffs detected for styled roundtrip");
        for diff in report
            .differences
            .iter()
            .filter(|d| d.severity == xlsx_diff::Severity::Critical)
        {
            eprintln!("{diff}");
        }
        panic!("styled workbook did not round-trip cleanly");
    }

    Ok(())
}

