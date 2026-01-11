use std::fs;
use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::{CellRef, CellValue};
use formula_xlsx::{patch_xlsx_streaming, load_from_bytes, WorksheetCellPatch, XlsxDocument};
use zip::ZipArchive;

#[test]
fn streaming_noop_roundtrip_has_no_critical_diffs() -> Result<(), Box<dyn std::error::Error>> {
    let fixtures = [
        "calc_settings.xlsx",
        "comments.xlsx",
        "conditional_formatting_2007.xlsx",
        "rt_macro.xlsm",
    ];

    let tmpdir = tempfile::tempdir()?;

    for fixture_name in fixtures {
        let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures")
            .join(fixture_name);
        let bytes = fs::read(&fixture_path)?;

        let out_path = tmpdir.path().join(format!("roundtrip-{fixture_name}"));
        let out_file = fs::File::create(&out_path)?;

        patch_xlsx_streaming(Cursor::new(bytes), out_file, &[])?;

        let report = xlsx_diff::diff_workbooks(&fixture_path, &out_path)?;
        if report.has_at_least(xlsx_diff::Severity::Critical) {
            eprintln!(
                "Critical diffs detected for streaming no-op fixture {}",
                fixture_path.display()
            );
            for diff in report
                .differences
                .iter()
                .filter(|d| d.severity == xlsx_diff::Severity::Critical)
            {
                eprintln!("{diff}");
            }
            panic!("streaming no-op did not round-trip cleanly: {}", fixture_path.display());
        }
    }

    Ok(())
}

#[test]
fn streaming_patch_updates_cell_value_and_formula() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/merged-cells.xlsx");
    let bytes = fs::read(&fixture_path)?;

    let orig = load_from_bytes(&bytes)?;
    let sheet_id = orig.workbook.sheets[0].id;
    let sheet = orig.workbook.sheet(sheet_id).unwrap();
    let a1 = CellRef::from_a1("A1")?;
    let orig_style = sheet
        .cell(a1)
        .map(|c| c.style_id)
        .unwrap_or_default();

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        a1,
        CellValue::Number(2.0),
        Some("=1+1".to_string()),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;

    let out_bytes = out.get_ref();
    let doc = load_from_bytes(out_bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet(sheet_id).unwrap();
    let cell = sheet
        .cell(CellRef::from_a1("A1")?)
        .expect("patched cell should exist");

    assert_eq!(cell.value, CellValue::Number(2.0));
    assert_eq!(cell.formula.as_deref(), Some("1+1"));
    assert_eq!(cell.style_id, orig_style, "patcher should preserve cell style");

    Ok(())
}

#[test]
fn streaming_patch_expands_dimension_when_writing_out_of_range_cell(
) -> Result<(), Box<dyn std::error::Error>> {
    // Build a minimal in-memory workbook (via the existing writer) where dimension is `A1:A1`.
    let mut workbook = formula_model::Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1");
    let sheet = workbook.sheet_mut(sheet_id).unwrap();
    sheet.set_cell(CellRef::from_a1("A1")?, formula_model::Cell::new(CellValue::Number(1.0)));

    let doc = XlsxDocument::new(workbook);
    let bytes = doc.save_to_vec()?;

    // Patch a cell well outside the original used range.
    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("C3")?,
        CellValue::Number(9.0),
        None,
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;

    let mut archive = ZipArchive::new(Cursor::new(out.into_inner()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    let xml_doc = roxmltree::Document::parse(&sheet_xml)?;
    let worksheet = xml_doc.root_element();
    let dimension = worksheet
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "dimension")
        .expect("dimension element should exist");
    assert_eq!(
        dimension.attribute("ref"),
        Some("A1:C3"),
        "dimension should expand to cover patched cell"
    );

    Ok(())
}
