use std::path::Path;

use pretty_assertions::assert_eq;

use formula_model::{Cell, CellRef, Style};

#[test]
fn roundtrip_fixtures_no_critical_diffs() -> Result<(), Box<dyn std::error::Error>> {
    let fixtures_root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx");
    let fixtures = xlsx_diff::collect_fixture_paths(&fixtures_root)?;
    assert!(!fixtures.is_empty(), "no fixtures found");

    for fixture in fixtures {
        let mut pkg = formula_xlsx::WorkbookPackage::load(&fixture)?;
        let tmpdir = tempfile::tempdir()?;
        let out = tmpdir.path().join("roundtripped.xlsx");
        pkg.save(&out)?;

        let report = xlsx_diff::diff_workbooks(&fixture, &out)?;
        if report.has_at_least(xlsx_diff::Severity::Critical) {
            eprintln!("Critical diffs detected for fixture {}", fixture.display());
            for diff in report
                .differences
                .iter()
                .filter(|d| d.severity == xlsx_diff::Severity::Critical)
            {
                eprintln!("{diff}");
            }
            panic!("fixture {} did not round-trip cleanly", fixture.display());
        }
    }

    Ok(())
}

#[test]
fn styles_part_appends_xfs_deterministically() -> Result<(), Box<dyn std::error::Error>> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/styles/styles.xlsx");
    let mut pkg = formula_xlsx::WorkbookPackage::load(&fixture)?;

    // Create a new style derived from default with italic applied.
    let new_style_id = pkg.workbook.intern_style(Style {
        font: Some(formula_model::Font {
            italic: true,
            ..Default::default()
        }),
        ..Default::default()
    });

    let first = pkg.xf_index_for_style(new_style_id)?;
    let second = pkg.xf_index_for_style(new_style_id)?;
    assert_eq!(first, second);

    // Ensure cellXfs count only increased by one (original fixture has 2).
    let xml_bytes = pkg.styles().to_xml_bytes();
    let xml = std::str::from_utf8(&xml_bytes)?;
    let doc = roxmltree::Document::parse(xml)?;
    let cell_xfs = doc
        .root_element()
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "cellXfs")
        .expect("cellXfs missing");
    assert_eq!(cell_xfs.attribute("count"), Some("3"));

    // Apply the new style to A1 and ensure it survives a save+reload.
    let sheet_id = pkg.workbook.sheets[0].id;
    let sheet = pkg.workbook.sheet_mut(sheet_id).unwrap();
    let mut cell = sheet
        .cell(CellRef::from_a1("A1").unwrap())
        .cloned()
        .unwrap_or_else(|| Cell::default());
    cell.style_id = new_style_id;
    sheet.set_cell(CellRef::from_a1("A1").unwrap(), cell);

    let tmpdir = tempfile::tempdir()?;
    let out = tmpdir.path().join("styled.xlsx");
    pkg.save(&out)?;

    let pkg2 = formula_xlsx::WorkbookPackage::load(&out)?;
    let sheet2_id = pkg2.workbook.sheets[0].id;
    let sheet2 = pkg2.workbook.sheet(sheet2_id).unwrap();
    let cell2 = sheet2.cell(CellRef::from_a1("A1").unwrap()).unwrap();
    let style2 = pkg2.workbook.styles.get(cell2.style_id).unwrap();
    assert_eq!(style2.font.as_ref().unwrap().italic, true);

    Ok(())
}
