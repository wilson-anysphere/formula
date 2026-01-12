use std::fs;
use std::io::{Cursor, Read, Write};
use std::path::Path;

use formula_model::rich_text::{RichText, RichTextRunStyle};
use formula_model::Color;
use formula_model::{CellRef, CellValue};
use formula_xlsx::{
    load_from_bytes, patch_xlsx_streaming, patch_xlsx_streaming_workbook_cell_patches, CellPatch,
    WorkbookCellPatches, WorksheetCellPatch, XlsxDocument,
};
use zip::ZipArchive;

fn build_minimal_xlsx(sheet_xml: &str) -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn worksheet_cell_value(sheet_xml: &str, cell_ref: &str) -> Option<String> {
    let xml_doc = roxmltree::Document::parse(sheet_xml).ok()?;
    let cell = xml_doc.descendants().find(|n| {
        n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some(cell_ref)
    })?;
    let v = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")?;
    Some(v.text().unwrap_or_default().to_string())
}

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
        Some(" =1+1".to_string()),
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

    let mut archive = ZipArchive::new(Cursor::new(out.get_ref()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;
    let xml_doc = roxmltree::Document::parse(&sheet_xml)?;
    let patched_cell = xml_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("A1 cell should exist");
    let f = patched_cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "f")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert!(
        !f.trim_start().starts_with('='),
        "patched <f> text must not include a leading '=' (got {f:?})"
    );
    assert_eq!(f, "1+1");

    Ok(())
}

#[test]
fn streaming_patch_expands_dimension_when_writing_out_of_range_cell(
) -> Result<(), Box<dyn std::error::Error>> {
    // Build a minimal in-memory workbook (via the existing writer) where dimension is `A1:A1`.
    let mut workbook = formula_model::Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").unwrap();
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

#[test]
fn streaming_patch_normalizes_formula_with_xlfn_prefixes() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = include_bytes!("fixtures/rt_simple.xlsx");

    let patches = [
        WorksheetCellPatch::new(
            "xl/worksheets/sheet1.xml",
            CellRef::from_a1("C1")?,
            CellValue::Number(1.0),
            Some("=SEQUENCE(3)".to_string()),
        ),
        WorksheetCellPatch::new(
            "xl/worksheets/sheet1.xml",
            CellRef::from_a1("C2")?,
            CellValue::Number(1.0),
            Some(r#"=TEXTAFTER("a_b","_")"#.to_string()),
        ),
        WorksheetCellPatch::new(
            "xl/worksheets/sheet1.xml",
            CellRef::from_a1("C3")?,
            CellValue::Number(1.0),
            Some(r#"=TEXTBEFORE("a_b","_")"#.to_string()),
        ),
        WorksheetCellPatch::new(
            "xl/worksheets/sheet1.xml",
            CellRef::from_a1("C4")?,
            CellValue::Number(1.0),
            Some("=VALUETOTEXT(1)".to_string()),
        ),
        WorksheetCellPatch::new(
            "xl/worksheets/sheet1.xml",
            CellRef::from_a1("C5")?,
            CellValue::Number(1.0),
            Some("=FORECAST.ETS(1,2,3)".to_string()),
        ),
        WorksheetCellPatch::new(
            "xl/worksheets/sheet1.xml",
            CellRef::from_a1("C6")?,
            CellValue::Number(1.0),
            Some("=FORECAST.ETS.CONFINT(1,2,3)".to_string()),
        ),
        WorksheetCellPatch::new(
            "xl/worksheets/sheet1.xml",
            CellRef::from_a1("C7")?,
            CellValue::Number(1.0),
            Some("=FORECAST.ETS.SEASONALITY(1,2)".to_string()),
        ),
        WorksheetCellPatch::new(
            "xl/worksheets/sheet1.xml",
            CellRef::from_a1("C8")?,
            CellValue::Number(1.0),
            Some("=FORECAST.ETS.STAT(1,2)".to_string()),
        ),
    ];

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes.as_slice()), &mut out, &patches)?;

    // Ensure the stored formula uses the _xlfn prefix.
    let mut archive = ZipArchive::new(Cursor::new(out.get_ref()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;
    let xml_doc = roxmltree::Document::parse(&sheet_xml)?;
    for (cell_ref, expected_file_text) in [
        ("C1", "_xlfn.SEQUENCE(3)"),
        ("C2", r#"_xlfn.TEXTAFTER("a_b","_")"#),
        ("C3", r#"_xlfn.TEXTBEFORE("a_b","_")"#),
        ("C4", "_xlfn.VALUETOTEXT(1)"),
        ("C5", "_xlfn.FORECAST.ETS(1,2,3)"),
        ("C6", "_xlfn.FORECAST.ETS.CONFINT(1,2,3)"),
        ("C7", "_xlfn.FORECAST.ETS.SEASONALITY(1,2)"),
        ("C8", "_xlfn.FORECAST.ETS.STAT(1,2)"),
    ] {
        let cell = xml_doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some(cell_ref))
            .unwrap_or_else(|| panic!("{cell_ref} cell should exist"));
        let f = cell
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "f")
            .and_then(|n| n.text())
            .unwrap_or_default();
        assert_eq!(
            f, expected_file_text,
            "expected {cell_ref} formula to be stored with _xlfn prefix"
        );
    }

    // Ensure the parsed model uses the display formula without the prefix.
    let doc = load_from_bytes(out.get_ref())?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet(sheet_id).unwrap();
    for (cell_ref, expected_display) in [
        ("C1", "SEQUENCE(3)"),
        ("C2", r#"TEXTAFTER("a_b","_")"#),
        ("C3", r#"TEXTBEFORE("a_b","_")"#),
        ("C4", "VALUETOTEXT(1)"),
        ("C5", "FORECAST.ETS(1,2,3)"),
        ("C6", "FORECAST.ETS.CONFINT(1,2,3)"),
        ("C7", "FORECAST.ETS.SEASONALITY(1,2)"),
        ("C8", "FORECAST.ETS.STAT(1,2)"),
    ] {
        let cell = sheet.cell(CellRef::from_a1(cell_ref)?).unwrap();
        assert_eq!(cell.formula.as_deref(), Some(expected_display));
    }

    Ok(())
}

#[test]
fn streaming_patch_detaches_textless_shared_formula() -> Result<(), Box<dyn std::error::Error>> {
    // Fixture contains a shared formula in A2 (master) and a textless shared reference in B2.
    let bytes = include_bytes!("fixtures/rt_simple.xlsx");

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("B2")?,
        CellValue::Number(2.0),
        Some("=1+1".to_string()),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes.as_slice()), &mut out, &[patch])?;

    let mut archive = ZipArchive::new(Cursor::new(out.get_ref()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    let xml_doc = roxmltree::Document::parse(&sheet_xml)?;
    let cell = xml_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("B2"))
        .expect("B2 cell should exist");
    let f = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "f")
        .expect("B2 should have a formula");
    assert_eq!(
        f.attribute("t"),
        None,
        "patched textless shared formula should become a standalone formula"
    );
    assert_eq!(f.attribute("si"), None);

    Ok(())
}

#[test]
fn streaming_patch_preserves_prefixes_when_expanding_empty_sheetdata(
) -> Result<(), Box<dyn std::error::Error>> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:dimension ref="A1"/>
  <x:sheetData/>
</x:worksheet>"#;

    let bytes = build_minimal_xlsx(worksheet_xml);

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A1")?,
        CellValue::Number(2.0),
        None,
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;

    let mut archive = ZipArchive::new(Cursor::new(out.into_inner()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;
    roxmltree::Document::parse(&sheet_xml)?;
    assert!(
        sheet_xml.contains("<x:sheetData>") && sheet_xml.contains("</x:sheetData>"),
        "expected prefixed sheetData expansion, got: {sheet_xml}"
    );
    assert!(
        sheet_xml.contains("<x:row") && sheet_xml.contains("<x:c r=\"A1\""),
        "expected prefixed row/cell insertion, got: {sheet_xml}"
    );
    assert!(
        sheet_xml.contains("<x:v>2</x:v>"),
        "expected prefixed value element insertion, got: {sheet_xml}"
    );

    Ok(())
}

#[test]
fn streaming_patch_preserves_prefixes_when_inserting_missing_sheetdata(
) -> Result<(), Box<dyn std::error::Error>> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:dimension ref="A1"/>
</x:worksheet>"#;

    let bytes = build_minimal_xlsx(worksheet_xml);

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A1")?,
        CellValue::Number(2.0),
        None,
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;

    let mut archive = ZipArchive::new(Cursor::new(out.into_inner()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;
    roxmltree::Document::parse(&sheet_xml)?;
    assert!(
        sheet_xml.contains("<x:sheetData>") && sheet_xml.contains("</x:sheetData>"),
        "expected prefixed sheetData insertion, got: {sheet_xml}"
    );
    assert!(
        sheet_xml.contains("<x:row") && sheet_xml.contains("<x:c r=\"A1\""),
        "expected prefixed row/cell insertion, got: {sheet_xml}"
    );
    assert!(
        sheet_xml.contains("<x:v>2</x:v>"),
        "expected prefixed value element insertion, got: {sheet_xml}"
    );

    Ok(())
}

#[test]
fn streaming_patch_drops_calc_chain_when_formulas_change() -> Result<(), Box<dyn std::error::Error>>
{
    use zip::result::ZipError;

    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/calc_settings.xlsx");
    let bytes = fs::read(&fixture_path)?;

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A1")?,
        CellValue::Number(2.0),
        Some("=1+1".to_string()),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;

    let mut archive = ZipArchive::new(Cursor::new(out.get_ref()))?;
    assert!(
        matches!(archive.by_name("xl/calcChain.xml").err(), Some(ZipError::FileNotFound)),
        "expected streaming patcher to drop xl/calcChain.xml after formula edits"
    );

    let mut workbook_xml = String::new();
    archive
        .by_name("xl/workbook.xml")?
        .read_to_string(&mut workbook_xml)?;
    let workbook_doc = roxmltree::Document::parse(&workbook_xml)?;
    let calc_pr = workbook_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "calcPr")
        .expect("workbook.xml should include <calcPr>");
    assert_eq!(
        calc_pr.attribute("fullCalcOnLoad"),
        Some("1"),
        "streaming patcher should force fullCalcOnLoad after formula edits"
    );

    let mut rels_xml = String::new();
    archive
        .by_name("xl/_rels/workbook.xml.rels")?
        .read_to_string(&mut rels_xml)?;
    assert!(
        !rels_xml.contains("relationships/calcChain"),
        "workbook.xml.rels should not contain calcChain relationship after formula edits"
    );

    let mut content_types = String::new();
    archive
        .by_name("[Content_Types].xml")?
        .read_to_string(&mut content_types)?;
    assert!(
        !content_types.contains("calcChain.xml"),
        "[Content_Types].xml should not reference calcChain.xml after formula edits"
    );

    Ok(())
}

#[test]
fn streaming_patch_drops_calc_chain_when_formulas_are_removed(
) -> Result<(), Box<dyn std::error::Error>> {
    use zip::result::ZipError;

    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/recalc_policy.xlsx");
    let bytes = fs::read(&fixture_path)?;

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("C1")?,
        CellValue::Number(123.0),
        None,
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;

    let mut archive = ZipArchive::new(Cursor::new(out.get_ref()))?;
    assert!(
        matches!(archive.by_name("xl/calcChain.xml").err(), Some(ZipError::FileNotFound)),
        "expected streaming patcher to drop xl/calcChain.xml after removing formulas"
    );

    let mut workbook_xml = String::new();
    archive
        .by_name("xl/workbook.xml")?
        .read_to_string(&mut workbook_xml)?;
    let workbook_doc = roxmltree::Document::parse(&workbook_xml)?;
    let calc_pr = workbook_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "calcPr")
        .expect("workbook.xml should include <calcPr>");
    assert_eq!(
        calc_pr.attribute("fullCalcOnLoad"),
        Some("1"),
        "streaming patcher should force fullCalcOnLoad after removing formulas"
    );

    let mut rels_xml = String::new();
    archive
        .by_name("xl/_rels/workbook.xml.rels")?
        .read_to_string(&mut rels_xml)?;
    assert!(
        !rels_xml.contains("relationships/calcChain"),
        "workbook.xml.rels should not contain calcChain relationship after removing formulas"
    );

    let mut content_types = String::new();
    archive
        .by_name("[Content_Types].xml")?
        .read_to_string(&mut content_types)?;
    assert!(
        !content_types.contains("calcChain.xml"),
        "[Content_Types].xml should not reference calcChain.xml after removing formulas"
    );

    Ok(())
}

#[test]
fn streaming_workbook_cell_patches_resolve_sheet_names() -> Result<(), Box<dyn std::error::Error>> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/basic.xlsx");
    let bytes = fs::read(&fixture)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        // Sheet names are case-insensitive in Excel; accept patches keyed by any casing.
        "sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value(CellValue::Number(42.0)),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes), &mut out, &patches)?;

    let tmpdir = tempfile::tempdir()?;
    let out_path = tmpdir.path().join("patched.xlsx");
    fs::write(&out_path, out.get_ref())?;

    let report = xlsx_diff::diff_workbooks(&fixture, &out_path)?;
    for diff in &report.differences {
        assert_ne!(diff.kind, "missing_part", "missing part {}", diff.part);
        assert_ne!(diff.kind, "extra_part", "extra part {}", diff.part);
    }

    let changed_parts: std::collections::BTreeSet<String> = report
        .differences
        .iter()
        .map(|d| d.part.clone())
        .collect();
    assert_eq!(
        changed_parts,
        std::collections::BTreeSet::from(["xl/worksheets/sheet1.xml".to_string()])
    );

    Ok(())
}

#[test]
fn streaming_workbook_cell_patches_resolve_worksheet_part(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/multi-sheet.xlsx");
    let bytes = fs::read(&fixture)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "xl/worksheets/sheet2.xml",
        CellRef::from_a1("A1")?,
        CellPatch::set_value(CellValue::Number(123.0)),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes), &mut out, &patches)?;

    let tmpdir = tempfile::tempdir()?;
    let out_path = tmpdir.path().join("patched.xlsx");
    fs::write(&out_path, out.get_ref())?;

    let report = xlsx_diff::diff_workbooks(&fixture, &out_path)?;
    for diff in &report.differences {
        assert_ne!(diff.kind, "missing_part", "missing part {}", diff.part);
        assert_ne!(diff.kind, "extra_part", "extra part {}", diff.part);
    }

    let changed_parts: std::collections::BTreeSet<String> = report
        .differences
        .iter()
        .map(|d| d.part.clone())
        .collect();
    assert_eq!(
        changed_parts,
        std::collections::BTreeSet::from(["xl/worksheets/sheet2.xml".to_string()])
    );

    let mut archive = ZipArchive::new(Cursor::new(out.get_ref()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet2.xml")?
        .read_to_string(&mut sheet_xml)?;
    assert_eq!(worksheet_cell_value(&sheet_xml, "A1"), Some("123".to_string()));

    Ok(())
}

#[test]
fn streaming_workbook_cell_patches_resolve_rel_id() -> Result<(), Box<dyn std::error::Error>> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/multi-sheet.xlsx");
    let bytes = fs::read(&fixture)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "rId2",
        CellRef::from_a1("A1")?,
        CellPatch::set_value(CellValue::Number(321.0)),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes), &mut out, &patches)?;

    let tmpdir = tempfile::tempdir()?;
    let out_path = tmpdir.path().join("patched.xlsx");
    fs::write(&out_path, out.get_ref())?;

    let report = xlsx_diff::diff_workbooks(&fixture, &out_path)?;
    for diff in &report.differences {
        assert_ne!(diff.kind, "missing_part", "missing part {}", diff.part);
        assert_ne!(diff.kind, "extra_part", "extra part {}", diff.part);
    }

    let changed_parts: std::collections::BTreeSet<String> = report
        .differences
        .iter()
        .map(|d| d.part.clone())
        .collect();
    assert_eq!(
        changed_parts,
        std::collections::BTreeSet::from(["xl/worksheets/sheet2.xml".to_string()])
    );

    let mut archive = ZipArchive::new(Cursor::new(out.get_ref()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet2.xml")?
        .read_to_string(&mut sheet_xml)?;
    assert_eq!(worksheet_cell_value(&sheet_xml, "A1"), Some("321".to_string()));

    Ok(())
}

#[test]
fn streaming_patch_preserves_unknown_cell_types() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/date-type.xlsx");
    let bytes = fs::read(&fixture_path)?;

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("C1")?,
        CellValue::String("2028-05-06T00:00:00Z".to_string()),
        None,
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;

    let mut archive = ZipArchive::new(Cursor::new(out.get_ref()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    let xml_doc = roxmltree::Document::parse(&sheet_xml)?;
    let cell = xml_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("C1"))
        .expect("C1 cell should exist");
    assert_eq!(cell.attribute("t"), Some("d"));
    let v = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(v, "2028-05-06T00:00:00Z");

    Ok(())
}

#[test]
fn streaming_patch_inserts_sheet_data_when_missing() -> Result<(), Box<dyn std::error::Error>> {
    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipWriter};

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    // Worksheet XML intentionally omits `<sheetData>` to ensure the streaming patcher inserts it.
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;

    let mut input = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut input);
        let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);
        zip.start_file("xl/workbook.xml", options)?;
        zip.write_all(workbook_xml.as_bytes())?;
        zip.start_file("xl/_rels/workbook.xml.rels", options)?;
        zip.write_all(workbook_rels.as_bytes())?;
        zip.start_file("xl/worksheets/sheet1.xml", options)?;
        zip.write_all(worksheet_xml.as_bytes())?;
        zip.finish()?;
    }
    input.set_position(0);

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A1")?,
        CellValue::Number(1.0),
        None,
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(input, &mut out, &[patch])?;

    let mut archive = ZipArchive::new(Cursor::new(out.into_inner()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;

    assert!(
        sheet_xml.contains("<sheetData>"),
        "patched worksheet should include inserted <sheetData>"
    );
    assert!(
        sheet_xml.contains(r#"<c r=\"A1\""#) || sheet_xml.contains(r#"<c r="A1""#),
        "patched worksheet should include the new cell"
    );

    Ok(())
}

#[test]
fn streaming_patch_preserves_shared_string_cells() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/shared-strings.xlsx");
    let bytes = fs::read(&fixture_path)?;

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A1")?,
        CellValue::String("Patched".to_string()),
        None,
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;

    let mut archive = ZipArchive::new(Cursor::new(out.get_ref()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;
    let doc = roxmltree::Document::parse(&sheet_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let cell = doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some("A1"))
        .expect("A1 should exist");
    assert_eq!(
        cell.attribute("t"),
        Some("s"),
        "patched shared-string cell should remain t=\"s\""
    );

    let mut shared_strings_xml = String::new();
    archive
        .by_name("xl/sharedStrings.xml")?
        .read_to_string(&mut shared_strings_xml)?;
    assert!(
        shared_strings_xml.contains("Patched"),
        "sharedStrings.xml should contain the new string"
    );
    let shared_doc = roxmltree::Document::parse(&shared_strings_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let sst = shared_doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "sst")))
        .expect("sharedStrings.xml should contain <sst>");
    assert_eq!(
        sst.attribute("count"),
        Some("2"),
        "sharedStrings/@count should remain the original reference count"
    );
    assert_eq!(
        sst.attribute("uniqueCount"),
        Some("3"),
        "sharedStrings/@uniqueCount should reflect the appended entry"
    );

    let tmpdir = tempfile::tempdir()?;
    let out_path = tmpdir.path().join("patched.xlsx");
    fs::write(&out_path, out.get_ref())?;
    let report = xlsx_diff::diff_workbooks(&fixture_path, &out_path)?;
    for diff in &report.differences {
        assert_ne!(diff.kind, "missing_part", "missing part {}", diff.part);
        assert_ne!(diff.kind, "extra_part", "extra part {}", diff.part);
    }

    let changed_parts: std::collections::BTreeSet<String> = report
        .differences
        .iter()
        .map(|d| d.part.clone())
        .collect();
    assert_eq!(
        changed_parts,
        std::collections::BTreeSet::from([
            "xl/sharedStrings.xml".to_string(),
            "xl/worksheets/sheet1.xml".to_string(),
        ])
    );

    Ok(())
}

#[test]
fn streaming_patch_writes_rich_text_via_shared_strings() -> Result<(), Box<dyn std::error::Error>>
{
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/styles/rich-text-shared-strings.xlsx");
    let bytes = fs::read(&fixture_path)?;

    let rich = RichText::from_segments(vec![
        ("Red".to_string(), RichTextRunStyle::default()),
        (
            "Blue".to_string(),
            RichTextRunStyle {
                bold: Some(true),
                color: Some(Color::new_argb(0xFFFF0000)),
                ..Default::default()
            },
        ),
    ]);

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A2")?,
        CellValue::RichText(rich),
        None,
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;

    let mut archive = ZipArchive::new(Cursor::new(out.get_ref()))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;
    let doc = roxmltree::Document::parse(&sheet_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let cell = doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some("A2"))
        .expect("A2 should exist");
    assert_eq!(cell.attribute("t"), Some("s"));

    let mut shared_strings_xml = String::new();
    archive
        .by_name("xl/sharedStrings.xml")?
        .read_to_string(&mut shared_strings_xml)?;

    assert!(
        shared_strings_xml.contains("Red") && shared_strings_xml.contains("Blue"),
        "sharedStrings.xml should contain rich text segments"
    );
    assert!(
        shared_strings_xml.contains("<r>"),
        "sharedStrings.xml should include rich text runs"
    );

    Ok(())
}

#[test]
fn streaming_patch_writes_inline_rich_text_when_shared_strings_missing(
) -> Result<(), Box<dyn std::error::Error>> {
    use zip::result::ZipError;

    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/basic.xlsx");
    let bytes = fs::read(&fixture_path)?;

    let rich = RichText::from_segments(vec![
        ("Hello ".to_string(), RichTextRunStyle::default()),
        (
            "World".to_string(),
            RichTextRunStyle {
                bold: Some(true),
                ..Default::default()
            },
        ),
    ]);

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("B1")?,
        CellValue::RichText(rich),
        None,
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;

    let mut archive = ZipArchive::new(Cursor::new(out.get_ref()))?;
    assert!(
        matches!(archive.by_name("xl/sharedStrings.xml").err(), Some(ZipError::FileNotFound)),
        "workbook without sharedStrings.xml should not gain one during inline rich text patching"
    );

    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;
    let doc = roxmltree::Document::parse(&sheet_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let cell = doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some("B1"))
        .expect("B1 should exist");
    assert_eq!(
        cell.attribute("t"),
        Some("inlineStr"),
        "expected rich text to be written as inlineStr when sharedStrings are missing"
    );
    let inline = cell
        .children()
        .find(|n| n.has_tag_name((ns, "is")))
        .expect("inline string <is> should exist");
    assert!(
        inline
            .descendants()
            .any(|n| n.has_tag_name((ns, "r"))),
        "expected inline rich text to contain <r> runs"
    );

    Ok(())
}

const PHONETIC_MARKER: &str = "PHO_MARKER_123";
const EXTLST_MARKER: &str = "EXT_MARKER_456";

fn build_phonetic_shared_strings_fixture_xlsx() -> Vec<u8> {
    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipWriter};

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let shared_strings_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <t>Base</t>
    <rPh sb="0" eb="4"><t>{PHONETIC_MARKER}</t></rPh>
  </si>
  <extLst>
    <ext uri="{{{EXTLST_MARKER}}}">
      <marker>{EXTLST_MARKER}</marker>
    </ext>
  </extLst>
</sst>"#
    );

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/sharedStrings.xml", options).unwrap();
    zip.write_all(shared_strings_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn streaming_patch_preserves_phonetic_shared_strings_and_extlst(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_phonetic_shared_strings_fixture_xlsx();

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A2")?,
        CellValue::String("Patched".to_string()),
        None,
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes.clone()), &mut out, &[patch])?;

    let mut archive = ZipArchive::new(Cursor::new(out.get_ref()))?;
    let mut shared_strings_xml = String::new();
    archive
        .by_name("xl/sharedStrings.xml")?
        .read_to_string(&mut shared_strings_xml)?;
    assert!(
        shared_strings_xml.contains(PHONETIC_MARKER),
        "expected sharedStrings.xml to preserve phonetic subtree"
    );
    assert!(
        shared_strings_xml.contains(EXTLST_MARKER),
        "expected sharedStrings.xml to preserve <extLst> subtree"
    );
    assert!(
        shared_strings_xml.contains("Patched"),
        "expected sharedStrings.xml to include newly inserted string"
    );

    let patched_pos = shared_strings_xml
        .find("Patched")
        .expect("sharedStrings.xml should contain Patched");
    let extlst_pos = shared_strings_xml
        .find("<extLst")
        .expect("sharedStrings.xml should contain <extLst>");
    assert!(
        patched_pos < extlst_pos,
        "expected inserted <si> to appear before <extLst> (got sharedStrings.xml: {shared_strings_xml})"
    );

    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;
    let doc = roxmltree::Document::parse(&sheet_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let cell = doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some("A2"))
        .expect("A2 should exist");
    assert_eq!(cell.attribute("t"), Some("s"));

    let tmpdir = tempfile::tempdir()?;
    let input_path = tmpdir.path().join("input.xlsx");
    let out_path = tmpdir.path().join("patched.xlsx");
    fs::write(&input_path, &bytes)?;
    fs::write(&out_path, out.get_ref())?;

    let report = xlsx_diff::diff_workbooks(&input_path, &out_path)?;
    for diff in &report.differences {
        assert_ne!(diff.kind, "missing_part", "missing part {}", diff.part);
        assert_ne!(diff.kind, "extra_part", "extra part {}", diff.part);
    }

    let changed_parts: std::collections::BTreeSet<String> = report
        .differences
        .iter()
        .map(|d| d.part.clone())
        .collect();
    assert_eq!(
        changed_parts,
        std::collections::BTreeSet::from([
            "xl/sharedStrings.xml".to_string(),
            "xl/worksheets/sheet1.xml".to_string(),
        ])
    );

    Ok(())
}
