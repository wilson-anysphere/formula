use std::fs;
use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::{CellRef, CellValue};
use formula_xlsx::{
    patch_xlsx_streaming, patch_xlsx_streaming_with_recalc_policy, RecalcPolicy, WorksheetCellPatch,
};
use zip::ZipArchive;

fn read_worksheet_xml(bytes: &[u8], part_name: &str) -> Result<String, Box<dyn std::error::Error>> {
    let mut archive = ZipArchive::new(Cursor::new(bytes))?;
    let mut xml = String::new();
    archive.by_name(part_name)?.read_to_string(&mut xml)?;
    Ok(xml)
}

#[test]
fn streaming_recalc_policy_can_clear_cached_values_for_formula_patches(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/rt_simple.xlsx");
    let bytes = fs::read(&fixture_path)?;

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("C1")?,
        CellValue::Number(2.0),
        Some("=1+1".to_string()),
    );

    // When enabled, a formula patch should omit the cached `<v>` value.
    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_with_recalc_policy(
        Cursor::new(bytes.clone()),
        &mut out,
        &[patch.clone()],
        RecalcPolicy {
            clear_cached_values_on_formula_change: true,
            ..Default::default()
        },
    )?;
    let sheet_xml = read_worksheet_xml(out.get_ref(), "xl/worksheets/sheet1.xml")?;
    let xml_doc = roxmltree::Document::parse(&sheet_xml)?;
    let patched_cell = xml_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("C1"))
        .expect("patched cell C1 should exist");
    assert!(
        patched_cell
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "f"),
        "patched cell should contain a <f> element"
    );
    assert!(
        !patched_cell
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "v"),
        "cached value should be omitted when clear_cached_values_on_formula_change is enabled"
    );

    // Default behavior should continue writing cached `<v>` values.
    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;
    let sheet_xml = read_worksheet_xml(out.get_ref(), "xl/worksheets/sheet1.xml")?;
    let xml_doc = roxmltree::Document::parse(&sheet_xml)?;
    let patched_cell = xml_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("C1"))
        .expect("patched cell C1 should exist");
    assert!(
        patched_cell
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "f"),
        "patched cell should contain a <f> element"
    );
    assert!(
        patched_cell
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "v"),
        "default streaming patcher should preserve cached <v> values"
    );

    Ok(())
}

