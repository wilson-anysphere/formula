use formula_model::{CellRef, CellValue};
use formula_xlsx::{CellPatch, RecalcPolicy, WorkbookCellPatches, XlsxPackage};

#[test]
fn in_memory_patcher_can_clear_cached_values_for_formula_edits() -> Result<(), Box<dyn std::error::Error>>
{
    let bytes = include_bytes!("fixtures/rt_simple.xlsx");

    // Patch C1 to a new formula + value.
    let mut pkg = XlsxPackage::from_bytes(bytes)?;
    let first_sheet = pkg
        .worksheet_parts()?
        .into_iter()
        .next()
        .expect("fixture should have at least one sheet");
    let sheet_name = first_sheet.name;
    let worksheet_part = first_sheet.worksheet_part;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        sheet_name.clone(),
        CellRef::from_a1("C1")?,
        CellPatch::set_value_with_formula(CellValue::Number(2.0), "=1+1"),
    );

    // When enabled, a formula change should omit the cached `<v>` value.
    pkg.apply_cell_patches_with_recalc_policy(
        &patches,
        RecalcPolicy {
            clear_cached_values_on_formula_change: true,
            ..Default::default()
        },
    )?;

    let sheet_xml = std::str::from_utf8(
        pkg.part(&worksheet_part)
            .expect("worksheet should exist"),
    )?;
    let xml_doc = roxmltree::Document::parse(sheet_xml)?;
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

    // Control: default policy should continue writing cached `<v>` values.
    let mut pkg = XlsxPackage::from_bytes(bytes)?;
    let first_sheet = pkg
        .worksheet_parts()?
        .into_iter()
        .next()
        .expect("fixture should have at least one sheet");
    let sheet_name = first_sheet.name;
    let worksheet_part = first_sheet.worksheet_part;
    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        sheet_name,
        CellRef::from_a1("C1")?,
        CellPatch::set_value_with_formula(CellValue::Number(2.0), "=1+1"),
    );
    pkg.apply_cell_patches(&patches)?;

    let sheet_xml = std::str::from_utf8(
        pkg.part(&worksheet_part)
            .expect("worksheet should exist"),
    )?;
    let xml_doc = roxmltree::Document::parse(sheet_xml)?;
    let patched_cell = xml_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("C1"))
        .expect("patched cell C1 should exist");
    assert!(
        patched_cell
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "v"),
        "default in-memory patcher should preserve cached <v> values"
    );

    Ok(())
}
