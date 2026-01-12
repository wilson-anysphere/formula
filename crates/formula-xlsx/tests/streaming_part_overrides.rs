use std::collections::{BTreeSet, HashMap};
use std::fs;
use std::io::{Cursor, Read};
use std::path::Path;

use formula_xlsx::{
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides,
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides_and_recalc_policy, CellPatch,
    PartOverride, RecalcPolicy, WorkbookCellPatches, XlsxPackage,
};
use formula_model::{CellRef, CellValue};
use zip::ZipArchive;

const POWER_QUERY_PART: &str = "xl/formula/power-query.xml";

fn fixture_basic_xlsx_bytes() -> Vec<u8> {
    let fixture_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/basic.xlsx");
    fs::read(&fixture_path).expect("read basic.xlsx fixture")
}

fn inject_power_query_part(base: &[u8], bytes: Vec<u8>) -> Vec<u8> {
    let mut pkg = XlsxPackage::from_bytes(base).expect("parse base package");
    pkg.set_part(POWER_QUERY_PART, bytes);
    pkg.write_to_bytes().expect("write injected package")
}

fn assert_parts_preserved_except(original: &[u8], patched: &[u8], except: &[&str]) {
    let original_pkg = XlsxPackage::from_bytes(original).expect("parse original package");
    let patched_pkg = XlsxPackage::from_bytes(patched).expect("parse patched package");

    let except: BTreeSet<&str> = except.iter().copied().collect();

    for (name, bytes) in original_pkg.parts() {
        if except.contains(name) {
            continue;
        }
        assert_eq!(
            Some(bytes),
            patched_pkg.part(name),
            "expected part {name} to be preserved byte-for-byte"
        );
    }
}

#[test]
fn streaming_part_override_replaces_power_query_xml_without_touching_other_parts(
) -> Result<(), Box<dyn std::error::Error>> {
    let base = fixture_basic_xlsx_bytes();
    let initial_xml = br#"<FormulaPowerQuery version="1"><![CDATA[{"queries":[{"id":"q1"}]}]]></FormulaPowerQuery>"#.to_vec();
    let input = inject_power_query_part(&base, initial_xml.clone());

    let updated_xml = br#"<FormulaPowerQuery version="1"><![CDATA[{"queries":[{"id":"q2"}]}]]></FormulaPowerQuery>"#.to_vec();
    let mut overrides = HashMap::new();
    overrides.insert(
        POWER_QUERY_PART.to_string(),
        PartOverride::Replace(updated_xml.clone()),
    );

    let patches = WorkbookCellPatches::default();
    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
        Cursor::new(input.clone()),
        &mut out,
        &patches,
        &overrides,
    )?;
    let out_bytes = out.into_inner();

    // Ensure the replaced part was updated.
    let out_pkg = XlsxPackage::from_bytes(&out_bytes)?;
    assert_eq!(
        out_pkg.part(POWER_QUERY_PART),
        Some(updated_xml.as_slice()),
        "expected power-query.xml to be replaced"
    );

    // Ensure everything else is preserved.
    assert_parts_preserved_except(&input, &out_bytes, &[POWER_QUERY_PART]);

    Ok(())
}

#[test]
fn streaming_part_override_removes_power_query_xml_without_touching_other_parts(
) -> Result<(), Box<dyn std::error::Error>> {
    let base = fixture_basic_xlsx_bytes();
    let initial_xml = br#"<FormulaPowerQuery version="1"><![CDATA[{"queries":[{"id":"q1"}]}]]></FormulaPowerQuery>"#.to_vec();
    let input = inject_power_query_part(&base, initial_xml);

    let mut overrides = HashMap::new();
    overrides.insert(POWER_QUERY_PART.to_string(), PartOverride::Remove);

    let patches = WorkbookCellPatches::default();
    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
        Cursor::new(input.clone()),
        &mut out,
        &patches,
        &overrides,
    )?;
    let out_bytes = out.into_inner();

    let out_pkg = XlsxPackage::from_bytes(&out_bytes)?;
    assert!(
        out_pkg.part(POWER_QUERY_PART).is_none(),
        "expected power-query.xml to be removed"
    );

    assert_parts_preserved_except(&input, &out_bytes, &[POWER_QUERY_PART]);
    Ok(())
}

#[test]
fn streaming_part_override_adds_power_query_xml_without_touching_other_parts(
) -> Result<(), Box<dyn std::error::Error>> {
    let input = fixture_basic_xlsx_bytes();

    let added_xml = br#"<FormulaPowerQuery version="1"><![CDATA[{"queries":[{"id":"added"}]}]]></FormulaPowerQuery>"#.to_vec();
    let mut overrides = HashMap::new();
    overrides.insert(
        POWER_QUERY_PART.to_string(),
        PartOverride::Add(added_xml.clone()),
    );

    let patches = WorkbookCellPatches::default();
    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
        Cursor::new(input.clone()),
        &mut out,
        &patches,
        &overrides,
    )?;
    let out_bytes = out.into_inner();

    let out_pkg = XlsxPackage::from_bytes(&out_bytes)?;
    assert_eq!(
        out_pkg.part(POWER_QUERY_PART),
        Some(added_xml.as_slice()),
        "expected power-query.xml to be added"
    );

    assert_parts_preserved_except(&input, &out_bytes, &[POWER_QUERY_PART]);
    Ok(())
}

#[test]
fn streaming_part_override_recalc_policy_can_clear_cached_values_for_formula_patches(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/rt_simple.xlsx");
    let bytes = fs::read(&fixture_path)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("C1")?,
        CellPatch::set_value_with_formula(CellValue::Number(2.0), "=1+1"),
    );

    let added_xml = br#"<FormulaPowerQuery version="1"><![CDATA[{"queries":[{"id":"added"}]}]]></FormulaPowerQuery>"#.to_vec();
    let mut overrides = HashMap::new();
    overrides.insert(POWER_QUERY_PART.to_string(), PartOverride::Add(added_xml.clone()));

    // With the recalc policy flag enabled, the formula patch should omit cached `<v>` values.
    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides_and_recalc_policy(
        Cursor::new(bytes.clone()),
        &mut out,
        &patches,
        &overrides,
        RecalcPolicy {
            clear_cached_values_on_formula_change: true,
            ..Default::default()
        },
    )?;
    let out_bytes = out.into_inner();

    // Ensure the override was applied.
    let out_pkg = XlsxPackage::from_bytes(&out_bytes)?;
    assert_eq!(
        out_pkg.part(POWER_QUERY_PART),
        Some(added_xml.as_slice()),
        "expected power-query.xml to be added"
    );

    // Ensure the cached value was cleared for the patched formula cell.
    let mut archive = ZipArchive::new(Cursor::new(&out_bytes))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;
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

    // Control: the existing entrypoint still writes cached values (default policy).
    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
        Cursor::new(bytes),
        &mut out,
        &patches,
        &overrides,
    )?;
    let out_bytes = out.into_inner();
    let mut archive = ZipArchive::new(Cursor::new(&out_bytes))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;
    let xml_doc = roxmltree::Document::parse(&sheet_xml)?;
    let patched_cell = xml_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("C1"))
        .expect("patched cell C1 should exist");
    assert!(
        patched_cell
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "v"),
        "default part-overrides streaming patcher should preserve cached <v> values"
    );

    Ok(())
}
