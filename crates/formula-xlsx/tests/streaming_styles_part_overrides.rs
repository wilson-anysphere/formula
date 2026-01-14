use std::collections::{BTreeSet, HashMap};
use std::fs;
use std::io::Cursor;
use std::path::Path;

use formula_model::{CellRef, CellValue, Style, StyleTable};
use formula_xlsx::{
    load_from_bytes, patch_xlsx_streaming_workbook_cell_patches_with_styles_and_part_overrides,
    CellPatch, PartOverride, WorkbookCellPatches, XlsxPackage,
};
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

fn part_name_set(bytes: &[u8]) -> BTreeSet<String> {
    XlsxPackage::from_bytes(bytes)
        .expect("parse package")
        .part_names()
        .map(str::to_string)
        .collect()
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

fn cell_xfs_count(styles_xml: &str) -> u32 {
    let doc = roxmltree::Document::parse(styles_xml).expect("valid xml");
    doc.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "cellXfs")
        .and_then(|n| n.attribute("count"))
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(0)
}

fn cell_s_attr(sheet_xml: &str, a1: &str) -> Option<String> {
    let doc = roxmltree::Document::parse(sheet_xml).expect("valid xml");
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    doc.descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some(a1))
        .and_then(|n| n.attribute("s"))
        .map(str::to_string)
}

#[test]
fn streaming_styles_and_part_overrides_can_patch_style_id_and_replace_power_query_in_one_pass(
) -> Result<(), Box<dyn std::error::Error>> {
    let base = fixture_basic_xlsx_bytes();
    let initial_xml = br#"<FormulaPowerQuery version="1"><![CDATA[{"queries":[{"id":"q1"}]}]]></FormulaPowerQuery>"#.to_vec();
    let input = inject_power_query_part(&base, initial_xml.clone());

    // Use the model loader to get a StyleTable that already contains the workbook's styles.
    let doc = load_from_bytes(&input)?;
    let mut style_table = doc.workbook.styles.clone();

    let new_style_id = style_table.intern(Style {
        // A unique number format string to force a new xf record.
        number_format: Some("0.0000000000000000\"STYLE_TEST\"".to_string()),
        ..Default::default()
    });

    let input_pkg = XlsxPackage::from_bytes(&input)?;
    let before_styles =
        std::str::from_utf8(input_pkg.part("xl/styles.xml").expect("styles.xml exists"))?;
    let before_count = cell_xfs_count(before_styles);

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A1")?,
        CellPatch::set_value(CellValue::Number(123.0)).with_style_id(new_style_id),
    );

    let updated_xml = br#"<FormulaPowerQuery version="1"><![CDATA[{"queries":[{"id":"q2"}]}]]></FormulaPowerQuery>"#.to_vec();
    let mut overrides = HashMap::new();
    overrides.insert(
        POWER_QUERY_PART.to_string(),
        PartOverride::Replace(updated_xml.clone()),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_styles_and_part_overrides(
        Cursor::new(input.clone()),
        &mut out,
        &patches,
        &style_table,
        &overrides,
    )?;
    let out_bytes = out.into_inner();

    // Ensure the override was applied.
    let out_pkg = XlsxPackage::from_bytes(&out_bytes)?;
    assert_eq!(
        out_pkg.part(POWER_QUERY_PART),
        Some(updated_xml.as_slice()),
        "expected power-query.xml to be replaced"
    );

    // Ensure styles.xml grew deterministically and the patched cell references the new xf index.
    let out_styles =
        std::str::from_utf8(out_pkg.part("xl/styles.xml").expect("styles.xml exists"))?;
    let after_count = cell_xfs_count(out_styles);
    assert_eq!(
        after_count,
        before_count + 1,
        "expected a new xf record to be appended"
    );
    assert!(
        out_styles.contains("STYLE_TEST"),
        "expected styles.xml to contain STYLE_TEST custom format:\n{out_styles}"
    );
    let sheet_xml = std::str::from_utf8(
        out_pkg
            .part("xl/worksheets/sheet1.xml")
            .expect("sheet1.xml exists"),
    )?;
    assert_eq!(cell_s_attr(sheet_xml, "A1"), Some(before_count.to_string()));

    // Replacing an existing part should not change the set of ZIP entries.
    assert_eq!(part_name_set(&input), part_name_set(&out_bytes));

    // Ensure everything else is preserved.
    assert_parts_preserved_except(
        &input,
        &out_bytes,
        &["xl/styles.xml", "xl/worksheets/sheet1.xml", POWER_QUERY_PART],
    );

    Ok(())
}

#[test]
fn streaming_styles_and_part_overrides_appends_added_parts_in_lexicographic_order(
) -> Result<(), Box<dyn std::error::Error>> {
    let input = fixture_basic_xlsx_bytes();
    let style_table = StyleTable::default();
    let patches = WorkbookCellPatches::default();

    let mut overrides = HashMap::new();
    overrides.insert(
        "xl/formula/aaa.xml".to_string(),
        PartOverride::Add(b"<A/>".to_vec()),
    );
    overrides.insert(
        POWER_QUERY_PART.to_string(),
        PartOverride::Add(b"<PQ/>".to_vec()),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_styles_and_part_overrides(
        Cursor::new(input.clone()),
        &mut out,
        &patches,
        &style_table,
        &overrides,
    )?;
    let out_bytes = out.into_inner();

    let out_pkg = XlsxPackage::from_bytes(&out_bytes)?;
    assert_eq!(out_pkg.part("xl/formula/aaa.xml"), Some(b"<A/>".as_slice()));
    assert_eq!(out_pkg.part(POWER_QUERY_PART), Some(b"<PQ/>".as_slice()));

    // Adding should introduce only the new parts.
    let mut expected_names = part_name_set(&input);
    expected_names.insert("xl/formula/aaa.xml".to_string());
    expected_names.insert(POWER_QUERY_PART.to_string());
    assert_eq!(expected_names, part_name_set(&out_bytes));

    // Added parts are appended deterministically at the end of the ZIP.
    let mut archive = ZipArchive::new(Cursor::new(&out_bytes))?;
    let mut non_dir_names = Vec::new();
    for i in 0..archive.len() {
        let file = archive.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        non_dir_names.push(file.name().strip_prefix('/').unwrap_or(file.name()).to_string());
    }
    let n = non_dir_names.len();
    assert_eq!(non_dir_names[n - 2].as_str(), "xl/formula/aaa.xml");
    assert_eq!(non_dir_names[n - 1].as_str(), POWER_QUERY_PART);

    assert_parts_preserved_except(&input, &out_bytes, &["xl/formula/aaa.xml", POWER_QUERY_PART]);
    Ok(())
}

#[test]
fn streaming_styles_and_part_overrides_skips_removed_parts(
) -> Result<(), Box<dyn std::error::Error>> {
    let base = fixture_basic_xlsx_bytes();
    let initial_xml =
        br#"<FormulaPowerQuery version="1"><![CDATA[{"queries":[{"id":"q1"}]}]]></FormulaPowerQuery>"#
            .to_vec();
    let input = inject_power_query_part(&base, initial_xml);

    let style_table = StyleTable::default();
    let patches = WorkbookCellPatches::default();

    let mut overrides = HashMap::new();
    overrides.insert(POWER_QUERY_PART.to_string(), PartOverride::Remove);

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches_with_styles_and_part_overrides(
        Cursor::new(input.clone()),
        &mut out,
        &patches,
        &style_table,
        &overrides,
    )?;
    let out_bytes = out.into_inner();

    let out_pkg = XlsxPackage::from_bytes(&out_bytes)?;
    assert!(
        out_pkg.part(POWER_QUERY_PART).is_none(),
        "expected power-query.xml to be removed"
    );

    let mut expected_names = part_name_set(&input);
    expected_names.remove(POWER_QUERY_PART);
    assert_eq!(expected_names, part_name_set(&out_bytes));

    assert_parts_preserved_except(&input, &out_bytes, &[POWER_QUERY_PART]);
    Ok(())
}
