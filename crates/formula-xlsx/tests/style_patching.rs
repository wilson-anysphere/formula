use std::fs;
use std::io::{Cursor, Read};

use formula_model::{CellRef, CellValue, Style};
use formula_xlsx::{
    load_from_bytes, patch_xlsx_streaming, CellPatch, WorkbookCellPatches, WorksheetCellPatch,
    XlsxPackage,
};
use zip::ZipArchive;

fn load_fixture() -> Vec<u8> {
    fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/styles/varied_styles.xlsx"
    ))
    .expect("fixture exists")
}

fn sheet1_xml(bytes: &[u8]) -> String {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("valid zip");
    let mut xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet1.xml exists")
        .read_to_string(&mut xml)
        .expect("read sheet xml");
    xml
}

fn cell_s_attr(sheet_xml: &str, a1: &str) -> Option<String> {
    let doc = roxmltree::Document::parse(sheet_xml).expect("valid xml");
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    doc.descendants()
        .find(|n| n.has_tag_name((ns, "c")) && n.attribute("r") == Some(a1))
        .and_then(|n| n.attribute("s"))
        .map(str::to_string)
}

fn cell_xfs_count(styles_xml: &str) -> u32 {
    let doc = roxmltree::Document::parse(styles_xml).expect("valid xml");
    doc.descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "cellXfs")
        .and_then(|n| n.attribute("count"))
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(0)
}

#[test]
fn patch_preserves_existing_style_when_xf_index_none() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = load_fixture();
    let orig = load_from_bytes(&bytes)?;
    let sheet_id = orig.workbook.sheets[0].id;
    let sheet = orig.workbook.sheet(sheet_id).expect("sheet1 exists");

    let a1 = CellRef::from_a1("A1")?;
    let orig_style = sheet
        .cell(a1)
        .map(|c| c.style_id)
        .expect("A1 should exist in fixture");

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        a1,
        CellValue::String("Updated".to_string()),
        None,
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;

    let patched = load_from_bytes(out.get_ref())?;
    let sheet_id = patched.workbook.sheets[0].id;
    let sheet = patched
        .workbook
        .sheet(sheet_id)
        .expect("sheet1 exists in patched workbook");
    let cell = sheet.cell(a1).expect("A1 should exist after patch");

    assert_eq!(cell.value, CellValue::String("Updated".to_string()));
    assert_eq!(
        cell.style_id, orig_style,
        "style should be preserved by default"
    );
    Ok(())
}

#[test]
fn builtin_number_formats_use_builtin_num_fmt_ids_on_write() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = load_fixture();

    // Use the model loader to get a StyleTable that already contains the workbook's styles.
    let doc = load_from_bytes(&bytes)?;
    let mut style_table = doc.workbook.styles.clone();

    // Built-in id 1 is `0`.
    let style_id = style_table.intern(Style {
        number_format: Some("0".to_string()),
        ..Default::default()
    });

    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let a1 = CellRef::from_a1("A1")?;
    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        a1,
        CellPatch::set_value(CellValue::Number(123.0)).with_style_id(style_id),
    );

    pkg.apply_cell_patches_with_styles(&patches, &style_table)?;

    let styles_xml = std::str::from_utf8(pkg.part("xl/styles.xml").expect("styles.xml exists"))?;

    // We should not introduce a custom <numFmt ... formatCode="0"> entry.
    assert!(
        !styles_xml.contains(r#"formatCode="0""#),
        "styles.xml should not contain a custom numFmt for built-in '0' format:\n{styles_xml}"
    );

    // The new XF should reference the built-in numFmtId.
    assert!(
        styles_xml.contains(r#"numFmtId="1""#),
        "expected built-in numFmtId=1 to appear in styles.xml:\n{styles_xml}"
    );

    Ok(())
}

#[test]
fn patch_can_override_style_xf_index() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = load_fixture();
    let a1 = CellRef::from_a1("A1")?;

    // Remove the style by setting `xf_index = Some(0)`.
    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        a1,
        CellValue::String("NoStyle".to_string()),
        None,
    )
    .with_xf_index(Some(0));

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes.clone()), &mut out, &[patch])?;
    let sheet_xml = sheet1_xml(out.get_ref());
    assert_eq!(cell_s_attr(&sheet_xml, "A1"), None);

    // Overwrite the style to an existing non-zero `xf` index.
    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        a1,
        CellValue::String("StyledAgain".to_string()),
        None,
    )
    .with_xf_index(Some(2));

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch])?;
    let sheet_xml = sheet1_xml(out.get_ref());
    assert_eq!(cell_s_attr(&sheet_xml, "A1"), Some("2".to_string()));

    Ok(())
}

#[test]
fn package_apply_cell_patches_updates_styles_xml_when_new_style_added(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = load_fixture();

    // Use the model loader to get a StyleTable that already contains the workbook's styles.
    let doc = load_from_bytes(&bytes)?;
    let mut style_table = doc.workbook.styles.clone();

    let new_style_id = style_table.intern(Style {
        // A unique number format string to force a new xf record.
        number_format: Some("0.0000000000000000\"STYLE_TEST\"".to_string()),
        ..Default::default()
    });

    let mut pkg = XlsxPackage::from_bytes(&bytes)?;
    let before_styles = std::str::from_utf8(pkg.part("xl/styles.xml").expect("styles.xml exists"))?;
    let before_count = cell_xfs_count(before_styles);

    let a1 = CellRef::from_a1("A1")?;
    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        a1,
        CellPatch::set_value(CellValue::String("Styled".to_string())).with_style_id(new_style_id),
    );

    pkg.apply_cell_patches_with_styles(&patches, &style_table)?;

    let after_styles = std::str::from_utf8(pkg.part("xl/styles.xml").expect("styles.xml exists"))?;
    let after_count = cell_xfs_count(after_styles);
    assert_eq!(
        after_count,
        before_count + 1,
        "expected a new xf record to be appended"
    );

    let expected_xf = before_count;
    let sheet_xml = std::str::from_utf8(
        pkg.part("xl/worksheets/sheet1.xml")
            .expect("sheet1.xml exists"),
    )?;
    assert_eq!(cell_s_attr(sheet_xml, "A1"), Some(expected_xf.to_string()));

    Ok(())
}
