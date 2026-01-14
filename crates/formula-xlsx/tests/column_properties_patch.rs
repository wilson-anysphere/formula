use std::collections::BTreeMap;
use std::io::Cursor;
use std::path::Path;

use formula_model::ColProperties;

use formula_xlsx::{
    patch_xlsx_streaming_workbook_cell_patches, read_workbook_model_from_bytes, WorkbookCellPatches,
    XlsxPackage,
};

fn build_basic_col_patches() -> WorkbookCellPatches {
    let mut cols: BTreeMap<u32, ColProperties> = BTreeMap::new();
    cols.insert(
        0,
        ColProperties {
            width: Some(15.0),
            ..Default::default()
        },
    );
    cols.insert(
        1,
        ColProperties {
            hidden: true,
            ..Default::default()
        },
    );

    let mut patches = WorkbookCellPatches::default();
    patches.sheet_mut("Sheet1").set_col_properties(cols);
    patches
}

fn count_elements(xml: &str, local_name: &str) -> usize {
    let doc = roxmltree::Document::parse(xml).expect("parse xml");
    doc.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == local_name)
        .count()
}

#[test]
fn apply_cell_patches_persists_column_width_and_hidden() -> Result<(), Box<dyn std::error::Error>> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/basic.xlsx");
    let bytes = std::fs::read(&fixture)?;

    let mut pkg = XlsxPackage::from_bytes(&bytes)?;
    let patches = build_basic_col_patches();
    pkg.apply_cell_patches(&patches)?;

    let out = pkg.write_to_bytes()?;
    let workbook = read_workbook_model_from_bytes(&out)?;
    let sheet = &workbook.sheets[0];

    assert_eq!(sheet.name, "Sheet1");
    assert_eq!(sheet.col_properties(0).and_then(|p| p.width), Some(15.0));
    assert!(sheet.col_properties(1).is_some_and(|p| p.hidden));
    assert!(sheet.col_properties(2).is_none(), "expected sparse col props");

    // Column-only patches should not introduce a `<dimension>` element.
    let sheet_xml = std::str::from_utf8(
        pkg.part("xl/worksheets/sheet1.xml")
            .expect("sheet1.xml exists"),
    )?;
    assert!(!sheet_xml.contains("<dimension"), "unexpected dimension: {sheet_xml}");

    Ok(())
}

#[test]
fn apply_cell_patches_preserves_existing_col_attributes() -> Result<(), Box<dyn std::error::Error>> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/basic.xlsx");
    let bytes = std::fs::read(&fixture)?;

    let mut pkg = XlsxPackage::from_bytes(&bytes)?;
    let sheet_xml = std::str::from_utf8(
        pkg.part("xl/worksheets/sheet1.xml")
            .expect("sheet1.xml exists"),
    )?;

    // Inject an existing `<cols>` section with an unsupported attribute. Column-metadata patches
    // should update width/hidden while preserving unrelated `<col>` attributes.
    let injected = sheet_xml.replace(
        "<sheetData>",
        r#"<cols><col min="1" max="1" outlineLevel="2"/></cols><sheetData>"#,
    );
    pkg.set_part("xl/worksheets/sheet1.xml", injected.into_bytes());

    let patches = build_basic_col_patches();
    pkg.apply_cell_patches(&patches)?;

    let updated_xml = std::str::from_utf8(
        pkg.part("xl/worksheets/sheet1.xml")
            .expect("sheet1.xml exists"),
    )?;
    assert!(
        updated_xml.contains("outlineLevel=\"2\""),
        "expected outlineLevel to be preserved, got:\n{updated_xml}"
    );

    Ok(())
}

#[test]
fn streaming_patch_column_metadata_is_idempotent() -> Result<(), Box<dyn std::error::Error>> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/basic.xlsx");
    let bytes = std::fs::read(&fixture)?;
    let patches = build_basic_col_patches();

    // Apply the same patch twice via the streaming pipeline.
    let mut first = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes), &mut first, &patches)?;
    let first_bytes = first.into_inner();

    let mut second = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(first_bytes), &mut second, &patches)?;
    let second_bytes = second.into_inner();

    let pkg = XlsxPackage::from_bytes(&second_bytes)?;
    let sheet_xml = std::str::from_utf8(
        pkg.part("xl/worksheets/sheet1.xml")
            .expect("sheet1.xml exists"),
    )?;

    assert_eq!(count_elements(sheet_xml, "cols"), 1);
    assert_eq!(count_elements(sheet_xml, "col"), 2);

    Ok(())
}
