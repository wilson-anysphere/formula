use std::io::{Cursor, Read};

use base64::Engine;
use formula_model::{CellRef, CellValue};
use formula_xlsx::{
    load_from_bytes, CellPatch as WorkbookCellPatch, PackageCellPatch, WorkbookCellPatches,
    XlsxPackage,
};
use rust_xlsxwriter::{Image, Workbook};
use zip::ZipArchive;

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn zip_part_names(zip_bytes: &[u8]) -> Vec<String> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    (0..archive.len())
        .filter_map(|idx| archive.by_index(idx).ok().map(|f| f.name().to_string()))
        .collect()
}

fn worksheet_cell_vm(sheet_xml: &str, cell_ref: &str) -> Option<String> {
    let doc = roxmltree::Document::parse(sheet_xml).ok()?;
    let cell = doc.descendants().find(|n| {
        n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some(cell_ref)
    })?;
    cell.attribute("vm").map(|s| s.to_string())
}

fn build_xlsx_with_in_cell_image() -> Vec<u8> {
    // 1x1 transparent PNG.
    let png_bytes = base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/58HAQUBAO3+2NoAAAAASUVORK5CYII=")
        .expect("valid base64 png");

    let mut workbook = Workbook::new();
    {
        let worksheet = workbook.add_worksheet();

        let image = Image::new_from_buffer(&png_bytes).expect("create image");
        worksheet
            .embed_image(0, 0, &image)
            .expect("embed image in cell");
    }

    workbook.save_to_buffer().expect("save workbook")
}

#[test]
fn patching_other_cells_preserves_in_cell_image_vm_and_richdata_parts(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_xlsx_with_in_cell_image();

    // Sanity check: the embedded image cell should have a vm pointer.
    let sheet_xml_bytes = zip_part(&bytes, "xl/worksheets/sheet1.xml");
    let sheet_xml = std::str::from_utf8(&sheet_xml_bytes)?;
    assert_eq!(
        worksheet_cell_vm(sheet_xml, "A1").as_deref(),
        Some("1"),
        "expected embedded image placeholder cell to contain vm=\"1\""
    );

    let pkg = XlsxPackage::from_bytes(&bytes)?;
    let patched_bytes = pkg.apply_cell_patches_to_bytes(&[PackageCellPatch::for_sheet_name(
        "Sheet1",
        CellRef::from_a1("B1")?,
        CellValue::Number(123.0),
        None,
    )])?;

    let patched_sheet_xml_bytes = zip_part(&patched_bytes, "xl/worksheets/sheet1.xml");
    let patched_sheet_xml = std::str::from_utf8(&patched_sheet_xml_bytes)?;
    assert_eq!(
        worksheet_cell_vm(patched_sheet_xml, "A1").as_deref(),
        Some("1"),
        "patching an unrelated cell should not drop vm from the embedded-image placeholder"
    );

    let part_names = zip_part_names(&patched_bytes);
    assert!(
        part_names.iter().any(|p| p == "xl/metadata.xml"),
        "expected xl/metadata.xml to be preserved"
    );
    assert!(
        part_names.iter().any(|p| p.starts_with("xl/richData/")),
        "expected xl/richData/* parts to be preserved"
    );

    Ok(())
}

#[test]
fn patching_in_cell_image_cell_to_normal_value_drops_vm(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_xlsx_with_in_cell_image();
    let pkg = XlsxPackage::from_bytes(&bytes)?;

    let patched_bytes = pkg.apply_cell_patches_to_bytes(&[PackageCellPatch::for_sheet_name(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellValue::Number(42.0),
        None,
    )])?;

    let patched_sheet_xml_bytes = zip_part(&patched_bytes, "xl/worksheets/sheet1.xml");
    let patched_sheet_xml = std::str::from_utf8(&patched_sheet_xml_bytes)?;
    let doc = roxmltree::Document::parse(patched_sheet_xml)?;
    let cell_a1 = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("A1 cell exists");
    assert_eq!(
        cell_a1.attribute("vm"),
        None,
        "vm must be dropped when patching away from the embedded-image placeholder semantics"
    );

    Ok(())
}

#[test]
fn apply_cell_patches_drops_vm_when_updating_in_cell_image_cell(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_xlsx_with_in_cell_image();
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        WorkbookCellPatch::set_value(CellValue::Number(1.0)),
    );
    pkg.apply_cell_patches(&patches)?;

    let sheet_xml = std::str::from_utf8(
        pkg.part("xl/worksheets/sheet1.xml")
            .expect("sheet1.xml exists"),
    )?;
    let doc = roxmltree::Document::parse(sheet_xml)?;
    let cell_a1 = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("A1 exists");
    assert_eq!(
        cell_a1.attribute("vm"),
        None,
        "vm must be dropped when updating away from the embedded-image placeholder semantics"
    );

    Ok(())
}

#[test]
fn xlsx_document_roundtrip_preserves_vm_for_in_cell_image_placeholder(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_xlsx_with_in_cell_image();
    let doc = load_from_bytes(&bytes)?;
    let saved = doc.save_to_vec()?;

    let sheet_xml_bytes = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let sheet_xml = std::str::from_utf8(&sheet_xml_bytes)?;
    assert_eq!(
        worksheet_cell_vm(sheet_xml, "A1").as_deref(),
        Some("1"),
        "round-tripping should preserve vm for the embedded-image placeholder"
    );

    Ok(())
}

#[test]
fn xlsx_document_editing_in_cell_image_cell_drops_vm(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_xlsx_with_in_cell_image();
    let mut doc = load_from_bytes(&bytes)?;

    let sheet_id = doc.workbook.sheets[0].id;
    doc.workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .set_value(CellRef::from_a1("A1")?, CellValue::Number(7.0));

    let saved = doc.save_to_vec()?;
    let sheet_xml_bytes = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let sheet_xml = std::str::from_utf8(&sheet_xml_bytes)?;
    let parsed = roxmltree::Document::parse(sheet_xml)?;
    let cell_a1 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("A1 exists");
    assert_eq!(
        cell_a1.attribute("vm"),
        None,
        "vm must be dropped when editing away from the embedded-image placeholder semantics"
    );

    Ok(())
}
