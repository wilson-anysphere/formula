use std::io::Cursor;

use base64::Engine;
use formula_model::{CellRef, CellValue};
use formula_xlsx::{
    patch_xlsx_streaming_workbook_cell_patches, CellPatch, WorkbookCellPatches, XlsxPackage,
};
use rust_xlsxwriter::{Format, Image, Workbook};

fn assert_cell_vm_error(sheet_xml: &str, a1: &str, vm: &str, error_value: &str) {
    let doc = roxmltree::Document::parse(sheet_xml).expect("parse worksheet xml");
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some(a1))
        .unwrap_or_else(|| panic!("expected cell {a1} to exist"));

    assert_eq!(
        cell.attribute("vm"),
        Some(vm),
        "expected cell {a1} to have vm={vm}"
    );
    assert_eq!(cell.attribute("t"), Some("e"), "expected cell {a1} to be an error cell");
    let v = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(v, error_value, "expected cell {a1} to preserve error value");
}

fn assert_rich_data_parts_present(pkg: &XlsxPackage) {
    for part in [
        "xl/metadata.xml",
        "xl/richData/rdrichvalue.xml",
        "xl/richData/rdrichvaluestructure.xml",
        "xl/richData/rdRichValueTypes.xml",
        "xl/richData/richValueRel.xml",
        "xl/richData/_rels/richValueRel.xml.rels",
        "xl/media/image1.png",
    ] {
        assert!(
            pkg.part(part).is_some(),
            "expected embedded-image richData part {part} to be present"
        );
    }
}

#[test]
fn streaming_embedded_images_vm_preserved_on_row_rewrite() -> Result<(), Box<dyn std::error::Error>>
{
    // 1x1 transparent PNG.
    let png_bytes = base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/58HAQUBAO3+2NoAAAAASUVORK5CYII=")
        .expect("valid base64 png");

    // Generate a workbook with an embedded image in B2 (row=1,col=1). Excel stores this as an
    // error cell with a `vm="..."` attribute pointing at richData parts.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let image = Image::new_from_buffer(&png_bytes)?;
    let format = Format::new();
    worksheet.embed_image_with_format(1, 1, &image, &format)?;

    let input_bytes = workbook.save_to_buffer()?;

    // Sanity-check that rust_xlsxwriter emitted the vm cell + richData parts.
    let input_pkg = XlsxPackage::from_bytes(&input_bytes)?;
    assert_rich_data_parts_present(&input_pkg);
    let input_sheet_xml =
        std::str::from_utf8(input_pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert_cell_vm_error(input_sheet_xml, "B2", "1", "#VALUE!");

    // Patch another cell in the same row (A2), which forces the streaming patcher to rewrite row
    // 2 while leaving B2 unpatched.
    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A2")?,
        CellPatch::set_value(CellValue::Number(123.0)),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(input_bytes), &mut out, &patches)?;

    let out_pkg = XlsxPackage::from_bytes(out.get_ref())?;
    assert_rich_data_parts_present(&out_pkg);
    let out_sheet_xml =
        std::str::from_utf8(out_pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();

    // Ensure the embedded-image cell is still present with its vm attribute + cached error value.
    assert_cell_vm_error(out_sheet_xml, "B2", "1", "#VALUE!");

    // Ensure A2 was patched.
    let doc = roxmltree::Document::parse(out_sheet_xml)?;
    let cell_a2 = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A2"))
        .expect("expected patched cell A2 to exist");
    let v = cell_a2
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(v, "123");

    Ok(())
}

