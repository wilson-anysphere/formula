use std::io::Write as _;

use formula_model::{CellRef, CellValue};
use formula_xlsx::{CellPatch, WorkbookCellPatches, XlsxPackage};
use rust_xlsxwriter::{Format, Image, Workbook};
use tempfile::NamedTempFile;

// A valid 1x1 PNG (grayscale + alpha), base64-decoded from:
// iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/wwAAgMBgE+0bQAAAABJRU5ErkJggg==
const TINY_PNG_BYTES: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
    0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x04, 0x00, 0x00,
    0x00, 0xB5, 0x1C, 0x0C, 0x02, 0x00, 0x00, 0x00, 0x0B, 0x49, 0x44, 0x41, 0x54, 0x78,
    0xDA, 0x63, 0xFC, 0xFF, 0x0C, 0x00, 0x02, 0x03, 0x01, 0x80, 0x4F, 0xB4, 0x6D, 0x00,
    0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
];

fn assert_part_present(pkg: &XlsxPackage, part: &str) {
    assert!(
        pkg.part(part).is_some(),
        "expected part {part} to exist. Available parts: {:?}",
        pkg.part_names().collect::<Vec<_>>()
    );
}

#[test]
fn embedded_images_place_in_cell_roundtrip_preserves_metadata_and_richdata_parts() {
    // Generate an XLSX with an embedded (place-in-cell) image using rust_xlsxwriter.
    // This should create the xl/metadata.xml + xl/richData/* parts plus the image payload.
    let mut tmp_png = NamedTempFile::new().expect("create temp png");
    tmp_png
        .write_all(TINY_PNG_BYTES)
        .expect("write png bytes");
    tmp_png.flush().expect("flush png");

    let image = Image::new(tmp_png.path()).expect("load png as rust_xlsxwriter::Image");

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Use embed_image_with_format to ensure the cell carries an `s="..."` style attribute.
    let format = Format::new().set_bold();
    worksheet
        .embed_image_with_format(0, 0, &image, &format)
        .expect("embed image");

    let input_bytes = workbook.save_to_buffer().expect("save workbook");

    // Sanity-check that the source package actually contains the expected parts.
    let mut pkg = XlsxPackage::from_bytes(&input_bytes).expect("parse generated xlsx");
    for part in [
        "xl/metadata.xml",
        "xl/richData/rdrichvalue.xml",
        "xl/richData/rdrichvaluestructure.xml",
        "xl/richData/rdRichValueTypes.xml",
        "xl/richData/richValueRel.xml",
        "xl/richData/_rels/richValueRel.xml.rels",
        "xl/media/image1.png",
    ] {
        assert_part_present(&pkg, part);
    }

    // Apply a benign patch to a different cell to exercise the streaming patch pipeline.
    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("B2").expect("valid cell ref"),
        CellPatch::set_value(CellValue::Number(123.0)),
    );
    pkg.apply_cell_patches(&patches)
        .expect("apply patches to pkg");

    let out_bytes = pkg.write_to_bytes().expect("write patched xlsx");
    let pkg2 = XlsxPackage::from_bytes(&out_bytes).expect("parse patched xlsx");

    // The embedded-image OOXML parts should still exist after patching.
    for part in [
        "xl/metadata.xml",
        "xl/richData/rdrichvalue.xml",
        "xl/richData/rdrichvaluestructure.xml",
        "xl/richData/rdRichValueTypes.xml",
        "xl/richData/richValueRel.xml",
        "xl/richData/_rels/richValueRel.xml.rels",
        "xl/media/image1.png",
    ] {
        assert_part_present(&pkg2, part);
    }

    // The embedded-image cell should still look like an error cell with a value metadata index:
    // `<c r="A1" t="e" vm="..."><v>#VALUE!</v></c>`.
    let sheet_xml =
        std::str::from_utf8(pkg2.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let doc = roxmltree::Document::parse(sheet_xml).expect("parse sheet xml");
    let a1 = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected embedded image cell at A1");
    assert_eq!(a1.attribute("t"), Some("e"), "expected A1 to be an error cell");
    assert!(
        a1.attribute("vm").is_some(),
        "expected embedded image cell to have vm attribute, got: {sheet_xml}"
    );
    assert!(
        a1.attribute("s").is_some(),
        "expected embedded image cell to have an s (style) attribute, got: {sheet_xml}"
    );
    let v = a1
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .expect("expected cached <v> child for embedded image cell");
    assert_eq!(
        v.text(),
        Some("#VALUE!"),
        "expected embedded image cell cached value to be #VALUE!, got: {sheet_xml}"
    );

    // Optional: verify workbook relationships still include the richData + sheetMetadata rels.
    let workbook_rels =
        std::str::from_utf8(pkg2.part("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
    assert!(
        workbook_rels.contains("relationships/sheetMetadata"),
        "expected workbook.xml.rels to reference sheetMetadata, got: {workbook_rels}"
    );
    assert!(
        workbook_rels.contains("schemas.microsoft.com/office/2022/10/relationships/richValueRel"),
        "expected workbook.xml.rels to reference richValueRel, got: {workbook_rels}"
    );
    assert!(
        workbook_rels.contains("schemas.microsoft.com/office/2017/06/relationships/rdRichValue"),
        "expected workbook.xml.rels to reference rdRichValue relationships, got: {workbook_rels}"
    );
}

