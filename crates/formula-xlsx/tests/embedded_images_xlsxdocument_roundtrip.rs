use base64::Engine;
use formula_model::{CellRef, CellValue};
use formula_xlsx::{load_from_bytes, XlsxPackage};
use rust_xlsxwriter::{Color, Format, Image, Workbook};

fn decode_png_bytes() -> Vec<u8> {
    // 1x1 transparent PNG.
    base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/58HAQUBAO3+2NoAAAAASUVORK5CYII=")
        .expect("valid base64 PNG")
}

fn assert_embedded_image_parts_present(pkg: &XlsxPackage) {
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
            "expected embedded image part {part} to be present"
        );
    }
}

fn assert_embedded_image_parts_equal(before: &XlsxPackage, after: &XlsxPackage) {
    for part in [
        "xl/metadata.xml",
        "xl/richData/rdrichvalue.xml",
        "xl/richData/rdrichvaluestructure.xml",
        "xl/richData/rdRichValueTypes.xml",
        "xl/richData/richValueRel.xml",
        "xl/richData/_rels/richValueRel.xml.rels",
        "xl/media/image1.png",
    ] {
        assert_eq!(
            before.part(part),
            after.part(part),
            "expected embedded image part {part} to be byte-for-byte preserved"
        );
    }
}

fn assert_sheet_cell_is_embedded_image(sheet_xml: &str, cell_a1: &str) {
    let doc = roxmltree::Document::parse(sheet_xml).expect("parse sheet XML");
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some(cell_a1))
        .unwrap_or_else(|| panic!("expected {cell_a1} cell to exist in worksheet xml"));

    assert!(
        cell.attribute("vm").is_some(),
        "expected {cell_a1} cell to have vm attribute, got: {sheet_xml}"
    );
    assert!(
        cell.attribute("s").is_some(),
        "expected {cell_a1} cell to have s attribute, got: {sheet_xml}"
    );

    let value = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(
        value,
        "#VALUE!",
        "expected embedded image cell {cell_a1} to have #VALUE! cached value"
    );
}

#[test]
fn xlsxdocument_roundtrip_preserves_embedded_images_in_cells() {
    // Create a temporary PNG file for rust_xlsxwriter to embed.
    let dir = tempfile::tempdir().expect("create temp dir");
    let png_path = dir.path().join("cell-image.png");
    std::fs::write(&png_path, decode_png_bytes()).expect("write png bytes");

    let image = Image::new(&png_path).expect("load image");

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Ensure we have at least one normal cell and one embedded image cell.
    worksheet.write_string(0, 0, "hello").expect("write A1");
    let format = Format::new()
        .set_background_color(Color::Yellow)
        .set_bold();
    worksheet
        .embed_image_with_format(1, 1, &image, &format)
        .expect("embed image in B2");

    let input = workbook.save_to_buffer().expect("write xlsx");

    // Sanity-check the generated workbook: verify the embedded image parts and the sheet cell.
    let input_pkg = XlsxPackage::from_bytes(&input).expect("open generated xlsx");
    assert_embedded_image_parts_present(&input_pkg);
    let input_sheet_xml =
        std::str::from_utf8(input_pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert_sheet_cell_is_embedded_image(input_sheet_xml, "B2");

    // Load into the higher-fidelity XlsxDocument representation.
    let mut doc = load_from_bytes(&input).expect("load xlsx into XlsxDocument");
    let sheet_id = doc.workbook.sheets[0].id;

    // Make a minimal model edit unrelated to the embedded image.
    doc.set_cell_value(
        sheet_id,
        CellRef::from_a1("C3").unwrap(),
        CellValue::String("world".to_string()),
    );

    let output = doc.save_to_vec().expect("save XlsxDocument");

    // Reload as a package and assert that the embedded image richData/metadata/media parts were
    // preserved.
    let output_pkg = XlsxPackage::from_bytes(&output).expect("open saved xlsx");
    assert_embedded_image_parts_present(&output_pkg);
    assert_embedded_image_parts_equal(&input_pkg, &output_pkg);

    // Ensure the embedded image cell still has `vm="..."` and the `#VALUE!` cached value.
    let output_sheet_xml =
        std::str::from_utf8(output_pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert_sheet_cell_is_embedded_image(output_sheet_xml, "B2");
}
