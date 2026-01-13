use base64::Engine as _;
use formula_model::CellRef;
use formula_xlsx::XlsxPackage;
use rust_xlsxwriter::{Image, Workbook};
use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

const ONE_BY_ONE_PNG_BASE64: &str =
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/w8AAgMBApZ9xO4AAAAASUVORK5CYII=";

fn one_by_one_png_bytes() -> Vec<u8> {
    base64::engine::general_purpose::STANDARD
        .decode(ONE_BY_ONE_PNG_BASE64)
        .expect("decode png base64")
}

fn write_temp_png(bytes: &[u8]) -> (tempfile::TempDir, std::path::PathBuf) {
    let dir = tempfile::tempdir().expect("tempdir");
    let path = dir.path().join("image.png");
    std::fs::write(&path, bytes).expect("write png");
    (dir, path)
}

fn build_workbook_with_embedded_image(
    alt_text: Option<&str>,
    include_dynamic_array: bool,
) -> Vec<u8> {
    let png = one_by_one_png_bytes();
    let (_dir, image_path) = write_temp_png(&png);

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    if include_dynamic_array {
        // Force Excel to emit XLDAPR metadata alongside rich value metadata.
        // `SEQUENCE()` is a dynamic array function.
        worksheet
            .write_dynamic_array_formula(0, 1, 2, 1, "=SEQUENCE(3)")
            .unwrap();
    }

    let mut image = Image::new(&image_path).expect("create image");
    if let Some(text) = alt_text {
        image = image.set_alt_text(text);
    }

    worksheet.embed_image(0, 0, &image).unwrap();
    workbook.save_to_buffer().unwrap()
}

fn build_minimal_vm_indexed_embedded_image_xlsx(png_bytes: &[u8]) -> Vec<u8> {
    // Some real-world workbooks omit `xl/metadata.xml` and the rich value tables but still encode
    // embedded-in-cell images by using the worksheet cell `vm=` value as a direct index into
    // `xl/richData/richValueRel.xml` relationship slots.
    //
    // This fixture is intentionally minimal: it includes only the parts required by
    // `XlsxPackage::worksheet_parts()` + `extract_embedded_cell_images()`.

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#;

    // Use vm="1" to cover a 1-based producer; the extractor should tolerate both 0-based and
    // 1-based slot indices when metadata mapping is missing.
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="1"/>
    </row>
  </sheetData>
</worksheet>"#;

    let rich_value_rel_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRel>"#;

    let rich_value_rel_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/richValueRel.xml", options)
        .unwrap();
    zip.write_all(rich_value_rel_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/_rels/richValueRel.xml.rels", options)
        .unwrap();
    zip.write_all(rich_value_rel_rels.as_bytes()).unwrap();

    zip.start_file("xl/media/image1.png", options).unwrap();
    zip.write_all(png_bytes).unwrap();

    zip.finish().unwrap().into_inner()
}

fn build_minimal_vm_indexed_embedded_images_xlsx_two(
    img1_bytes: &[u8],
    img2_bytes: &[u8],
) -> Vec<u8> {
    // Same as `build_minimal_vm_indexed_embedded_image_xlsx`, but with two image relationship slots
    // and worksheet `vm` values encoded as a 0-based index (vm="0", vm="1", ...).

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="0"/>
    </row>
    <row r="2">
      <c r="A2" vm="1"/>
    </row>
  </sheetData>
</worksheet>"#;

    let rich_value_rel_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
  <rel r:id="rId2"/>
</richValueRel>"#;

    let rich_value_rel_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/richValueRel.xml", options)
        .unwrap();
    zip.write_all(rich_value_rel_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/_rels/richValueRel.xml.rels", options)
        .unwrap();
    zip.write_all(rich_value_rel_rels.as_bytes()).unwrap();

    zip.start_file("xl/media/image1.png", options).unwrap();
    zip.write_all(img1_bytes).unwrap();
    zip.start_file("xl/media/image2.png", options).unwrap();
    zip.write_all(img2_bytes).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn extracts_single_embedded_image_in_cell() {
    let bytes = build_workbook_with_embedded_image(None, false);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read xlsx");

    let images = pkg
        .extract_embedded_cell_images()
        .expect("extract embedded cell images");
    assert_eq!(images.len(), 1);

    let key = ("xl/worksheets/sheet1.xml".to_string(), CellRef::new(0, 0));
    let cell_img = images.get(&key).expect("expected A1 embedded image");
    let stored = &cell_img.image_bytes;
    assert_eq!(
        stored,
        &one_by_one_png_bytes(),
        "expected extracted image bytes to match the inserted PNG"
    );
}

#[test]
fn extracts_embedded_image_even_without_metadata_xml() {
    let png = one_by_one_png_bytes();
    let bytes = build_minimal_vm_indexed_embedded_image_xlsx(&png);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read xlsx");
    assert!(
        pkg.part("xl/metadata.xml").is_none(),
        "fixture should omit xl/metadata.xml"
    );

    let images = pkg
        .extract_embedded_cell_images()
        .expect("extract embedded cell images");
    assert_eq!(images.len(), 1);

    let key = ("xl/worksheets/sheet1.xml".to_string(), CellRef::new(0, 0));
    let cell_img = images.get(&key).expect("expected A1 embedded image");
    assert_eq!(cell_img.image_part, "xl/media/image1.png");
    assert_eq!(cell_img.image_bytes, png);
    // Without `xl/metadata.xml` / `rdRichValue` parts we can't recover CalcOrigin. We default to
    // `0` (unknown).
    assert_eq!(cell_img.calc_origin, 0);
    assert!(cell_img.alt_text.is_none());
}

#[test]
fn extracts_embedded_images_without_metadata_when_vm_is_zero_based() {
    let img1 = b"img1".to_vec();
    let img2 = b"img2".to_vec();
    let bytes = build_minimal_vm_indexed_embedded_images_xlsx_two(&img1, &img2);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read xlsx");
    assert!(
        pkg.part("xl/metadata.xml").is_none(),
        "fixture should omit xl/metadata.xml"
    );

    let images = pkg
        .extract_embedded_cell_images()
        .expect("extract embedded cell images");
    assert_eq!(images.len(), 2);

    let key_a1 = ("xl/worksheets/sheet1.xml".to_string(), CellRef::new(0, 0));
    let key_a2 = ("xl/worksheets/sheet1.xml".to_string(), CellRef::new(1, 0));

    let a1 = images.get(&key_a1).expect("expected A1 embedded image");
    assert_eq!(a1.image_part, "xl/media/image1.png");
    assert_eq!(a1.image_bytes, img1);
    assert_eq!(a1.calc_origin, 0);
    assert!(a1.alt_text.is_none());

    let a2 = images.get(&key_a2).expect("expected A2 embedded image");
    assert_eq!(a2.image_part, "xl/media/image2.png");
    assert_eq!(a2.image_bytes, img2);
    assert_eq!(a2.calc_origin, 0);
    assert!(a2.alt_text.is_none());
}

#[test]
fn extracts_alt_text_from_embedded_image() {
    let bytes = build_workbook_with_embedded_image(Some("hello alt text"), false);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read xlsx");

    let images = pkg
        .extract_embedded_cell_images()
        .expect("extract embedded cell images");
    assert_eq!(images.len(), 1);

    let key = ("xl/worksheets/sheet1.xml".to_string(), CellRef::new(0, 0));
    let cell_img = images.get(&key).expect("expected A1 image");
    assert_eq!(cell_img.alt_text.as_deref(), Some("hello alt text"));
    assert_eq!(cell_img.calc_origin, 5);
}

#[test]
fn dynamic_array_metadata_does_not_break_embedded_image_extraction() {
    let bytes = build_workbook_with_embedded_image(None, true);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read xlsx");

    let metadata_xml = std::str::from_utf8(pkg.part("xl/metadata.xml").unwrap()).unwrap();
    assert!(
        metadata_xml.contains("XLDAPR"),
        "expected dynamic array metadata type in xl/metadata.xml"
    );
    assert!(
        metadata_xml.contains("XLRICHVALUE"),
        "expected rich value metadata type in xl/metadata.xml"
    );

    let images = pkg
        .extract_embedded_cell_images()
        .expect("extract embedded cell images");
    let addr = ("xl/worksheets/sheet1.xml".to_string(), CellRef::new(0, 0));
    assert!(
        images.contains_key(&addr),
        "expected embedded image mapping even with dynamic array metadata present"
    );
}
