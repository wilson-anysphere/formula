use formula_xlsx::{load_from_bytes, XlsxPackage};
use formula_model::{CellRef, CellValue};
use std::collections::BTreeMap;
use std::io::{Cursor, Read, Write};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

fn build_two_sheet_drawing_workbook() -> Vec<u8> {
    let base = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");
    let cursor = Cursor::new(base);
    let mut archive = ZipArchive::new(cursor).expect("open base fixture zip");

    let mut parts = BTreeMap::<String, Vec<u8>>::new();
    for idx in 0..archive.len() {
        let mut file = archive.by_index(idx).expect("zip entry");
        if file.is_dir() {
            continue;
        }
        let name = file.name().trim_start_matches('/').to_string();
        let mut buf = Vec::new();
        file.read_to_end(&mut buf).expect("read zip entry");
        parts.insert(name, buf);
    }

    // Duplicate the worksheet + its `.rels` but point at a new drawing part.
    let sheet1 = parts
        .get("xl/worksheets/sheet1.xml")
        .expect("base sheet1.xml")
        .clone();
    parts.insert("xl/worksheets/sheet2.xml".to_string(), sheet1);

    let sheet1_rels = String::from_utf8(
        parts
            .get("xl/worksheets/_rels/sheet1.xml.rels")
            .expect("base sheet1 rels")
            .clone(),
    )
    .expect("utf8 sheet rels");
    let sheet2_rels = sheet1_rels.replace("drawing1.xml", "drawing2.xml");
    parts.insert(
        "xl/worksheets/_rels/sheet2.xml.rels".to_string(),
        sheet2_rels.into_bytes(),
    );

    // Duplicate drawing1 -> drawing2 (and its `.rels`).
    let drawing1 = parts
        .get("xl/drawings/drawing1.xml")
        .expect("base drawing1.xml")
        .clone();
    let drawing2 = String::from_utf8(drawing1.clone()).expect("drawing xml utf-8");
    let drawing2 = drawing2.replace("Picture 1", "Picture 2");
    parts.insert("xl/drawings/drawing2.xml".to_string(), drawing2.into_bytes());
    let drawing1_rels = parts
        .get("xl/drawings/_rels/drawing1.xml.rels")
        .expect("base drawing1 rels")
        .clone();
    parts.insert("xl/drawings/_rels/drawing2.xml.rels".to_string(), drawing1_rels);

    // Replace workbook.xml to reference both worksheets.
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>
"#;
    parts.insert("xl/workbook.xml".to_string(), workbook_xml.to_vec());

    // Replace workbook.xml.rels with a consistent relationship table.
    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"#;
    parts.insert("xl/_rels/workbook.xml.rels".to_string(), workbook_rels.to_vec());

    // Update `[Content_Types].xml` so the new parts have overrides.
    let content_types = String::from_utf8(
        parts
            .get("[Content_Types].xml")
            .expect("base content types")
            .clone(),
    )
    .expect("utf8 [Content_Types].xml");
    let content_types = content_types.replace(
        r#"<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
        r#"<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
    );
    let content_types = content_types.replace(
        r#"<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>"#,
        r#"<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>"#,
    );
    parts.insert("[Content_Types].xml".to_string(), content_types.into_bytes());

    // Repack into a new XLSX zip.
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);
    for (name, bytes) in parts {
        zip.start_file(name, options).expect("zip start_file");
        zip.write_all(&bytes).expect("zip write_all");
    }
    zip.finish().expect("zip finish").into_inner()
}

#[test]
fn noop_save_preserves_drawing_parts_and_media_bytes() {
    let original_bytes = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");

    let doc = load_from_bytes(original_bytes).expect("load fixture");
    let saved = doc.save_to_vec().expect("save");

    let before = XlsxPackage::from_bytes(original_bytes).expect("read original pkg");
    let after = XlsxPackage::from_bytes(&saved).expect("read saved pkg");

    assert_eq!(
        before.part("xl/drawings/drawing1.xml").unwrap(),
        after.part("xl/drawings/drawing1.xml").unwrap(),
        "drawing XML should be preserved byte-for-byte on no-op save"
    );
    assert_eq!(
        before.part("xl/drawings/_rels/drawing1.xml.rels").unwrap(),
        after.part("xl/drawings/_rels/drawing1.xml.rels").unwrap(),
        "drawing relationship XML should be preserved byte-for-byte on no-op save"
    );
    assert_eq!(
        before.part("xl/media/image1.png").unwrap(),
        after.part("xl/media/image1.png").unwrap(),
        "image media bytes should be preserved byte-for-byte on no-op save"
    );
    assert_eq!(
        before
            .part("xl/worksheets/_rels/sheet1.xml.rels")
            .unwrap(),
        after.part("xl/worksheets/_rels/sheet1.xml.rels").unwrap(),
        "worksheet relationship part should be preserved byte-for-byte on no-op save"
    );
}

#[test]
fn editing_cells_does_not_rewrite_drawing_parts_or_media() {
    let original_bytes = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");

    let mut doc = load_from_bytes(original_bytes).expect("load fixture");
    let sheet_id = doc.workbook.sheets[0].id;
    doc.workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .set_value(CellRef::from_a1("A1").expect("valid A1"), CellValue::Number(123.0));
    let saved = doc.save_to_vec().expect("save");

    let before = XlsxPackage::from_bytes(original_bytes).expect("read original pkg");
    let after = XlsxPackage::from_bytes(&saved).expect("read saved pkg");

    assert_eq!(
        before.part("xl/drawings/drawing1.xml").unwrap(),
        after.part("xl/drawings/drawing1.xml").unwrap(),
        "drawing XML should be preserved byte-for-byte when editing unrelated cells"
    );
    assert_eq!(
        before.part("xl/drawings/_rels/drawing1.xml.rels").unwrap(),
        after.part("xl/drawings/_rels/drawing1.xml.rels").unwrap(),
        "drawing relationship XML should be preserved byte-for-byte when editing unrelated cells"
    );
    assert_eq!(
        before.part("xl/media/image1.png").unwrap(),
        after.part("xl/media/image1.png").unwrap(),
        "image media bytes should be preserved byte-for-byte when editing unrelated cells"
    );
    assert_eq!(
        before
            .part("xl/worksheets/_rels/sheet1.xml.rels")
            .unwrap(),
        after.part("xl/worksheets/_rels/sheet1.xml.rels").unwrap(),
        "worksheet relationship part should be preserved byte-for-byte when editing unrelated cells"
    );
}

#[test]
fn noop_save_preserves_drawing_parts_even_without_drawings_snapshot() {
    let original_bytes = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");

    let mut doc = load_from_bytes(original_bytes).expect("load fixture");
    let sheet_id = doc.workbook.sheets[0].id;
    // Simulate a document that has worksheet drawings populated but is missing the baseline
    // `drawings_snapshot` (e.g. metadata was dropped while persisting/restoring state).
    doc.xlsx_meta_mut().drawings_snapshot.remove(&sheet_id);

    let saved = doc.save_to_vec().expect("save");

    let before = XlsxPackage::from_bytes(original_bytes).expect("read original pkg");
    let after = XlsxPackage::from_bytes(&saved).expect("read saved pkg");

    assert_eq!(
        before.part("xl/drawings/drawing1.xml").unwrap(),
        after.part("xl/drawings/drawing1.xml").unwrap(),
        "drawing XML should be preserved byte-for-byte on no-op save even without a snapshot"
    );
    assert_eq!(
        before.part("xl/drawings/_rels/drawing1.xml.rels").unwrap(),
        after.part("xl/drawings/_rels/drawing1.xml.rels").unwrap(),
        "drawing relationship XML should be preserved byte-for-byte on no-op save even without a snapshot"
    );
    assert_eq!(
        before.part("xl/media/image1.png").unwrap(),
        after.part("xl/media/image1.png").unwrap(),
        "image media bytes should be preserved byte-for-byte on no-op save even without a snapshot"
    );
    assert_eq!(
        before
            .part("xl/worksheets/_rels/sheet1.xml.rels")
            .unwrap(),
        after.part("xl/worksheets/_rels/sheet1.xml.rels").unwrap(),
        "worksheet relationship part should be preserved byte-for-byte on no-op save even without a snapshot"
    );
}

#[test]
fn multi_sheet_noop_save_preserves_all_drawing_parts() {
    let original_bytes = build_two_sheet_drawing_workbook();

    let doc = load_from_bytes(&original_bytes).expect("load synthetic workbook");
    let saved = doc.save_to_vec().expect("save");

    let before = XlsxPackage::from_bytes(&original_bytes).expect("read original pkg");
    let after = XlsxPackage::from_bytes(&saved).expect("read saved pkg");

    for part in [
        "xl/drawings/drawing1.xml",
        "xl/drawings/drawing2.xml",
        "xl/drawings/_rels/drawing1.xml.rels",
        "xl/drawings/_rels/drawing2.xml.rels",
        "xl/worksheets/_rels/sheet1.xml.rels",
        "xl/worksheets/_rels/sheet2.xml.rels",
        "xl/media/image1.png",
    ] {
        assert_eq!(
            before.part(part).unwrap(),
            after.part(part).unwrap(),
            "{part} should be preserved byte-for-byte on no-op save"
        );
    }
}

#[test]
fn multi_sheet_cell_edit_preserves_all_drawing_parts() {
    let original_bytes = build_two_sheet_drawing_workbook();

    let mut doc = load_from_bytes(&original_bytes).expect("load synthetic workbook");
    let sheet1_id = doc.workbook.sheets[0].id;
    doc.workbook
        .sheet_mut(sheet1_id)
        .expect("sheet1 exists")
        .set_value(CellRef::from_a1("B2").expect("valid B2"), CellValue::Number(42.0));
    let saved = doc.save_to_vec().expect("save");

    let before = XlsxPackage::from_bytes(&original_bytes).expect("read original pkg");
    let after = XlsxPackage::from_bytes(&saved).expect("read saved pkg");

    for part in [
        "xl/drawings/drawing1.xml",
        "xl/drawings/drawing2.xml",
        "xl/drawings/_rels/drawing1.xml.rels",
        "xl/drawings/_rels/drawing2.xml.rels",
        "xl/worksheets/_rels/sheet1.xml.rels",
        "xl/worksheets/_rels/sheet2.xml.rels",
        "xl/media/image1.png",
    ] {
        assert_eq!(
            before.part(part).unwrap(),
            after.part(part).unwrap(),
            "{part} should be preserved byte-for-byte when editing unrelated cells"
        );
    }
}
